import os
import re
import html
import base64
import logging
from bs4 import BeautifulSoup

logger = logging.getLogger("FaultMaintenanceService")

def list_html_files(src_folder: str = "Fault assistance") -> list:
    """List all HTML fault template files, sorted by code."""
    files_list = []
    if not os.path.exists(src_folder):
        return []
    
    for f in os.listdir(src_folder):
        if f.endswith(".html") and f != "index.html" and f != "All fault messages.html":
            # Extract fault code from file name
            match = re.match(r'^([A-Za-z0-9]+)_(.*)\.html$', f)
            if match:
                code = match.group(1)
                text = match.group(2)
                files_list.append({
                    "filename": f,
                    "code": code,
                    "name": text,
                    "display": f"{code} - {text}"
                })
    
    # Sort by code
    files_list.sort(key=lambda x: x["code"])
    return files_list

def save_markdown_formatting(text: str) -> str:
    """Escape text and map markdown styles (**bold**, *italic*, {code}) to ST4 HTML tags."""
    esc = html.escape(text)
    # Convert code braces: {something} -> {<code>something</code>}
    esc = re.sub(r'\{([^{}]+)\}', r'{<code>\1</code>}', esc)
    # Convert bold: **something** -> <span class="Bold">something</span>
    esc = re.sub(r'\*\*([^*]+)\*\*', r'<span class="Bold">\1</span>', esc)
    # Convert italic: *something* -> <em>something</em>
    esc = re.sub(r'\*([^*]+)\*', r'<em>\1</em>', esc)
    return esc

def restore_markdown_formatting(html_str: str) -> str:
    """Map ST4 HTML tags back to markdown syntax (**bold**, *italic*, {code})."""
    # Replace code tag: {<code>something</code>} -> {something}
    html_str = re.sub(r'\{<code>([^<]+)</code>\}', r'{\1}', html_str)
    # Also check code tag without braces: <code>something</code> -> {something}
    html_str = re.sub(r'<code>([^<]+)</code>', r'{\1}', html_str)
    # Replace bold tag: <span class="Bold">something</span> -> **something**
    html_str = re.sub(r'<span class="Bold">([^<]+)</span>', r'**\1**', html_str)
    # Replace em tag: <em>something</em> -> *something*
    html_str = re.sub(r'<em>([^<]+)</em>', r'*\1*', html_str)
    return html_str

def html_to_markdown_cell(html_str: str) -> str:
    """Helper to convert cell HTML back to editable text with bullet points, preserving formatting."""
    # Recursively unescape html_str to heal any nested escapes from previous bugs
    prev_str = ""
    while html_str != prev_str:
        prev_str = html_str
        html_str = html.unescape(html_str)
        
    cell_soup = BeautifulSoup(html_str, "html.parser")
    
    # 1. Process unordered lists
    ul_list = cell_soup.find_all("ul")
    for ul in ul_list:
        items = []
        for dd in ul.find_all("dd", class_="UnorderedList"):
            # Decompose bullet point dt if present
            dl = dd.parent
            if dl:
                dt = dl.find("dt")
                if dt:
                    dt.decompose()
            # Process formatting inside the list item dd
            for tag in dd.find_all("span", class_="Bold"):
                tag.replace_with(f"**{tag.get_text()}**")
            for tag in dd.find_all("em"):
                tag.replace_with(f"*{tag.get_text()}*")
            for tag in dd.find_all("code"):
                tag.replace_with(f"{{{tag.get_text()}}}")
            items.append("- " + dd.get_text().strip())
        ul.replace_with("\n".join(items) + "\n")
        
    # 2. Process formatting in paragraphs outside lists
    for tag in cell_soup.find_all("span", class_="Bold"):
        tag.replace_with(f"**{tag.get_text()}**")
    for tag in cell_soup.find_all("em"):
        tag.replace_with(f"*{tag.get_text()}*")
    for tag in cell_soup.find_all("code"):
        tag.replace_with(f"{{{tag.get_text()}}}")
        
    cell_text = cell_soup.get_text().strip()
    cell_text = re.sub(r'\{\{([^{}]+)\}\}', r'{\1}', cell_text)
    return cell_text

def markdown_to_html_cell(text: str) -> str:
    """Helper to convert editable text back to ST4 cell HTML structure, preserving formatting."""
    lines = text.strip().split("\n")
    html_parts = []
    in_list = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("- ") or line.startswith("* "):
            if not in_list:
                html_parts.append('<ul class="list" style="margin-bottom:20px;">')
                in_list = True
            content = line[2:].strip()
            # Convert formatting markdown to HTML
            formatted_content = save_markdown_formatting(content)
            html_parts.append(
                f'<div><dl class="UnorderedList"><dt class="UnorderedList">▪</dt>'
                f'<dd class="UnorderedList">{formatted_content}</dd></dl></div>'
            )
        else:
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            # Convert formatting markdown to HTML
            formatted_line = save_markdown_formatting(line)
            html_parts.append(f'<div class="Paragraph">{formatted_line}</div>')
            
    if in_list:
        html_parts.append('</ul>')
        
    # Wrap in TableCellLeft wrapper
    wrapped_html = f'<div class="TableCellLeft">{"".join(html_parts)}</div>'
    return wrapped_html

def parse_st4_to_blocks(html_content: str) -> tuple:
    """Parse ST4 HTML to extract Heading and dynamic block elements in order. 
    Uses leaf textmodules to prevent nested wrapper blocks matching children out of order."""
    try:
        soup = BeautifulSoup(html_content, "html.parser")
        content_div = soup.find("div", class_="content")
        if not content_div:
            return "", []
            
        # Get Heading
        heading = ""
        heading_div = soup.find("div", class_="Heading")
        if heading_div:
            code_tag = heading_div.find("code")
            heading = code_tag.get_text().strip() if code_tag else heading_div.get_text().strip()
            
        blocks = []
        textmodules = content_div.find_all("div", class_="textmodule")
        # Filter to only leaf textmodules (do not contain any other textmodule tags)
        leaf_textmodules = [tm for tm in textmodules if not tm.find("div", class_="textmodule")]
        
        for tm in leaf_textmodules:
            # 1. Check if Safety block is present
            safety_text_div = tm.find("div", class_="safetytext")
            if safety_text_div:
                classes = safety_text_div.get("class", [])
                w_type = "caution"
                for c in ["caution", "warning", "danger", "notice"]:
                    if c in classes:
                        w_type = c
                        break
                        
                word_div = safety_text_div.find(class_="SafetyWord")
                word = word_div.get_text().strip() if word_div else w_type.upper()
                
                cause_div = safety_text_div.find(class_="SafetyCause")
                cause = cause_div.get_text().strip() if cause_div else ""
                
                consequence_divs = safety_text_div.find_all(class_="SafetyConsequence")
                consequences = [c.get_text().strip() for c in consequence_divs]
                
                # Extract specific safety warning icon filename
                icon_div = tm.find(class_="safetyicon")
                icon_img = icon_div.find("img") if icon_div else None
                icon_name = "general_general"
                if icon_img:
                    src = icon_img.get("src", "")
                    basename = os.path.basename(src)
                    icon_name = os.path.splitext(basename)[0]
                
                blocks.append({
                    "type": "Safety",
                    "warn_type": w_type,
                    "word": word,
                    "cause": cause,
                    "consequences": consequences,
                    "icon": icon_name
                })
                continue

            # 2. Check if Cause/Elimination table is present
            table = tm.find("table")
            if table:
                tr_list = table.find_all("tr")
                if tr_list and ("cause" in tr_list[0].get_text().lower() or "elimination" in tr_list[0].get_text().lower()):
                    rows = []
                    for tr in tr_list[1:]:
                        tds = tr.find_all("td")
                        if len(tds) >= 2:
                            cause_text = html_to_markdown_cell(str(tds[0]))
                            elim_text = html_to_markdown_cell(str(tds[1]))
                            rows.append({"cause": cause_text, "elimination": elim_text})
                    blocks.append({"type": "Table", "rows": rows})
                    continue
                    
            # 2b. Check if Image is present
            anchor_image_div = tm.find("div", class_="anchor_image")
            if anchor_image_div:
                img = anchor_image_div.find("img")
                if img:
                    src = img.get("src", "")
                    blocks.append({"type": "Image", "src": src})
                    continue
                    
            # 3. Check for SubHeading
            subheading = tm.find("div", class_="SubHeading")
            if subheading:
                blocks.append({"type": "SubHeading", "text": subheading.get_text().strip()})
                continue
                
            # 4. Check for Paragraphs
            paragraphs = tm.find_all("div", class_="Paragraph")
            if paragraphs:
                for p in paragraphs:
                    p_icon = ""
                    dl = p.find("dl")
                    if dl:
                        img = dl.find("img")
                        if img:
                            alt = img.get("alt", "").lower()
                            src = img.get("src", "").lower()
                            if "listiconaction" in alt or "listiconaction" in src:
                                p_icon = "action"
                            elif "listicontouch" in alt or "listicontouch" in src:
                                p_icon = "touch"
                        dl.decompose()
                    
                    for tag in p.find_all("span", class_="Bold"):
                        tag.replace_with(f"**{tag.get_text()}**")
                    for tag in p.find_all("em"):
                        tag.replace_with(f"*{tag.get_text()}*")
                    for tag in p.find_all("code"):
                        tag.replace_with(f"{{{tag.get_text()}}}")
                    
                    p_text = p.get_text().strip()
                    p_text = re.sub(r'\{\{([^{}]+)\}\}', r'{\1}', p_text)
                    blocks.append({"type": "Paragraph", "text": p_text, "icon": p_icon})
                continue
                
            # 5. Check for Lists
            ul = tm.find("ul", class_="list")
            if ul:
                items = []
                for dd in ul.find_all("dd", class_="UnorderedList"):
                    items.append(dd.get_text().strip())
                blocks.append({"type": "List", "items": items})
                continue
                
        return heading, blocks
    except Exception as e:
        logger.error(f"Error parsing ST4 blocks: {e}")
        return "", []

def save_html_blocks(file_path: str, heading: str, blocks: list) -> tuple:
    """Reconstruct ONLY the <div class='content'> element of ST4 HTML from blocks and save."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            html_content = f.read()
            
        soup = BeautifulSoup(html_content, "html.parser")
        content_div = soup.find("div", class_="content")
        if not content_div:
            return False, "Malformed HTML file: no <div class=\"content\"> found."
            
        # Reconstruct the inner HTML of content_div
        new_content_parts = [
            '<div class="anchor_clear"></div>\n',
            '<div class="anchor_margin"><div class="HeadingMargin"></div></div>\n',
            f'<div class="anchor_text"><div class="Heading"><code>{html.escape(heading)}</code><span class="NotReleased"></span></div></div>\n',
            '<div class="metalist"></div>\n'
        ]
        
        for block in blocks:
            b_type = block["type"]
            if b_type == "SubHeading":
                esc_txt = html.escape(block["text"])
                new_content_parts.append(
                    f'<div class="textmodule"><div class="anchor_text"><div class="SubHeading">{esc_txt}</div></div></div>\n'
                )
            elif b_type == "Paragraph":
                icon_v = block.get("icon", "")
                dl_prefix = ""
                if icon_v == "action":
                    dl_prefix = '<dl class="Gen18999aca82174ea99a898a0ba2ffb434"><dt class="Gen18999aca82174ea99a898a0ba2ffb434"><span class="Gen97bdbcaeda7d442d8b32ca4accd03b32"><a name=""><span style="display: none;"></span></a><img style="max-width: 100%;" title="Titel : ListIconAction &#xA;ID :" alt="ListIconAction" src="Images/ListIconAction.gif" /></span></dt><dd class="Gen18999aca82174ea99a898a0ba2ffb434"></dd></dl>'
                elif icon_v == "touch":
                    dl_prefix = '<dl class="Genb96b8d69637f4d9fbf94faafab8ca869"><dt class="Genb96b8d69637f4d9fbf94faafab8ca869"><span class="Genc92fde6808ff4a6baa01d3056719364d"><a name=""><span style="display: none;"></span></a><img style="max-width: 100%;" title="Titel : ListIconTouch&#xA;ID : 2530709643" alt="ListIconTouch" src="Images/ListIconTouch.gif" /></span></dt><dd class="Genb96b8d69637f4d9fbf94faafab8ca869"></dd></dl>'
                
                formatted_txt = save_markdown_formatting(block["text"])
                new_content_parts.append(
                    f'<div class="textmodule"><div class="anchor_text" style="margin-bottom:20px;"><div class="Paragraph">{dl_prefix}{formatted_txt}</div></div></div>\n'
                )
            elif b_type == "List":
                list_items = []
                for item in block["items"]:
                    esc_item = html.escape(item)
                    list_items.append(
                        f'<div class="anchor_text"><dl class="UnorderedList"><dt class="UnorderedList">▪</dt>'
                        f'<dd class="UnorderedList">{esc_item}</dd></dl></div>\n'
                    )
                new_content_parts.append(
                    f'<div class="textmodule"><ul class="list" style="margin-bottom:20px;">\n'
                    f'{"".join(list_items)}'
                    f'</ul></div>\n'
                )
            elif b_type == "Table":
                td_class = "ID0E1NCI"
                table_rows_html = [
                    f'<tr valign="top">'
                    f'<td class="{td_class}"><div style="margin-bottom: 0.1cm;margin-left: 0.1cm;margin-right: 0.1cm;margin-top: 0.1cm;overflow: hidden;min-width:25px;"><div class="TableCellCenter"><em>Cause</em></div></div></td>'
                    f'<td class="{td_class}"><div style="margin-bottom: 0.1cm;margin-left: 0.1cm;margin-right: 0.1cm;margin-top: 0.1cm;overflow: hidden;min-width:25px;"><div class="TableCellCenter"><em>Elimination</em></div></div></td>'
                    f'</tr>\n'
                ]
                for row in block["rows"]:
                    cause_html = markdown_to_html_cell(row["cause"])
                    elim_html = markdown_to_html_cell(row["elimination"])
                    table_rows_html.append(
                        f'<tr valign="top">'
                        f'<td class="{td_class}"><div style="margin-bottom: 0.1cm;margin-left: 0.1cm;margin-right: 0.1cm;margin-top: 0.1cm;overflow: hidden;min-width:25px;">{cause_html}</div></td>'
                        f'<td class="{td_class}"><div style="margin-bottom: 0.1cm;margin-left: 0.1cm;margin-right: 0.1cm;margin-top: 0.1cm;overflow: hidden;min-width:25px;">{elim_html}</div></td>'
                        f'</tr>\n'
                    )
                new_content_parts.append(
                    f'<div class="textmodule"><div class="table-container framed" style="margin-top:-1px;"><div><div style="width:100%;"><div align="left">\n'
                    f'<table style="max-width:100%;min-width:100%;width:100%;direction: ltr;width: 100%;" class="ID0ECECI" cellpadding="0" cellspacing="0" border="0">\n'
                    f'<colgroup><col style="width: 50.0%;"></col><col style="width: 49.9%;"></col></colgroup>\n'
                    f'<tbody>\n'
                    f'{"".join(table_rows_html)}'
                    f'</tbody>\n'
                    f'</table>\n'
                    f'</div></div></div></div></div>\n'
                )
            elif b_type == "Safety":
                w_type = block["warn_type"]
                word = block["word"]
                cause = block["cause"]
                consequences = block["consequences"]
                icon_name = block.get("icon", "general_general")
                
                img_src = f"safety/{icon_name}.png"
                    
                consequence_htmls = []
                for cons in consequences:
                    esc_cons = html.escape(cons)
                    consequence_htmls.append(
                        f'<div><div class="SafetyConsequence" style="padding-top:5px;">{esc_cons}</div></div>\n'
                    )
                
                esc_cause = html.escape(cause)
                esc_word = html.escape(word)
                
                new_content_parts.append(
                    f'<div class="textmodule">\n'
                    f'<div class="safetyicon {w_type}" style="margin-bottom:-100px;">'
                    f'<div class="anchor_text"><div class="SafetyIcon">'
                    f'<img style="height:60px;width:auto;" src="{img_src}" alt="{w_type}" />'
                    f'</div></div></div>'
                    f'<div class="safetytext {w_type}" style="margin-top:20px;">'
                    f'<div class="safetyheader {w_type}" style="padding-bottom:10px;padding-left:15px;padding-right:15px;padding-top:5px;">'
                    + (f'<div><div class="SafetyWordIcon"><img style="max-width: 100%;" alt="SmallWarningYellow" src="Images/SmallWarningYellow.gif" /></div></div>' if w_type != "notice" else "") +
                    f'<div><div class="SafetyWord">{esc_word}</div></div>'
                    f'</div>'
                    f'<div class="safetycontent" style="padding-bottom:5px;padding-left:15px;padding-right:15px;">'
                    f'<div><div class="SafetyCause" style="padding-top:5px;">{esc_cause}</div></div>'
                    f'{"".join(consequence_htmls)}'
                    f'</div></div></div>\n'
                )
            elif b_type == "Image":
                src = block.get("src", "")
                new_content_parts.append(
                    f'<div class="textmodule"><div class="anchor_image" style="margin-bottom:20px;"><div class="image framed"><img src="{src}" /></div></div></div>\n'
                )
                
        # Replace child nodes in content_div
        new_inner_soup = BeautifulSoup("".join(new_content_parts), "html.parser")
        content_div.clear()
        for child in new_inner_soup.contents:
            content_div.append(child)
            
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(str(soup))
            
        return True, "File saved successfully."
    except Exception as e:
        return False, f"Save Error: {str(e)}"

def read_and_parse_html(file_path: str) -> dict:
    """Parse ST4 HTML (Fallback legacy parser)."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            html_content = f.read()
        return {
            "success": True,
            "raw_html": html_content
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

def save_raw_html_file(file_path: str, raw_html: str) -> tuple:
    """Directly save/overwrite raw HTML code."""
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(raw_html)
        return True, "Raw HTML file saved successfully."
    except Exception as e:
        return False, f"Save Error: {str(e)}"

def create_new_fault_file(src_folder: str, new_code: str, new_text: str) -> tuple:
    """Duplicate a base template file to create a new fault code entry."""
    try:
        new_filename = f"{new_code.strip()}_{new_text.strip()}.html"
        new_file_path = os.path.join(src_folder, new_filename)
        
        if os.path.exists(new_file_path):
            return False, f"File already exists: {new_filename}"
            
        # Search for a base template file (contains "@1s@") in the directory
        template_file = None
        for f in os.listdir(src_folder):
            if "@1s@" in f and f.endswith(".html"):
                template_file = os.path.join(src_folder, f)
                break
                
        # If no "@1s@" template found, fallback to any valid html file
        if not template_file:
            for f in os.listdir(src_folder):
                if f.endswith(".html") and f != "index.html" and f != "All fault messages.html":
                    template_file = os.path.join(src_folder, f)
                    break
                    
        if not template_file:
            return False, "No template base file found in Fault assistance folder to duplicate."
            
        with open(template_file, "r", encoding="utf-8", errors="ignore") as f:
            template_content = f.read()
            
        # Create default header based on new_text
        soup = BeautifulSoup(template_content, "html.parser")
        
        # Update title & heading
        title_tag = soup.find("title")
        if title_tag:
            title_tag.string = f"Project - {new_text}"
            
        heading_div = soup.find("div", class_="Heading")
        if heading_div:
            code_tag = heading_div.find("code")
            if code_tag:
                code_tag.string = new_text
            else:
                heading_div.string = new_text
                
        with open(new_file_path, "w", encoding="utf-8") as f:
            f.write(str(soup))
            
        return True, f"New manual created successfully: {new_filename}"
    except Exception as e:
        return False, f"Creation Error: {str(e)}"

def delete_fault_file(src_folder: str, filename: str) -> tuple:
    """Delete a manual template file from folder."""
    try:
        file_path = os.path.join(src_folder, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            return True, f"Deleted file: {filename}"
        return False, f"File not found: {filename}"
    except Exception as e:
        return False, f"Delete Error: {str(e)}"

def get_html_preview_with_styles(file_path: str, src_folder: str = "Fault assistance") -> str:
    """Read HTML file and inject local CSS sheets and base64 images for full preview rendering."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            html_content = f.read()
            
        soup = BeautifulSoup(html_content, "html.parser")
        
        # 1. Inline all local CSS styles
        links = soup.find_all("link", rel="stylesheet")
        for link in links:
            href = link.get("href", "")
            css_path = os.path.join(src_folder, href)
            if os.path.exists(css_path):
                try:
                    with open(css_path, "r", encoding="utf-8", errors="ignore") as css_f:
                        css_content = css_f.read()
                    style_tag = soup.new_tag("style")
                    style_tag.string = css_content
                    link.replace_with(style_tag)
                except Exception as css_err:
                    logger.warning(f"Error reading css file {href}: {css_err}")
                    
        # 2. Base64 encode and inline all local images
        imgs = soup.find_all("img")
        for img in imgs:
            src = img.get("src", "")
            img_path = os.path.join(src_folder, src)
            if not os.path.exists(img_path):
                if os.path.exists(src):
                    img_path = src
                else:
                    parent_src = os.path.join(os.path.dirname(src_folder), src)
                    if os.path.exists(parent_src):
                        img_path = parent_src
            if os.path.exists(img_path):
                try:
                    ext = os.path.splitext(src)[1].replace(".", "").lower()
                    mime = "image/png" if ext == "png" else ("image/gif" if ext == "gif" else "image/jpeg")
                    with open(img_path, "rb") as img_f:
                        b64_data = base64.b64encode(img_f.read()).decode("utf-8")
                    img["src"] = f"data:{mime};base64,{b64_data}"
                except Exception as img_err:
                    logger.warning(f"Error base64 encoding img {src}: {img_err}")
                    
        return str(soup)
    except Exception as e:
        return f"<h3>Preview Generation Error</h3><p>{str(e)}</p>"
