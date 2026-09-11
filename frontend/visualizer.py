from PIL import Image, ImageDraw, ImageFont

def draw_visual_highlights(img, unified_blocks, page_mappings, draw_gray_layout=True) -> Image.Image:
    """
    Draws blue boxes on anchors, green boxes on auto-detected values, and red boxes on manual crop values.
    If draw_gray_layout is False, it skips background layout boxes to keep the original page clean.
    """
    draw_img = img.copy()
    draw = ImageDraw.Draw(draw_img, "RGBA")
    
    # Load font
    try:
        font = ImageFont.truetype("tahoma.ttf", 16)
    except IOError:
        font = ImageFont.load_default()
        
    # Draw all detected blocks in thin light gray (background layout) if requested
    if draw_gray_layout:
        for block in unified_blocks:
            x1, y1, x2, y2 = block['bbox']
            draw.rectangle([x1, y1, x2, y2], outline=(148, 163, 184, 80), fill=(148, 163, 184, 15), width=1)
            
    # Draw matched anchors in blue and values in green/red
    for m in page_mappings:
        anchor_rect = m.get('anchor_rect')
        if anchor_rect:
            ax1 = anchor_rect['x']
            ay1 = anchor_rect['y']
            ax2 = ax1 + anchor_rect['width']
            ay2 = ay1 + anchor_rect['height']
            draw.rectangle([ax1, ay1, ax2, ay2], outline=(99, 102, 241, 255), fill=(99, 102, 241, 35), width=2)
            
        value_rect = m.get('value_rect')
        if value_rect:
            vx1 = value_rect['x']
            vy1 = value_rect['y']
            vx2 = vx1 + value_rect['width']
            vy2 = vy1 + value_rect['height']
            
            # Use red color for manually defined/adjusted crops, green for auto-detected crops
            is_manual = m.get('is_manual_crop', False)
            border_color = (239, 68, 68, 255) if is_manual else (34, 197, 94, 255)
            fill_color = (239, 68, 68, 35) if is_manual else (34, 197, 94, 35)
            
            draw.rectangle([vx1, vy1, vx2, vy2], outline=border_color, fill=fill_color, width=2)
            
            # Label badge (dynamic width based on text length)
            label_text = m.get('cell') or m.get('header') or "Value"
            badge_w = len(label_text) * 9 + 10
            draw.rectangle([vx1, max(0, vy1 - 20), vx1 + badge_w, vy1], fill=border_color)
            draw.text((vx1 + 5, max(0, vy1 - 18)), str(label_text), fill=(255, 255, 255, 255), font=font)
            
    return draw_img
