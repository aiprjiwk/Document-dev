import base64
import io
from PIL import Image
import json

def render_fabric_canvas(image: Image.Image, mapping_df) -> str:
    """
    Generates a full interactive Fabric.js HTML5 canvas component string.
    Renders ALL red bounding boxes simultaneously as draggable/resizable Fabric.js objects.
    """
    w, h = image.size
    
    # Scale image for display (e.g. max width 800px)
    max_disp_w = 800
    scale = max_disp_w / float(w) if w > max_disp_w else 1.0
    disp_w = int(w * scale)
    disp_h = int(h * scale)
    
    # Encode image to base64
    buf = io.BytesIO()
    image.save(buf, format="JPEG", quality=85)
    img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    
    # Build Fabric.js objects array
    objects = []
    for idx, row in mapping_df.iterrows():
        try:
            name = str(row["Field Name"])
            cell = str(row["Excel Cell"])
            top_pct = float(row["Top (%)"])
            left_pct = float(row["Left (%)"])
            h_pct = float(row["Height (%)"])
            w_pct = float(row["Width (%)"])
            
            box_left = (left_pct / 100.0) * disp_w
            box_top = (top_pct / 100.0) * disp_h
            box_width = (w_pct / 100.0) * disp_w
            box_height = (h_pct / 100.0) * disp_h
            
            objects.append({
                "field": name,
                "cell": cell,
                "idx": int(idx),
                "left": box_left,
                "top": box_top,
                "width": box_width,
                "height": box_height
            })
        except Exception:
            pass
            
    objects_json = json.dumps(objects)
    
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <script src="https://cdnjs.cloudflare.com/ajax/libs/fabric.js/5.3.1/fabric.min.js"></script>
      <style>
        body {{ margin: 0; padding: 0; font-family: sans-serif; background-color: #f8fafc; }}
        #canvas-container {{ position: relative; border: 2px solid #cbd5e1; border-radius: 8px; overflow: hidden; display: inline-block; }}
        .info-panel {{ margin-top: 10px; padding: 10px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 13px; }}
        .badge {{ background: #ef4444; color: white; padding: 2px 6px; border-radius: 4px; font-size: 11px; margin-right: 5px; }}
      </style>
    </head>
    <body>
      <div id="canvas-container">
        <canvas id="c" width="{disp_w}" height="{disp_h}"></canvas>
      </div>
      
      <div class="info-panel">
        <strong>🎯 Direct Mouse Multi-Box Adjuster (Fabric.js)</strong><br/>
        <span>Click and drag ANY red box directly on the document image to move or resize it freely with your mouse.</span>
        <div id="status-out" style="margin-top: 5px; color: #2563eb; font-weight: bold;">Selected: None</div>
      </div>

      <script>
        const canvas = new fabric.Canvas('c', {{
          selection: false
        }});

        const scale = {scale};
        const dispW = {disp_w};
        const dispH = {disp_h};

        // Set background image
        fabric.Image.fromURL("data:image/jpeg;base64,{img_b64}", function(img) {{
          img.set({{
            scaleX: scale,
            scaleY: scale,
            selectable: false,
            evented: false
          }});
          canvas.setBackgroundImage(img, canvas.renderAll.bind(canvas));
        }});

        const initialBoxes = {objects_json};

        initialBoxes.forEach(box => {{
          const rect = new fabric.Rect({{
            left: box.left,
            top: box.top,
            width: box.width,
            height: box.height,
            fill: 'rgba(239, 68, 68, 0.15)',
            stroke: '#ef4444',
            strokeWidth: 2,
            cornerColor: '#ef4444',
            cornerSize: 8,
            transparentCorners: false,
            fieldIndex: box.idx,
            fieldName: box.field,
            excelCell: box.cell
          }});

          const text = new fabric.Text(box.field + " (" + box.cell + ")", {{
            fontSize: 12,
            fill: '#ef4444',
            backgroundColor: 'rgba(255, 255, 255, 0.85)',
            left: box.left,
            top: Math.max(0, box.top - 16),
            selectable: false,
            evented: false
          }});

          rect.on('moving', function() {{
            text.set({{ left: rect.left, top: Math.max(0, rect.top - 16) }});
            updateStatus(rect);
          }});

          rect.on('scaling', function() {{
            text.set({{ left: rect.left, top: Math.max(0, rect.top - 16) }});
            updateStatus(rect);
          }});

          rect.on('selected', function() {{
            updateStatus(rect);
          }});

          canvas.add(rect);
          canvas.add(text);
        }});

        function updateStatus(rect) {{
          const leftPct = ((rect.left / dispW) * 100).toFixed(1);
          const topPct = ((rect.top / dispH) * 100).toFixed(1);
          const widthPct = (((rect.width * rect.scaleX) / dispW) * 100).toFixed(1);
          const heightPct = (((rect.height * rect.scaleY) / dispH) * 100).toFixed(1);
          
          document.getElementById('status-out').innerHTML = 
            "<span class='badge'>" + rect.fieldName + "</span> " +
            "Left: " + leftPct + "%, Top: " + topPct + "%, Width: " + widthPct + "%, Height: " + heightPct + "%";
        }}
      </script>
    </body>
    </html>
    """
    return html_code
