"""
MCP Server for CATIA V5
Provides tools to interact with CATIA V5 through COM automation
"""

from fastmcp import FastMCP
from fastmcp.tools import Tool
try:
    import win32com.client
    import pythoncom
    CATIA_AVAILABLE = True
except ImportError:
    CATIA_AVAILABLE = False

# Initialize FastMCP server
app = FastMCP("catia_v5_server")

def get_catia_application():
    """Get or create CATIA application instance"""
    if not CATIA_AVAILABLE:
        raise Exception("pywin32 is required for CATIA automation. Install with: pip install pywin32")
    
    try:
        pythoncom.CoInitialize()
        catia = win32com.client.Dispatch("CATIA.Application")
        return catia
    except Exception as e:
        raise Exception(f"Failed to connect to CATIA V5: {str(e)}")

def release_com_object(obj):
    """Properly release COM object"""
    if obj:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


@app.tool("get_catia_info")
def get_catia_info():
    """Get CATIA application information"""
    catia = get_catia_application()
    info = {
        "version": catia.SystemConfiguration.Version,
        "visible": catia.Visible,
        "caption": catia.Caption,
        "full_name": catia.FullName,
    }
    return info


@app.tool("list_documents")
def list_documents():
    """List all open documents"""
    catia = get_catia_application()
    documents = catia.Documents
    doc_list = []
    for i in range(1, documents.Count + 1):
        doc = documents.Item(i)
        doc_list.append({
            "name": doc.Name,
            "full_name": doc.FullName,
            "type": doc.Name.split(".")[-1],
        })
    return doc_list


@app.tool("get_active_document")
def get_active_document():
    """Get active document information"""
    catia = get_catia_application()
    if catia.Documents.Count == 0:
        return "No documents are currently open"
    
    active_doc = catia.ActiveDocument
    info = {
        "name": active_doc.Name,
        "full_name": active_doc.FullName,
        "path": active_doc.Path if hasattr(active_doc, 'Path') else "Not saved",
        "saved": active_doc.Saved,
    }
    return info

@app.tool("create_part")
def create_part(name: str):
    """Create a new part document"""
    catia = get_catia_application()
    documents = catia.Documents
    part_doc = documents.Add("Part")
    part_doc.Part.PartDocument.Part.set_Name(name)
    return f"Created new Part document: {name}"
        
@app.tool("create_product")
def create_product(name: str):
    """Create a new product document"""
    catia = get_catia_application()
    documents = catia.Documents
    product_doc = documents.Add("Product")
    product_doc.Product.set_Name(name)
    return f"Created new Product document: {name}"

@app.tool("create_drawing")
def create_drawing(name: str):
    """Create a new drawing document"""
    catia = get_catia_application()
    documents = catia.Documents
    drawing_doc = documents.Add("Drawing")
    return f"Created new Drawing document: {name}"

@app.tool("open_document")
def open_document(file_path: str):
    """Open a document from file"""
    catia = get_catia_application()
    doc = catia.Documents.Open(file_path)
    return f"Opened document: {doc.Name}"

@app.tool("save_document")
def save_document(file_path: str = None):
    """Save the active document"""
    catia = get_catia_application()
    active_doc = catia.ActiveDocument
    if file_path:
        active_doc.SaveAs(file_path)
        return f"Document saved to: {file_path}"
    else:
        active_doc.Save()
        return f"Document saved: {active_doc.Name}"
        
@app.tool("close_document")
def close_document(document_name: str):
    """Close a document by name"""
    catia = get_catia_application()
    documents = catia.Documents
    for i in range(1, documents.Count + 1):
        doc = documents.Item(i)
        if doc.Name == document_name:
            doc.Close()
            return f"Closed document: {document_name}"
    return f"Document not found: {document_name}"

@app.tool("create_sketch")
def create_sketch(plane: str, name: str = None):
    """Create a new sketch on specified plane"""
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    bodies = part.Bodies
    body = bodies.Item(1)
    sketches = body.Sketches
    
    # Get the plane reference
    plane_map = {
        "xy": part.OriginElements.PlaneXY,
        "yz": part.OriginElements.PlaneYZ,
        "zx": part.OriginElements.PlaneZX,
    }
    ref_plane = plane_map[plane]
    
    sketch = sketches.Add(ref_plane)
    # Attempt to set the sketch name if provided
    if name:
        try:
            # try common name setters
            setattr(sketch, 'Name', name)
        except Exception:
            try:
                sketch.set_Name(name)
            except Exception:
                pass
    return f"Created sketch on {plane} plane" + (f" with name {name}" if name else "")

@app.tool("create_pad")
def create_pad(length: float):
    """Create a pad from the current sketch"""
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    shape_factory = part.ShapeFactory
    
    # Get the last sketch
    bodies = part.Bodies
    body = bodies.Item(1)
    sketches = body.Sketches
    sketch = sketches.Item(sketches.Count)
    
    pad = shape_factory.AddNewPad(sketch, length)
    part.Update()
    return f"Created pad with length {length} mm"
        
@app.tool("create_pocket")
def create_pocket(depth: float):
    """Create a pocket from the current sketch"""
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    shape_factory = part.ShapeFactory
    
    # Get the last sketch
    bodies = part.Bodies
    body = bodies.Item(1)
    sketches = body.Sketches
    sketch = sketches.Item(sketches.Count)
    
    pocket = shape_factory.AddNewPocket(sketch, depth)
    part.Update()
    return f"Created pocket with depth {depth} mm"

@app.tool("get_part_bodies")
def get_part_bodies():
    """Get all bodies in the active part"""
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    bodies = part.Bodies
    
    body_list = []
    for i in range(1, bodies.Count + 1):
        body = bodies.Item(i)
        body_list.append({
            "name": body.Name,
        })
    return body_list
        
@app.tool("update_part")
def update_part():
    """Update the active part"""
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    part.Update()
    return "Part updated successfully"



@app.tool("create_rectangle")
def create_rectangle(x, y, width, height, centered=True):
    catia = get_catia_application()
    part_doc = catia.ActiveDocument
    part = part_doc.Part
    bodies = part.Bodies
    body = bodies.Item(1)
    sketches = body.Sketches

    # Get the current (last) sketch
    sketch = sketches.Item(sketches.Count)
    # Make sure the sketch is the in-work object and open it for edition
    part.InWorkObject = sketch
    factory = sketch.OpenEdition()

    # Calculate corner coordinates
    if centered:
        x1 = x - width / 2
        y1 = y - height / 2
        x2 = x + width / 2
        y2 = y + height / 2
    else:
        x1 = x
        y1 = y
        x2 = x + width
        y2 = y + height

    # Create four lines to form a rectangle
    # (Using the factory returned by OpenEdition)
    _ = factory.CreatePoint(x1, y1)
    line1 = factory.CreateLine(x1, y1, x2, y1)  # Bottom
    line2 = factory.CreateLine(x2, y1, x2, y2)  # Right
    line3 = factory.CreateLine(x2, y2, x1, y2)  # Top
    line4 = factory.CreateLine(x1, y2, x1, y1)  # Left

    # Optional: set ReportName (left as-is for debugging/traceability)
    try:
        line1.ReportName = 1
        line2.ReportName = 2
        line3.ReportName = 3
        line4.ReportName = 4
    except Exception:
        # Some CATIA objects may not expose ReportName; ignore if unavailable
        pass

    # Close sketch edition and update part
    try:
        sketch.CloseEdition()
    except Exception:
        # If CloseEdition isn't available or fails, continue — the next
        # operations will most likely raise a descriptive COM error.
        pass
    try:
        part.Update()
    except Exception:
        pass

    result = {
        "message": f"Created rectangle at ({x}, {y}) with width={width}mm, height={height}mm",
        "centered": centered,
        "corners": {
            "bottom_left": [x1, y1],
            "top_right": [x2, y2]
        }
    }
    return result


@app.tool("execute_macro")
def execute_macro(macro_path: str, module_name: str, function_name: str):
    """Execute a CATIA macro"""
    catia = get_catia_application()
    system_service = catia.SystemService
    system_service.ExecuteScript(
        macro_path,
        module_name,
        function_name,
        []
    )
    return f"Executed macro: {function_name}"

if __name__ == "__main__":
    app.run()
