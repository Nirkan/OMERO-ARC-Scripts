import omero
import omero.gateway
import omero.model
import getpass
import json
import os

def export_rois_as_json(image_id, conn):
    """Export ROIs from the given OMERO image as a JSON string."""
    roi_service = conn.getRoiService()
    result = roi_service.findByImage(image_id, None)
    roi_list = []
    
    for roi in result.rois:
        shapes = []
        for s in roi.copyShapes():
            shape = {
                "id": s.getId().getValue(),
                "theT": s.getTheT().getValue() if s.getTheT() else None,
                "theZ": s.getTheZ().getValue() if s.getTheZ() else None,
                "fillColor": s.getFillColor().getValue() if s.getFillColor() else None,
                "strokeColor": s.getStrokeColor().getValue() if s.getStrokeColor() else None,
                "strokeWidth": s.getStrokeWidth().getValue() if s.getStrokeWidth() else None,
                "fontFamily": s.getFontFamily().getValue() if s.getFontFamily() else None,
                "fontSize": s.getFontSize().getValue() if s.getFontSize() else None,
                "fontStyle": s.getFontStyle().getValue() if s.getFontStyle() else None,
                "textValue": s.getTextValue().getValue() if s.getTextValue() else None
            }
            
            if type(s) == omero.model.RectangleI:
                shape.update({
                    "type": "Rectangle",
                    "x": s.getX().getValue(),
                    "y": s.getY().getValue(),
                    "width": s.getWidth().getValue(),
                    "height": s.getHeight().getValue()
                })
            elif type(s) == omero.model.EllipseI:
                shape.update({
                    "type": "Ellipse",
                    "x": s.getX().getValue(),
                    "y": s.getY().getValue(),
                    "radiusX": s.getRadiusX().getValue(),
                    "radiusY": s.getRadiusY().getValue()
                })
            elif type(s) == omero.model.PointI:
                shape.update({
                    "type": "Point",
                    "x": s.getX().getValue(),
                    "y": s.getY().getValue()
                })
            elif type(s) == omero.model.LineI:
                shape.update({
                    "type": "Line",
                    "x1": s.getX1().getValue(),
                    "x2": s.getX2().getValue(),
                    "y1": s.getY1().getValue(),
                    "y2": s.getY2().getValue(),
                    "markerStart": s.getMarkerStart().getValue() if s.getMarkerStart() is not None else None,
                    "markerEnd": s.getMarkerEnd().getValue() if s.getMarkerEnd() is not None else None
                })
            elif type(s) == omero.model.MaskI:
                shape.update({
                    "type": "Mask",
                    "x": s.getX().getValue(),
                    "y": s.getY().getValue(),
                    "width": s.getWidth().getValue(),
                    "height": s.getHeight().getValue()
                })
            elif type(s) == omero.model.PolygonI:
                shape.update({
                    "type": "Polygon",
                    "points": s.getPoints().getValue()
                })
            elif type(s) == omero.model.PolylineI:
                shape.update({
                    "type": "Polyline",                    
                    "points": s.getPoints().getValue()
                })
            elif type(s) == omero.model.LabelI:
                shape.update({
                    "type": "Label",
                    "text": s.getTextValue().getValue(),
                    "x": s.getX().getValue() if s.getX() else 0,
                    "y": s.getY().getValue() if s.getY() else 0,
                })
            shapes.append(shape)
        
        roi_list.append({
            "ROI_ID": roi.getId().getValue(),
            "Shapes": shapes
        })
        print(roi_list)
    
    return json.dumps({"Image_ID": image_id, "ROIs": roi_list}, indent=4)

def main():
    host = input("Enter OMERO host: ")
    username = input("Enter OMERO username: ")
    password = getpass.getpass("Enter OMERO password: ")
    
    conn = omero.gateway.BlitzGateway(username, password, host=host)
    conn.connect()
    
    user_input = input("Enter OMERO Dataset ID or Image ID: ")
    
    if user_input.isdigit():
        dataset = conn.getObject("Dataset", int(user_input))
        if dataset:
            folder_name = dataset.getName().replace(" ", "_")
            os.makedirs(folder_name, exist_ok=True)
            
            for image in dataset.listChildren():
                image_id = image.getId()
                image_name = image.getName().replace(" ", "_")
                json_output = export_rois_as_json(image_id, conn)
                filename = os.path.join(folder_name, f"{image_name}_ID{image_id}_rois.json")
                with open(filename, "w") as json_file:
                    json_file.write(json_output)
                print(f"ROIs exported to {filename}")
        else:
            image = conn.getObject("Image", int(user_input))
            if image:
                image_id = image.getId()
                image_name = image.getName().replace(" ", "_")
                json_output = export_rois_as_json(image_id, conn)
                filename = f"{image_name}_ID{image_id}_rois.json"
                with open(filename, "w") as json_file:
                    json_file.write(json_output)
                print(f"ROIs exported to {filename}")
            else:
                print("Invalid ID. No Dataset or Image found.")
    else:
        print("Invalid input. Please enter a numeric Dataset or Image ID.")
    
    conn.close()

if __name__ == "__main__":
    main()
