import omero
import omero.gateway
import omero.model
import omero.rtypes
import json
import getpass

def import_rois_from_json(json_file, conn):
    """Import ROIs from a JSON file into OMERO if the Image ID matches."""
    with open(json_file, 'r') as file:
        data = json.load(file)
    
    image_id = data["Image_ID"]
    image = conn.getObject("Image", image_id)
    
    if not image:
        print(f"Image with ID {image_id} not found in OMERO.")
        return
    
    roi_service = conn.getUpdateService()
    for roi_data in data["ROIs"]:
        roi = omero.model.RoiI()
        roi.setImage(image._obj)
        
        for shape_data in roi_data["Shapes"]:
            shape = None
            if shape_data["type"] == "Ellipse":
                shape = omero.model.EllipseI()
                shape.setX(omero.rtypes.rdouble(shape_data["x"]))
                shape.setY(omero.rtypes.rdouble(shape_data["y"]))
                shape.setRadiusX(omero.rtypes.rdouble(shape_data["radiusX"]))
                shape.setRadiusY(omero.rtypes.rdouble(shape_data["radiusY"]))
            elif shape_data["type"] == "Polygon":
                shape = omero.model.PolygonI()
                shape.setPoints(omero.rtypes.rstring(shape_data["points"]))
            elif shape_data["type"] == "Polyline":
                shape = omero.model.PolylineI()
                shape.setPoints(omero.rtypes.rstring(shape_data["points"]))
            elif shape_data["type"] == "Label":
                shape = omero.model.LabelI()
                shape.setTextValue(omero.rtypes.rstring(shape_data["text"]))
                shape.setX(omero.rtypes.rdouble(shape_data["x"]))
                shape.setY(omero.rtypes.rdouble(shape_data["y"]))
            elif shape_data["type"] == "Rectangle":
                shape = omero.model.RectangleI()
                shape.setX(omero.rtypes.rdouble(shape_data["x"]))
                shape.setY(omero.rtypes.rdouble(shape_data["y"]))
                shape.setWidth(omero.rtypes.rdouble(shape_data["width"]))
                shape.setHeight(omero.rtypes.rdouble(shape_data["height"]))
            elif shape_data["type"] == "Line":
                shape = omero.model.LineI()
                shape.setX1(omero.rtypes.rdouble(shape_data["x1"]))
                shape.setY1(omero.rtypes.rdouble(shape_data["y1"]))
                shape.setX2(omero.rtypes.rdouble(shape_data["x2"]))
                shape.setY2(omero.rtypes.rdouble(shape_data["y2"]))
            elif shape_data["type"] == "Mask":
                shape = omero.model.MaskI()
                shape.setX(omero.rtypes.rdouble(shape_data["x"]))
                shape.setY(omero.rtypes.rdouble(shape_data["y"]))
                shape.setWidth(omero.rtypes.rdouble(shape_data["width"]))
                shape.setHeight(omero.rtypes.rdouble(shape_data["height"]))
            
            if shape:
                shape.setTheZ(omero.rtypes.rint(shape_data["theZ"]))
                shape.setTheT(omero.rtypes.rint(shape_data["theT"]))
                
                if "fillColor" in shape_data and shape_data["fillColor"] is not None:
                    shape.setFillColor(omero.rtypes.rint(shape_data["fillColor"]))
                if "strokeColor" in shape_data and shape_data["strokeColor"] is not None:
                    shape.setStrokeColor(omero.rtypes.rint(shape_data["strokeColor"]))
                if "strokeWidth" in shape_data and shape_data["strokeWidth"] is not None:
                    shape.setStrokeWidth(omero.model.LengthI(shape_data["strokeWidth"], omero.model.enums.UnitsLength.PIXEL))
                if "fontFamily" in shape_data:
                    shape.setFontFamily(omero.rtypes.rstring(shape_data["fontFamily"]))
                if "fontSize" in shape_data and shape_data["fontSize"] is not None:
                    shape.setFontSize(omero.model.LengthI(shape_data["fontSize"], omero.model.enums.UnitsLength.POINT))
                if "fontStyle" in shape_data:
                    shape.setFontStyle(omero.rtypes.rstring(shape_data["fontStyle"]))
                
                roi.addShape(shape)
        
        roi_service.saveAndReturnObject(roi)
    print(f"ROIs successfully imported for Image ID {image_id}.")

def main():
    host = input("Enter OMERO server hostname: ")
    username = input("Enter OMERO username: ")
    password = getpass.getpass("Enter OMERO password: ")
    json_file = input("Enter path to JSON file: ")
    
    conn = omero.gateway.BlitzGateway(username, password, host=host)
    conn.connect()
    
    import_rois_from_json(json_file, conn)
    
    conn.close()

if __name__ == "__main__":
    main()
