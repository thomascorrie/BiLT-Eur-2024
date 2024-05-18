#!python3
"""Creates new items in Notion Asset Model from selected Revit elements"""

import os
import sys
import json
sys.path.append(r"C:\ProgramData\Anaconda3\Lib\site-packages")
from pyrevit import DB, revit, script
from pyrevit.revit.db.transaction import Transaction
from notion_client import Client

def get_secrets():
    with open(os.path.dirname(__file__) + '/../../../../secrets.json') as secrets_file:
        secrets = json.load(secrets_file)
    return secrets

secrets = get_secrets()
token = secrets.get("token")
spaces_db = secrets.get("spaces_db")
assets_db = secrets.get("assets_db")
os.environ['NOTION KEY'] = token
notion = Client(auth=os.environ['NOTION KEY'])

fam_name = "71LA_Asset"

doc = revit.doc

pvp = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_MARK))
rule = DB.FilterStringRule(pvp, DB.FilterStringBeginsWith(), "AST-")
filter = DB.ElementParameterFilter(rule)
elements = DB.FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()

# Get Generic Asset Family Type 
fam_sym = None
family_symbols = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_GenericModel).WhereElementIsElementType() 
for f in family_symbols:
    try: 
        tn = f.GetParameters("Type Name")[0].AsString()
        if tn == fam_name:
            fam_sym = f
            print("Generic Asset Family Type Found: " + tn)
    except:
        continue

if fam_sym:
    marks = set()
    for e in elements:
        mark = e.GetParameters("Mark")[0]
        mark_value = mark.AsString()
        marks.add(mark_value)

    # Get spaces from Notion
    results_spaces = notion.databases.query(
        **{
            "database_id": spaces_db,
            },
        ).get("results")

    dict_space_id_name = {}
    dict_space_name_id = {}
    for r in results_spaces:
        # print(r["id"] + ": " + r['properties']['Name']['title'][0]['plain_text'])
        dict_space_name_id[r['properties']['Name']['title'][0]['plain_text']] = r["id"]
        dict_space_id_name[r["id"]] = r['properties']['Name']['title'][0]['plain_text']

    # Get rooms from Revit and creates a dictionary
    dict_spaces_id_room = {}
    rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
    for r in rooms:
        room_name = r.GetParameters("Name")[0].AsString()
        dict_spaces_id_room[dict_space_name_id[room_name]] = r

    # Get Assets from Notion set to be in Model
    results = notion.databases.query(
        **{
            "database_id": assets_db,
            "filter": {
                "property": "In Model?",
                    "checkbox": {
                        "equals": True
                    }
            },
        }
        ).get("results")

    # Process results
    processed = 0
    for result in results:
        mark_value = result['properties']['ID']['unique_id']['prefix'] + "-" + str(result['properties']['ID']['unique_id']['number'])
        if mark_value in marks: # Check whether asset already exists in the model
            continue
        report = mark_value
        space_name = None
        space_element = None
        # Space
        try:
            space_id = result['properties']['Spaces']['relation'][0]['id']
            space_name = dict_space_id_name[space_id]
            space_element = dict_spaces_id_room[space_id]
        except:
            space_element = None
        # Size
        try:
            size = result['properties']['Size (w*d*h)']['rich_text'][0]['plain_text']
            width = int(size.split('x')[0])
            depth = int(size.split('x')[1])
            height = int(size.split('x')[2])
            width_f = width / 304.8 # Revit API in decimal feet
            depth_f = depth / 304.8 # Revit API in decimal feet
            height_f = height / 304.8 # Revit API in decimal feet
        except:
            size = None
        # Create Element in Revit
        with Transaction('SetParam') as rvtxn:
            report = report + " " + result['properties']['Name']['title'][0]['plain_text']
            if space_element:
                e = doc.Create.NewFamilyInstance(space_element.Location.Point, fam_sym, DB.Structure.StructuralType.NonStructural)
                report = report + " placed in " + space_name
            else:
                e = doc.Create.NewFamilyInstance(DB.XYZ[0,0,0], fam_sym, DB.Structure.StructuralType.NonStructural)
            mark = e.GetParameters("Mark")[0]
            mark.Set(mark_value)
            comments = e.GetParameters("Comments")[0]
            comments.Set(result['properties']['Name']['title'][0]['plain_text'])
            if size:
                width_param = e.GetParameters("Width")[0]
                width_param.Set(width_f)
                depth_param = e.GetParameters("Depth")[0]
                depth_param.Set(depth_f)
                height_param = e.GetParameters("Height")[0]
                height_param.Set(height_f)
                report = report + " " + str(width) + "x" + str(depth) + "x" + str(height)
            else:
                report = report + " no size given"
            print(report)
        processed = processed + 1
    if processed == 0:
        print("No new assets to place in the model")
else:
    print("Generic Asset Family Type Not Found")

