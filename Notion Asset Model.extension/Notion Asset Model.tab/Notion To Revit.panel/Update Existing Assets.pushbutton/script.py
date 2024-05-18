#!python3
"""Updates Existing Assets in Revit with data from Notion"""

import os
import sys
import json
sys.path.append(r"C:\ProgramData\Anaconda3\Lib\site-packages")
from pyrevit import DB, revit
from pyrevit.revit.db.transaction import Transaction
from notion_client import Client

def get_secrets():
    with open(os.path.dirname(__file__) + '/../../../../secrets.json') as secrets_file:
        secrets = json.load(secrets_file)
    return secrets

def set_parameter(e, name, value):
    param = e.GetParameters(name)[0]
    with Transaction('SetParam') as rvtxn:
        param.Set(value)

secrets = get_secrets()
token = secrets.get("token")
spaces_db = secrets.get("spaces_db")
assets_db = secrets.get("assets_db")
os.environ['NOTION KEY'] = token
notion = Client(auth=os.environ['NOTION KEY'])

doc = revit.doc
uidoc = revit.uidoc

selection = uidoc.Selection.GetElementIds()

pvp = DB.ParameterValueProvider(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_MARK))
rule = DB.FilterStringRule(pvp, DB.FilterStringBeginsWith(), "AST-")
filter = DB.ElementParameterFilter(rule)
elements = DB.FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements()

for e in elements:
    type = doc.GetElement(e.GetTypeId())
    mark = e.GetParameters("Mark")[0]
    mark_value = mark.AsString()
    mark_num = int(mark_value.split('-')[-1])
    print(e.Name, mark_value)
    results = notion.databases.query(
    **{
        "database_id": assets_db,
        "filter": {
            "property": "ID",
                "number": {
                    "equals": mark_num
                }
        },
    }
    ).get("results")
    uniclass_code = None
    try:
        uniclass_code = results[0]['properties']['Uniclass Code']['rich_text'][0]['plain_text']
    except:
        uniclass_code = None
    if uniclass_code:
        set_parameter(type, "Classification.Uniclass.Pr.Number", uniclass_code)
    uniclass_description = None
    try:
        uniclass_description = results[0]['properties']['Uniclass Description']['rich_text'][0]['plain_text']
    except:
        uniclass_description = None
    if uniclass_description:
        set_parameter(type, "Classification.Uniclass.Pr.Description", uniclass_description)

print("Complete")
    


