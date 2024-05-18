#!python3
"""Creates new items in Notion Asset Model from selected Revit elements"""

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

def addToNotion(e):
    print(e.Name)

    new_page = {
            "Name": {"title": [{"text": {"content": e.Name}}]},
            "In Model?": {"checkbox": True},

            }

    room = None
    try:
        room = e.Room.GetParameters("Name")[0].AsString()
        results = notion.databases.query(
        **{
            "database_id": spaces_db,
            "filter": {
                "property": "Name",
                    "rich_text": {
                        "equals": room
                    }
            },
        }
        ).get("results")
        room_id = results[0]['id']
        new_page["Spaces"] =  {"relation": [
                        {"id": room_id}
                        ]
                        }
        print(room, room_id)
    except:
        print("-----No Room Found")
    
    
    new_page = notion.pages.create(parent={"database_id": assets_db}, properties=new_page)
    print("----Added to Notion Asset Model")
    #print(new_page)
    asset_id = new_page['properties']['ID']['unique_id']['prefix'] + "-" + str(new_page['properties']['ID']['unique_id']['number'])
    asset_url = str(new_page['url'])
    return asset_id, asset_url

secrets = get_secrets()
token = secrets.get("token")
spaces_db = secrets.get("spaces_db")
assets_db = secrets.get("assets_db")
os.environ['NOTION KEY'] = token
notion = Client(auth=os.environ['NOTION KEY'])

doc = revit.doc
uidoc = revit.uidoc

selection = uidoc.Selection.GetElementIds()

if len(selection) > 0:
    for id in selection:
        e = doc.GetElement(id)
        mark = e.GetParameters("Mark")[0]
        mark_value = mark.AsString()
        if mark_value and mark_value.startswith("AST-"):
            print(e.Name + " already has an Asset Reference of " + mark_value)
            continue
        print(e.GetParameters("Mark")[0].AsString())
        asset_id, asset_url = addToNotion(e) # Add to Notion
        type = doc.GetElement(e.GetTypeId())
        url_type = type.GetParameters("URL")[0]
        with Transaction('CommandName') as rvtp:
            mark.Set(asset_id)
            url_type.Set(asset_url)
else:
    print("No elements selected")


