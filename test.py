import json

import pandas as pd
from onshape_client.client import Client
from onshape_client.onshape_url import OnshapeElement


def findIDs(bom_dict, IDName):
    global process1ID, process2ID
    for head in bom_dict["headers"]:
        if head['name'] == IDName:
            return head['id']


def getPartsDict(bom_dict):
    partDict = {}
    rows = bom_dict.get("rows", [])
    for row in rows:
        part_name = row.get("headerIdToValue", {}).get("57f3fb8efa3416c06701d60d", "Unknown")
        quantity = row.get("headerIdToValue", {}).get("5ace84d3c046ad611c65a0dd", "N/A")
        part_material = row.get("headerIdToValue", {}).get("57f3fb8efa3416c06701d615", "Unknown")
        if part_material != "N/A" and part_material is not None:
            partDict[part_name] = (int(quantity), part_material["displayName"])
        else:
            partDict[part_name] = (int(quantity), "No material")

    return partDict


def getExcelBom(Process1ID, Process2ID, DescriptionID):
    # Assuming `bom_dict` is the dictionary parsed from your JSON string
    rows = bom_dict.get("rows", [])

    # Define the desired columns
    columns = ["Part Name", "Assem", "Part", "Revision", "Qty", "Material",
               "Dimensions - PLEASE DO NOT COUNT ON THIS INFORMATION", "CAD",
               "Pre-process", "Done Pre-Process", "Process 1", "Done Process 1", "Process 2"]

    # Prepare data for each column based on your BOM structure
    data = []
    for row in rows:
        part_name = row.get("headerIdToValue", {}).get("57f3fb8efa3416c06701d60d", "Unknown")
        qty = row.get("headerIdToValue", {}).get("5ace84d3c046ad611c65a0dd", "N/A")
        material = row.get("headerIdToValue", {}).get("57f3fb8efa3416c06701d615", "Unknown")
        if material != "N/A" and material is not None:
            material = material['displayName']
        else:
            material = "No material"

        # Example placeholders for other fields; fill them in based on your BOM structure
        data.append({
            "Part Name": part_name,
            "Assem": "Assembly Info",  # Fill this with actual assembly data if available
            "Part": row.get("itemSource", {}).get("partId", "Unknown"),
            "Revision": row.get("headerIdToValue", {}).get("57f3fb8efa3416c06701d610", "Unknown"),
            "Qty": qty,
            "Material": material,
            "Dimensions - PLEASE DO NOT COUNT ON THIS INFORMATION": row.get("headerIdToValue", {}).get(DescriptionID,
                                                                                                       "Unknown"),
            "CAD": row.get("itemSource", {}).get("viewHref", ""),
            "Pre-process": "Pre-process data",
            "Done Pre-Process": False,  # Set based on BOM or FeatureScript data
            "Process 1": row.get("headerIdToValue", {}).get(Process1ID, "Unknown"),
            # To be set via FeatureScript
            "Done Process 1": False,  # Set based on process completion
            "Process 2": row.get("headerIdToValue", {}).get(Process2ID, "Unknown")
            # To be set via FeatureScript
        })

    # Create DataFrame and write to Excel
    df = pd.DataFrame(data, columns=columns)
    df.to_excel("BOM_output.xlsx", index=False)
    print("Excel file generated as BOM_output.xlsx")


if __name__ == '__main__':
    # Replace with your actual API keys
    access_key = 'iVTJDrE6RTFeWKRTj8cF4VCa'
    secret_key = 'hjhZYvSX1ylafeku5a7e4wDsBXUNQ6oKynl6HnocHTTddy0Q'
    base = 'https://cad.onshape.com'

    # Initialize the Onshape client
    client = Client(configuration={"base_url": base,
                                   "access_key": access_key,
                                   "secret_key": secret_key})

    # URL of the Onshape document (example BOM document link)
    document_url = 'https://cad.onshape.com/documents/97c0beec3c1a39d69e523379/w/efc327b803d3dae6436b287d/e/e0c6823fb5fe0aeef6fc5939'

    # Parse the document URL to extract the document, workspace, and element IDs
    element = OnshapeElement(document_url)
    did = element.did
    wid = element.wvmid
    eid = element.eid
    fixed_url = '/api/v9/assemblies/d/did/w/wid/e/eid/bom'

    method = 'GET'

    params = {}
    payload = {}
    headers = {'Accept': 'application/vnd.onshape.v1+json; charset=UTF-8;qs=0.1',
               'Content-Type': 'application/json'}

    fixed_url = fixed_url.replace('did', did)
    fixed_url = fixed_url.replace('wid', wid)
    fixed_url = fixed_url.replace('eid', eid)
    print("Connecting to Onshape's API...")
    response = client.api_client.request(method, url=base + fixed_url, query_params=params, headers=headers,
                                         body=payload)
    print("Onshape API Connected.")


    bom_dict = dict(json.loads(response.data))

    part_Dict = getPartsDict(bom_dict)
    process1ID = findIDs(bom_dict, "Process 1")
    process2ID = findIDs(bom_dict, "Process 2")
    DescriptionID = findIDs(bom_dict, "Description")
    getExcelBom(process1ID, process2ID, DescriptionID)

