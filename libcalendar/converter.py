import json
import pandas as pd

def convertFromExcel(path):
    sheets_dict = pd.read_excel(path, sheet_name=None, index_col=None)

    items = []
    calendar = {
        "items": items
    }

    reserved = ["collections", "date", "label", "thumbnail", "url"]

    for name, sheet in sheets_dict.items():
        if name == "metadata":
            links = []
            for index, row in sheet.iterrows():
                key = row["key"]
                literal = row["literal"]
                uri = row["uri"]

                if key == "link":
                    links.append({
                        "label": literal,
                        "value": uri
                    })

                else:
                    calendar[key] = literal

            calendar["links"] = links

        elif name == "items":
            for index, row in sheet.iterrows():
                metadata = []
                item = {
                    "metadata": metadata
                }
                items.append(item)
                for key, col in sheet.iteritems():
                    value = row[key]

                    if pd.isnull(value):
                        continue

                    value = str(value)

                    if key == "date":
                        value = value.split(" ")[0]

                    if key not in reserved:
                        value = value.split("|")
                        metadata.append({
                            "label": key,
                            "value": value
                        })
                    else:
                        if key == "collections":
                            value = value.split("|")
                    item[key] = value

    return calendar
