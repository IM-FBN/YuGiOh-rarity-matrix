import os
import json
import requests
import sqlite3
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Border, Side

# Constants
COMMON_LIMIT = 3
FULL_REPORT = 1
ALTERNATE_BG_COLOR = 1
DB_DIR = str(Path.cwd()) + os.sep + "expansions"

# SQL Statement
SELECT_STATEMENT = """
    SELECT datas.id, datas.alias, texts.name 
    FROM datas
    INNER JOIN texts ON datas.id = texts.id 
    WHERE instr(texts.name, '(GOAT)') OR instr(texts.name, '(Pre-Errata)')
    ORDER BY texts.name
"""

def find_main(alt_id_list):
    """
    Function to find the main id from the alternate id list
    """
    alt_id_list = set(int(item) for item in alt_id_list)
    hits = find_hits()

    # Convert the JSON string into a list of dictionaries
    cards = json.loads(json.dumps(hits, indent=4, ensure_ascii=False))

    # Search the card list and output main_id if alt_id is in the list
    return [str(card['main_id']) for card in cards if card['alt_id'] in alt_id_list]


def find_hits():
    """
    Function to find hits from the databases
    """
    databases = sorted(filter(lambda f: f.endswith(".cdb"), os.listdir(DB_DIR)))
    hits = []
    encountered_ids = set()

    for db in databases:
        db_path = (DB_DIR + os.sep + db)
        with sqlite3.connect(db_path) as con:
            cursor = con.cursor()
            cursor.execute(SELECT_STATEMENT)
            rows = cursor.fetchall()
            for row in rows:
                alt_id, main_id, name = row
                name = name.strip()

                if alt_id in encountered_ids:
                    continue

                _type = "GOAT" if "(GOAT)" in name else ("Pre-Errata" if "(Pre-Errata)" in name else "Unknown")

                hits.append({
                    "alt_id": alt_id,
                    "main_id": main_id,
                    "type": _type,
                    "name": name,
                    "db": db
                })

                encountered_ids.add(alt_id)

    hits.sort(key=lambda hit: hit["name"])
    return hits

def parse_card_data(card_data):
    """
    Function to parse card data
    """
    # Create an empty data structure to store the results
    result = {}

    if "data" in card_data:
        for info in card_data["data"]:
            # Create an empty list to store the rarity values
            rarities = {}

            # Go through each element in the "card_sets" list
            for card_set in info["card_sets"]:
                # Add set code and price to the appropriate rarity list
                rarity_list = rarities.setdefault(card_set["set_rarity"], [])
                rarity_list.append(card_set["set_code"].split("-")[0] + " " + card_set["set_price"] + "$")

            # Add the card name, type and the list of rarity values to the result
            result.update({"name": info["name"], "type": info["frameType"], "rarity": rarities})

    else:
        print(card_data)
        print("No data in card_data")
    
    return result
def decklist_request(decklist):
    """
    Function to request decklist from the API
    """
    decklist = set(decklist)
    response_data = []
    error_list = []

    def update_progress(current, total):
        """
        Function to update and print progress
        """
        if total - current < 4:
            progress = 100.0
        else:
            progress = (current + 1) / total * 100
        print(f"Progress: {progress:.2f}%")

    total_cards = len(decklist)

    for i, id in enumerate(decklist):
        response = requests.get(f'https://db.ygoprodeck.com/api/v7/cardinfo.php?id={id}').json()
        if "error" not in response:
            response_data.append(response)
        else:
            #print(f"Error with card ID {id}: {response['error']}")
            error_list.append(id)
        
        # Update progress every fourth card
        if (i + 1) % 4 == 0:
            update_progress(i, total_cards)

    main_id_list = find_main(error_list)

    # Check if any file has the extension ".cdb"
    if any(file.endswith(".cdb") for file in os.listdir(DB_DIR)):
        print("\n.cdb Files found")
        total_main_cards = len(main_id_list)
        for i, id in enumerate(main_id_list):
            response = requests.get(f'https://db.ygoprodeck.com/api/v7/cardinfo.php?id={id}').json()
            # Skip parsing if there was an error with the request
            if "error" in response:
                print(f"Error with card ID {id}: {response['error']}")
                continue
            response_data.append(response)
            
            # Update progress every fourth card
            if (i + 1) % 4 == 0:
                update_progress(i, total_main_cards)
    else:
        print("\nThere are no .cdb files in the specified directory.")
        # Write errors into a new text file
        print(error_list)
        with open("error_cards.ydk", "w") as file:
            # Iterate through each card in the error_list
            for card in error_list:
                # Write the card into the file
                file.write(str(card) + "\n")
    
    return response_data


def data_to_excel(parsed_data, full_report=FULL_REPORT, alternate_bg_color=ALTERNATE_BG_COLOR):
    """
    Function to write parsed data to an Excel file
    """
    # Sort card data first by type and then by name
    data_sorted = sorted(parsed_data, key=lambda card: (1 if card["type"].lower() in ["fusion", "synchro", "xyz", "link"] else 0, card["type"], card["name"]))

    # Create column names for the table
    rarities = ["Common", "Rare", "Super Rare", "Ultra Rare", "Secret Rare", "Ultimate Rare"]
    if full_report:
        for card in data_sorted:
            for rarity in card["rarity"].keys():
                if rarity not in rarities:
                    rarities.append(rarity)
    column_names = ["Type", "Name"] + rarities

    # Create Excel workbook
    workbook = Workbook()
    sheet = workbook.active

    # Write column names into the table
    for i, name in enumerate(column_names):
        cell = sheet.cell(row=1, column=i + 1, value=name)
        # Color every other column with a light gray background color if alternate_bg_color is True
        if alternate_bg_color and i % 2 == 1:
            fill = PatternFill(start_color='E3E3E3', end_color='E3E3E3', fill_type='solid')
            cell.fill = fill

    # Write data into the table
    # Freeze the first column
    sheet.freeze_panes = 'C2'

    for i, card in enumerate(data_sorted):
        row = [card["type"], card["name"]]
        for j, rarity in enumerate(rarities):
            sets = card["rarity"].get(rarity, [])
            if rarity == "Common":
                # Limit the number of entries for "Common"
                if len(sets) > COMMON_LIMIT:
                    row.append("\n".join(sets[:COMMON_LIMIT]) + "\n ({} weitere)".format(len(sets) - COMMON_LIMIT))
                else:
                    row.append("\n".join(sets))
            else:
                row.append("\n".join(sets))
        for j, value in enumerate(row):
            cell = sheet.cell(row=i+2, column=j+1, value=value)
            # Set alignment of cell to top and left
            cell.alignment = Alignment(horizontal="left", vertical="top")
            # Auto-size column width based on cell content
            column_letter = get_column_letter(j+1)
            column_dimension = sheet.column_dimensions[column_letter]
            column_dimension.auto_size = True
            # Color every other column with a light gray background color if alternate_bg_color is True
            # Set top border of cell
            border = Border(top=Side(border_style="thin", color="000000"))
            cell.border = border

            if j == 0:  # Column A
                if value.lower() == "effect":
                    cell.fill = PatternFill(start_color='FF8B53', end_color='FF8B53', fill_type='solid')
                elif value.lower() == "spell":
                    cell.fill = PatternFill(start_color='1D9E74', end_color='1D9E74', fill_type='solid')
                elif value.lower() == "trap":
                    cell.fill = PatternFill(start_color='BC5A84', end_color='BC5A84', fill_type='solid')
                elif value.lower() == "fusion":
                    cell.fill = PatternFill(start_color='A086B7', end_color='A086B7', fill_type='solid')
                elif value.lower() == "ritual":
                   cell.fill = PatternFill(start_color='9DB5CC', end_color='9DB5CC', fill_type='solid')
                elif value.lower() == "synchro":
                   cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
                elif value.lower() == "normal":
                   cell.fill = PatternFill(start_color='FDE68A', end_color='FDE68A', fill_type='solid')
                elif value.lower() == "xyz":
                   cell.fill = PatternFill(start_color='63666A', end_color='63666A', fill_type='solid')
                elif value.lower() == "link":
                   cell.fill = PatternFill(start_color='00008B', end_color='00008B', fill_type='solid')

            if alternate_bg_color and j % 2 == 1:
                fill = PatternFill(start_color='E3E3E3', end_color='E3E3E3', fill_type='solid')
                cell.fill = fill

    # Enable text wrap for all cells
    for row in sheet.rows:
        for cell in row:
            cell.alignment = Alignment(wrapText=True, horizontal="right", vertical="top")

    # Save Excel file
    workbook.save("rarity_matrix.xlsx")
    return "saved"





def main():
    decklist = []

    for file_name in os.listdir():
        if file_name.endswith(".ydk"):
            with open(file_name) as file:
                for line in file:
                    entry = line.strip()
                    if entry[0].isdigit():
                        decklist.append(entry)

    decklist_data = decklist_request(decklist)
    parsed_data = [parse_card_data(card_data) for card_data in decklist_data]
    data_to_excel(parsed_data, FULL_REPORT, ALTERNATE_BG_COLOR)
    print("FILE CREATED - FINISHED")


if __name__ == "__main__":
    main()
