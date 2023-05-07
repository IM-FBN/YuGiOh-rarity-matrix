# YuGiOh-rarity-matrix

Creating an Excel sheet from a Duelingbook or YGOPro Decklist to display all rarities of all individual cards

---

## Description

This tool is intended for checking the available card rarities of certain decklists. The code checks the script path for existing .ydk decklists, parses the data to the ygoprodeck API, and creates an Excel file providing the available released rarities and prices from TCGPlayer.com for each card.

---

## How to use

The code uses the following external libraries which must be installed in order to use the code:
- requests: pip install requests
- openpyxl: pip install openpyxl

1. Place any amount of .ydk files (tested for DuelingBook and EDOPro) inside the script folder.
2. Run the script via command shell or batch file.
3. An Excel file should be created.

*If you're using EDOPro files with pre-errata card versions, place the extensions folder from the EDOPro directory inside the script folder. The script checks if there is a pre-errata version and changes the ID to be properly found in the ygoprodeck API.

---

## For anyone using an Android device

There is an app called "Pydroid 3" where you can run Python scripts. For viewing the Excel file, I've found WPS Office quite comfortable. Just make sure to activate word wrap while marking all cells.

---

## Disclaimer

I'm quite new to programming and did this to reduce time while building decks I've tested online in a certain rarity. However, since I'm not too experienced, I'm glad for any code reviews or suggestions. 
