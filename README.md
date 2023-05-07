# YuGiOh-rarity-matrix
Creating a excel sheet from a Duelingbook or YGOPro Decklist to display all raritys of all individual cards

-----------------------------------------------------------------------------------------------------------
DESCRIPTION:
This tools use is inteded for checking the availiable card raritys of certain decklists.

The code checks the script path for existing .ydk decklists, parses the data to the ygoprodeck api 
and creates a excel file providing the availaible released raritys and price from tcgplayer.com
for each card.
-----------------------------------------------------------------------------------------------------------
HOW TO USE:
the code uses the following external libarys wich must be installed in oder to use the code
-> requests: pip install requests
-> openpyxl: pip install openpyxl

1) Place any amount of .ydk files (tested for DuelingBook and EDOPro) inside the script folder.
2) Run the script via command shell or batch file
3) A excel file should be created

*If your using EDOPro files with pre-errata card versions 
place the extensions folder from the EDOPro directory inside the script folder,
the script the checks if there is a pre-errata version and changes the id to be 
properly found in the ygoprodeck api.
-----------------------------------------------------------------------------------------------------------
For anyone whos using a android device: 
There is a app called "Pydroid 3" where you can run Python scripts.
For viewing the excel file ive found WPS Office quite comfortable, just make sure to activate word wrap
while marking all cells.
-----------------------------------------------------------------------------------------------------------
DISCLAIMER:
Im quite new to programming and did this to reduce time while building decks ive testes online in a certain 
rarity. However, since im not to experienced im glad for any code reviews or suggestions. 
-----------------------------------------------------------------------------------------------------------





