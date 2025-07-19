# C3_updater
📘 Excel Updater — README

This utility compares and updates Excel files based on contracts and Work Classes (WC). It processes data from Easy and Sub folders and generates updated versions accordingly.
To use the utility: 
1. Export C3 Forms from EasyPlant into "Easy" Folder divided by contract, 
2. Download Updated C3 Forms in "Sub" Folder
3. Open Utility via .exe file or open .py file.
4. Use rename buttons to unify naming system for all files.
5. Choose required Contract and SWClasses and click update. 
6. Updated files will appear in the Updated Folder. And log.csv with additional data for all errors and problems will be created in the main folder. 
7. Optionally Click "Rename updated" to rename resulting files with a following format "SWC Code"_"SWC Description"_"Contract Code".

🖱 Interface & Button Functions:

Button                     | Description
--------------------------|------------------------------------------------------------
Browse                    | Selects folders for Easy, Sub, and Updated
Rename Easy               | Renames files in Easy to `WC_CONTRACT_easy.xlsx`
Rename Sub                | Renames files in Sub to `WC_CONTRACT_sub.xlsx`
Rename Updated            | Renames files in Updated to "WC_code"_WC_Description_Contract using names from `WorkClass_Data.csv` to rename for construction convinience 
Open Easy                 | Opens "Easy" folder for selected Contracts
Open Sub                  | Opens "Sub" folder for selected Contracts
Open Updated              | Opens "Updated" folder for selected Contracts
Select Contracts / WC     | Multi-select lists to choose specific contracts and WCs
Update                    | Compares and updates progress data from C3 in steps percentage Sub → Easy, saves result to Updated

🎨 Formatting & Highlights:

Visual Indicator               | Meaning
------------------------------|------------------------------------------------------------
Yellow cell background     | Value was decreased
Bold red font              | Value was increased
Orange background          | Data type mismatch or text in numeric column
⚠ Missing tags            | Tags in Easy not found in Sub — affected cells set to 0
❗ Duplicates              | Duplicate tags found within the same sheet

📄 Additional Info:

• If folders Easy, Sub and Updated are located in the same directory as .exe it will be chosen by default
• When issues are found (duplicates or missing tags), buttons appear to open those files.

• WorkClass_Data.csv` and "WorkStep_Data.csv" must be placed in the same folder as the executable or script.
• Fully portable — no need to install Python when using the .exe version.
