# SICONV

# Automated Data Entry Script for the Brazilian Federal Government SICONV
This Python script automates the process of entering budget items into the SICONV system. It reads data from an Excel sheet and interacts with the SICONV web interface using the pyautogui library, facilitating the data entry process by mimicking keyboard and mouse actions.

## Features
Local Administration: The script is designed to automate local administration tasks by populating fields in the SICONV system based on data from an Excel spreadsheet.

Data Extraction: Extracts information from the 'ORÇAMENTO' and 'CÁLCULO' sheets of the Excel file.

Automated Input: Automatically fills in fields such as codes, descriptions, units, and prices in the SICONV system.

Image Comparison: Utilizes image comparison to ensure the system is ready for the next input step.

Multiple Items: Capable of handling multiple items (n_frentes) and their respective quantities.

Dynamic BDI Calculation: Calculates and enters BDI (Budget Difference Index) values dynamically based on the input data.

## Requirements
```
Python 3.x
Required Libraries:
pandas
pyautogui
PIL
xerox
matplotlib
```

You can install the required libraries using:

```
pip install pandas pyautogui pillow xerox matplotlib
```

## How to Use

### Prepare the Excel File:

Ensure your Excel file is named planilha.xlsx and contains the sheets 'ORÇAMENTO' and 'CÁLCULO'.
Place the Excel file in the same directory as the script.

### Configure Script Parameters:

Adjust the linhaInicial, linhaUltima, and colunaInicial variables to match the structure of your Excel file.
Set the time0, time1, and time2 variables to control the timing of the script's operations.
Define the n_frentes variable to match the number of work fronts.
Set item_times to the number of SINAPI items.
Update bdi1 and bdi2 with the appropriate BDI values.

### Run the Script:

Ensure that the SICONV system is open and Excel is ready with the correct cell selected.
Run the script, and it will automatically perform the data entry operations.

### Image Comparison:

The script uses screenshots to check if the SICONV system is ready for the next operation. Ensure that the required reference images (aguarde_trabalho.png, lupaPesquisar.png, etc.) are present in the same directory as the script.

### Automation Control:

The script includes functions like cadastroItem and cadastroMacroItem to manage item registration and macro item registration, respectively.
Notes
The script requires the user to provide appropriate timing for slower and faster operations using the time0, time1, and time2 variables.
Ensure that the SICONV system and Excel are in the correct state before running the script to avoid any interruptions.
The script performs real-time interactions with the system, so it should be monitored during execution to handle any unexpected prompts or errors.

Author
Felipe Brandão

