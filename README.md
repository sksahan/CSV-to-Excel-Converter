# **CSV to Excel Converter**

A simple desktop application built with Python to convert multiple CSV files in a folder into Excel (.xlsx) files. This tool is designed for users who need a quick and easy way to handle batch conversions of data files.

## **Features**

* **Batch Conversion:** Converts all .csv files within a selected folder and its subfolders.  
* **Separator Auto-Detection:** Automatically detects the correct separator (semicolon or comma) for your CSV files.  
* **Clean Output:** Automatically removes extra whitespace from column headers and drops blank, unnamed columns to ensure a clean final output.  
* **User-Friendly Interface:** A simple and intuitive GUI makes it easy for anyone to use without needing to write code.

## **Prerequisites**

To run this application, you need to have Python installed on your system.  
The following Python libraries are required:

* pandas  
* openpyxl  
* tkinter (usually included with Python)

You can install the required libraries using pip:  
pip install pandas openpyxl

## **How to Run**

1. **Clone the repository** (if available) or download the script.  
2. **Open a terminal or command prompt** in the directory where you saved the script.  
3. **Run the script** using the following command:

python your\_script\_name.py

*(Replace your\_script\_name.py with the name of your Python file)*

## **Usage**

1. **Select Folder:** Click the "Browse" button and choose the folder containing the CSV files you want to convert.  
2. **Choose Separator:** Use the dropdown menu to select your CSV separator. The "All (auto-detect)" option is recommended.  
3. **Convert:** Click the "Convert Files" button. The converted Excel files will be saved in the same folders as the original CSV files.

## **Troubleshooting**

* If the program does not run, ensure you have installed the required libraries as listed in the **Prerequisites** section.  
* Column widths are not automatically adjusted. You will need to manually adjust them in Excel for better readability if necessary.

## **Contributing**

If you would like to contribute to this project, please feel free to submit a pull request or open an issue.
