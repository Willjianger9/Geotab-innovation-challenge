# Delete Non-DOCX Files Script

This script helps you clean up the `data` directory by automatically deleting all files that are not Word documents (`.docx`). It will recursively scan through the directory and its subdirectories, identifying non-DOCX files for deletion.

## Requirements

- Python 3.x installed on your system

## How to Run

1. Make sure you have Python 3 installed on your computer
   - You can check by opening a terminal/command prompt and typing: `python3 --version`
   - If Python 3 is not installed, download and install it from [python.org](https://www.python.org/downloads/)

2. Place the `delete_non_docx.py` script in the same directory as your `data` folder (or adjust the path in the script)

3. Run the script from the terminal/command prompt:
   ```
   python3 delete_non_docx.py
   ```

4. The script will:
   - List all non-DOCX files it finds
   - Ask for your confirmation before deleting anything
   - Delete the files only after you type "yes" or "y"
   - Report how many files were successfully deleted

## Safety Features

- **Preview**: The script shows you exactly which files will be deleted before taking any action
- **Confirmation Required**: Nothing is deleted without your explicit confirmation
- **Detailed Reporting**: The script reports which files were successfully deleted and which ones failed (if any)

## Script Location

Make sure the script is located in the same directory as the `data` folder you want to clean up. If your data folder is located elsewhere, you'll need to modify the path in the script:

```python
# Change this line in the script
data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

# To point to your custom location, for example:
data_dir = '/path/to/your/data/folder'
```

## Troubleshooting

- If you get a "command not found" error, try using `python` instead of `python3`
- If you get a permission error, make sure the script has execute permissions:
  ```
  chmod +x delete_non_docx.py
  ```
  And then try running it with:
  ```
  ./delete_non_docx.py
  ```
