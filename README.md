# Excel Macro: Save All Sheets as Separate CSVs 💾

Tired of manually saving each sheet of your Excel workbook as a CSV? This simple VBA macro automates the entire process in one click!

It will loop through every visible worksheet in your workbook and save it as a separate `.csv` file in the same folder as your main Excel file.

## Is This Safe? Absolutely!

*   ✅ **Your original Excel file is never changed.** The macro only reads your sheets and saves copies.
*   ✅ **No data is sent anywhere.** This all happens locally on your computer.
*   ✅ **You can see the full code.** The entire script is available in the `SaveAllSheetsAsCSV.bas` file for you to review.

## Features
*   **One-Click Operation**: Run the macro and you're done.
*   **Automated Naming**: Each CSV is automatically named after its worksheet.
*   **Skips Hidden Sheets**: The macro is smart enough to ignore any hidden sheets.

## How to Use This Macro (Beginner's Guide)
Follow these simple steps to add this macro to your Excel workbook.

### Step 1: Open the VBA Editor
In your open Excel file, press the keyboard shortcut **Alt + F11**. This will open the Visual Basic for Applications (VBA) editor window. It's a separate window from your spreadsheet.

### Step 2: Insert a New Module
A macro needs a place to live. We'll put it in a new "Module."

1.  In the VBA editor's top menu, click **Insert**.
2.  From the dropdown menu, click **Module**.

A new blank white code window will appear on the right.

### Step 3: Copy and Paste the Code
1.  Click here to view the code file: [`SaveAllSheetsAsCSV.bas`](./SaveAllSheetsAsCSV.bas)
2.  On the page that opens, click the "Copy raw contents" button (it looks like two overlapping squares) in the top-right corner of the code box.
3.  Return to your Excel VBA editor and paste the copied code into the blank module window.

### Step 4: Save Your Workbook as Macro-Enabled
This is a crucial step! Regular `.xlsx` files cannot store macros.

1.  Go to **File -> Save As**.
2.  In the "Save as type" dropdown menu, choose **Excel Macro-Enabled Workbook (*.xlsm)**.
3.  Save the file.

### Step 5: Run the Macro!
Now for the easy part.

1.  Press the keyboard shortcut **Alt + F8**. This opens the Macro dialog box.
2.  You will see `SaveAllSheetsAsCSV` in the list.
3.  Click on it and then click the **Run** button.

That's it! A message box will appear when the process is complete, and you'll find all your new CSV files waiting for you in the same folder as your Excel file.

## License
This project is licensed under the MIT License. Feel free to use and modify it as you wish.
