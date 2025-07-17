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

<img width="2882" height="1566" alt="image" src="https://github.com/user-attachments/assets/fa3454b0-6917-4e20-b9a1-5054ea9a78c2" />

### Step 2: Insert a New Module
A macro needs a place to live. We'll put it in a new "Module."

1.  In the VBA editor's top menu, click **Insert**.
2.  From the dropdown menu, click **Module**.

A new blank white code window will appear on the right.

<img width="2875" height="1566" alt="image" src="https://github.com/user-attachments/assets/6f62d752-94e5-4fb1-bfbe-2780afa9955e" />

### Step 3: Copy and Paste the Code
1.  Click here to view the code file: [`SaveAllSheetsAsCSV.bas`](./SaveAllSheetsAsCSV.bas)
2.  On the page that opens, click the "Copy raw contents" button (it looks like two overlapping squares) in the top-right corner of the code box.
3.  Return to your Excel VBA editor and paste the copied code into the blank module window.

<img width="2532" height="1260" alt="image" src="https://github.com/user-attachments/assets/3edc273e-c9ba-4c19-92ab-1f07e768ff2c" />

<img width="2879" height="1569" alt="image" src="https://github.com/user-attachments/assets/6f604db4-3483-4cdf-bf82-2b0740b8bba6" />

### Step 4: Save Your Workbook as Macro-Enabled
This is a crucial step! Regular `.xlsx` files cannot store macros.

1.  Go to **File -> Save As**.
2.  In the "Save as type" dropdown menu, choose **Excel Macro-Enabled Workbook (*.xlsm)**.
3.  Save the file.

<img width="2895" height="1845" alt="image" src="https://github.com/user-attachments/assets/a7df1562-1859-4879-aecb-e35af0741bb4" />

<img width="2043" height="1176" alt="image" src="https://github.com/user-attachments/assets/3dc8b7f5-64d7-434e-a2e0-bbe86ab637ad" />


### Step 5: Run the Macro!
Now for the easy part.

1.  Press the keyboard shortcut **Alt + F8**. This opens the Macro dialog box.
2.  Click on `SaveAllSheetsAsCSV` in the list to select it.
3.  **Important:** If you don't see the macro, check the **"Macros in:"** dropdown at the bottom. It should be set to **"This Workbook"**.

	*   **Why?** This dropdown filters which macros you see. Here's what the options mean:
		*   `This Workbook`: Shows only macros saved in the current file. **(Choose this one!)**
		*   `All Open Workbooks`: Shows macros from every Excel file you have open.
		*   `Personal.xlsm`: Shows macros from your personal, global macro file (if you have one).

4.  Click the **Run** button.

<img width="581" height="865" alt="image" src="https://github.com/user-attachments/assets/97290c15-47b3-4799-8cbc-2653f4c92bcb" />


That's it! A message box will appear when the process is complete, and you'll find all your new CSV files waiting for you in the same folder as your Excel file.

<img width="633" height="525" alt="image" src="https://github.com/user-attachments/assets/19d4867e-b8af-4b35-879f-143c17340c38" />

## License
This project is licensed under the MIT License. Feel free to use and modify it as you wish.
