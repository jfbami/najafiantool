# Log a RASS folder

Install this sidebar in the lab **CD Archive Log - With Calibration** spreadsheet. The TEM log pages provide the calibration evidence. The sidebar records Brianna's entry in the archive spreadsheet.

## For Brianna

1. In the archive spreadsheet, open **RASS archive → Log a Drive folder**.
2. Paste the link to one case/block folder and click **Load folder**. Check the displayed path and example filenames. Confirm the CD number and case/block label.
3. Select the taps you confirmed in the TEM log and enter the year and page in **TEM log reference**. Leave the taps blank if you haven't confirmed them.
4. Click **Save to archive log**. The sidebar tells you the row number. If the row has different taps, review the existing entry before changing it.

Each submission records one case/block. For a matching block row, the sidebar can fill a blank Calibration cell and add the source reference and folder URL to Comments. It leaves image-level entries alone. Folder inspection uses Drive metadata and lists images directly inside the selected folder.

## One-time setup for a spreadsheet editor

1. Open the archive spreadsheet and select **Extensions → Apps Script**.
2. Copy [Code.gs](Code.gs) into the project and add [Intake.html](Intake.html) as an HTML file named `Intake`. If the project already has an `onOpen` function, merge the menu setup into it.
3. In **Project Settings**, enable **Show “appsscript.json” manifest file in editor**. Copy [appsscript.json](appsscript.json) into a new project's manifest. If the project already has other code, merge the Drive v3 dependency and the required scopes into its existing manifest.
4. Save and reload the spreadsheet. Open the sidebar from **RASS archive**. Brianna needs edit access to the sheet and access to each folder she logs. Google may ask her to authorize the script the first time she uses it.
