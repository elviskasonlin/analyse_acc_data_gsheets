# Google Sheets Add-On for Analysing Acceleration Data

This repository is created as part of the 10.015 Physical World 1D Project (SSP)

For a python implementation for acceleration data analysis, click ![here](https://github.com/seancze/analyse_acc_data_10.015). This will lead you to the "analyse_acc_data_10.015" repository.

___

# How to use this script

## Method 1

### 1. Cleaning up spreadsheets
* Ensure that the spreadsheets are all broken up into their own separate files based on their trips (Remember to leave in the data for where the train is at a stop) 
* Write down the rows for which the train is stationary (at a stop). We will need this later
* Additionally, write down the column letters for the following:
  * Timestamp (UNIX Time)
  * Sample Count
  * Acceleration X-axis
  * Acceleration Y-axis
  * Acceleration Z-axis

### 2. Creating a new script file in google drive
* In google drive, click "New > More > Google Apps Script"
* Name the Google Apps Script

### 3. Copy and paste the contents in 1d-script.js into the new script file
* Click on "1d-script.js" in this repository to open it
* Next, click "raw"
* Select the contents of "1d-script.js" file and copy
* Paste it in the newly created Google Apps Script file

### 4. Enable it for use in the spreadsheets
* Open the newly created Google Apps Script file
* Click "Run" in the menu bar
* Select "Test as add-on"
* In the modal window that shows up, go to "Configure New Test"
* Under "Select Version", select "Test with latest code"
* Under "Installation Config", select "Installed and Enabled"
* Next, click "Select Doc" to find and select the spreadsheet you wish to use this add-on in
* Select the spreadsheet and click "select"
* Finish enabling its use by clicking "Save"
* Repeat the steps to enable it for other spreadsheets

### 5. Use them in identified spreadsheets
* In the spreadsheets with the add-on enabled, select "Add-ons" from the menu bar
* In the dropdown menu, select the name of this add-on "1d-project"
* Next, set the variables for the script to work by following the modals
* Select "Process Acc Variable" to start the script
