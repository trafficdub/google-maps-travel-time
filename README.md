# google-maps-travel-time
Google Sheets utility to fetch real-time travel times data using Google Directions API.
## Retrieve Real-Time Travel Times Data with Google Sheets
Link to [Example Google Sheets](https://docs.google.com/spreadsheets/d/1_bfzHTEkwcLAnlFnZ1TTojtqJFJfM9ez62ZRmvHAwss/edit?usp=sharing)
### Instructions
#### Create a copy with your own Google Sheets
- Create a blank Google Sheets file
- Copy the sheet content from the example above
- Go to Tools >> Script Editor, and copy the javascript in google_sheets_script.js from this repository to Main.gs
- Right click on the Run button, select the 3 dots icon, then select "Assign Script"
- In the Assign Scrip window, paste in "writeCurrentVehTT" (the name of the main javascript function)
#### Enter Required Inputs
- Google Directions API key: Cell B17
[link to get API](https://developers.google.com/maps/documentation/directions/get-api-key)
- Route ID: Column D
- Start Location*: Column E
- End Location*: Column F

Notes:

For each route, Column D to Column F must all be filled out.

Start Location and End Location can be descriptive location names or geo-coordinates.)
#### Optional Inputs
Way Point: Column G
#### Run Script
- Run script by clicking on the Run button
- For first time use, may be prompted to give Google Sheets appropriate access to proceed
