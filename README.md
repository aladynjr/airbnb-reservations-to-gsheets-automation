# Airbnb Reservations Google Sheets Integration

This Google Apps Script project automates the process of importing Airbnb reservation data into a Google Sheets document. It provides a simple, one-click solution for hosts to keep their reservation information up-to-date.

## Features

- Fetches reservation data from Airbnb using their API
- Automatically updates a Google Sheet with new, modified, and canceled reservations
- Preserves custom data in additional columns
- Adds a custom menu to the Google Sheet for easy access
- Includes error handling and user notifications

## Sheet Example

Here's an example of how the sheet looks with data:

<img src="images/sheet_example.png" alt="Example of the Airbnb Reservations sheet with data" width="800">

## Setup

1. Create a new Google Sheet
2. Open the Script Editor (Tools > Script editor)
3. Copy the provided code into the Script Editor
4. Save the project
5. Refresh your Google Sheet to see the new "🏠 Airbnb Reservations" menu

## Configuration

Before using the script, you need to set up the configuration:

1. Open the Google Sheet
2. Look for the "config" sheet
3. Enter your Airbnb cookie value in cell B2
4. The API key in cell B3 should already be set

## Usage

1. Click on the "🏠 Airbnb Reservations" menu
2. Select "📅 Update Reservations"
3. The script will fetch the latest reservation data and update the sheet

## Sheet Structure

- Columns A to M: Reservation data from Airbnb
- Column N: City Tax (calculated)
- Columns O to Q: Checkboxes for Checked In, Checked Out, and Cleaned

## Notes

- The script adds dummy data for demonstration purposes. Remove or comment out the `addDummyDataToCSV` function call in production.
- Ensure you have the necessary permissions to access the Airbnb API with your account.
- This project was created to meet the specific needs of a particular client. You might need to customize it to fit your own requirements or use case.

## Disclaimer

This project is not officially associated with Airbnb. Use it at your own risk and ensure you comply with Airbnb's terms of service.