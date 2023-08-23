# Youtube Analytics Data to Excel Sheet  

## A tool that will convert YouTube data to an excel sheet and allows you to keep and remove each row of data manually.  

This tool requires the user to have a Google Client Secret Key from Google Cloud Services API. After inputing the client secret and channel data, the tool will extract the data
into a excel sheet. After converting it into an excel sheet, the user can remove rows from the excel sheet without manually deleting rows inside the excel sheet.  

## How to install and use the tool  
1. Clone this project
2. Go to the project - `cd yt-data-to-excel`
3. Run the main program - `python main.py`
4. Input the channel secret key, channel id, and data to scrape to
5. To remove rows run the gui - `python improvedGUI.py`
6. Add the data excel sheet in the first file
7. Add the output excel sheet in the second file
8. Keep the row with the "Keep" row button and "Next" to remove the row

