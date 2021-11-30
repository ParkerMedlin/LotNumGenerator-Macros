# Blending-Schedule
Version control for VBA macros in the blending schedule workbook.


### This workbook is used for:
 - Tracking inventory of finished chemical blends.
 - Planning batch sizes based on information from the production schedule (separate workbook).
 - Tracking raw material needs and predicting shortages before they become an issue.
 - Reporting of inventory transaction timelines for easier comparison between Sage 100 database and actual inventory counts.







*Update notes 11-30-2021*
 - updated lot number lookup macro to reflect new filepath
 - updated history report macro and daily count macro to reflect new filepaths
 - added startron planning macro
     - this will help us plan specifically for startron runs when we have large qtys of different blends all scheduled close together and we can't afford the space required to blend all of them beforehand 
   	 - This macro opens blendData, filters the list once by each different startron PN, copies each version of the list to a new sheet, then makes a table and orders it by start time.
   	 - I can't figure out why the loop I wrote to clear empty rows + the header rows where they copy over, isn't clearing out the very bottom header row in the table. (i am copying over the entire table including headers each time I re-filter the blendData table). I am using Rows(incrementorFromMyFORLoop).EntireRow.Delete to clear out empty rows and header rows. when the for loop gets to row 24, the function deletes row 25.
   	 - ^Leaving as-is because the report is functioning like i need it to. This is a cosmetic issue.
