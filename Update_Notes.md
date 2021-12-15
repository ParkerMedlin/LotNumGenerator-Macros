# Update Notes

<<<<<<< HEAD
### 12-15-2021
 - Made the issue sheet generator function faster by a factor of approximately 1000 by using an array instead of copypasting back & forth between windows. Also eliminated need for userform dictating to select line. 

### 12-13-2021 
 - Added incrementor functionality to the CopyLast() function 
	 - Issue arose when Shawn went to make multiple copies of a blend without first using the plus sign to make the first lot number. JRD pointed out that it probably makes more sense to change the program. 
 - Added a date and weekday inputbox to the increment function

### 12-1-2021
 - JRD pointed out a pitfall that I was running into with the empty row cleanup loop on startron report macro; resolved and pushed
     - rows would eventually be skipped because I was using the loop incrementor to store what row i was on, and number of the row I was on would change every time i deleted a row. once I added a separate `Integer` to track the row count, everything worked as intended. 

### 11-30-2021
 - updated lot number lookup macro to reflect new filepath
 - updated history report macro and daily count macro to reflect new filepaths
 - added startron planning macro
     - this will help us plan specifically for startron runs when we have large qtys of different blends all scheduled close together and we can't afford the space required to blend all of them beforehand 
   	 - This macro opens blendData, filters the list once by each different startron PN, copies each version of the list to a new sheet, then makes a table and orders it by start time.
   	 - I can't figure out why the loop I wrote to clear empty rows + the header rows where they copy over, isn't clearing out the very bottom header row in the table. (i am copying over the entire table including headers each time I re-filter the blendData table). I am using Rows(incrementorFromMyFORLoop).EntireRow.Delete to clear out empty rows and header rows. when the for loop gets to row 24, the function deletes row 25.
   	 - ^Leaving as-is because the report is functioning like i need it to. This is a cosmetic issue.
=======
### 12-6-2021
 - removed the extra (unnecessary) subroutine PullShtPrint. The only thing it was doing was displaying an inputBox for the number of copies to be printed. I was informed that we only ever need one copy of this sheet, so I removed the option (you can just click it twice if you really need multiple copies). 
>>>>>>> 578b58a19f3b52e321070b1361199d81fcb9a051
