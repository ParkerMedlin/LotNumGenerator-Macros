# Update Notes

### 1-26-2022
 - Light edits to the increment sub to eliminate formulas in the labelqty and blendsheetqty columns

### 1-25-2022
 - Resolved issue where the issue sheet multiprint was triggering for drum batch of same blend as Hx batches below it. Just made criteria more specific so now it checks for line as well as blend and run date.  

### 1-24-2022
 - Added the copyLast sub to the plus sign instead of having a separate button
 - Added logic to the blendsheet printing macro so that it keeps jumping to next lot and printing blend sheet for that one too as long as the line and the blend desc are matching 
 - Added similar logic to the issue sheet printing macro as well 
 
### 1-18-2022
 - Made it so the format options from the line selector UserForm will put in the borders. Also moved the workbook saving to post-incrementor instead of right in the middle
 - Updated the label printer so it goes back to the lot num gen workbook after printing labels

### 1-17-2022
 - Tremendously simplified the incrementor macro. 
 - **CHANGED THE SEED NUMBER AND LOT NUMBER PROCESS SO NOW SEED NUMBER IS EQUAL TO THE ENDING OF THE LOT NUMBER, RATHER THAN TWO MORE** 

### 1-11-2022
 - Various trimming down of subroutines. Also added text box for documentation of the different macro-linked icons. 

### 1-5-2022 AGAIN
 - Added a dialog box to labelPrinter sub so it doesn't automatically print the whole pre-calculated quantity of labels.

### 1-5-2022 
 - Added NumLock module so that I could re-enable NumLock after the labelPrinter Sub, which for some reason started turning off NumLock after i changed it to print the labels in bartender with Sendkeys.

### 12-21-2021
 - Updated the labelPrinter subroutine so it now uses an array instead of copypasta
 - Changed the issue sheet gen subroutine so it now points to NOT_Blending_Issue_Sheet.xlsb
 - Removed unused module "Timestamp"

### 12-20-2021
 - Renamed pullsheetgen to picksheetgen and updated content of it to reflect new table names.

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
