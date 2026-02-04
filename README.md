# VBA Coding Challenge Example #1

## Instructions

Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock. The result should match the following image:

<img width="718" alt="moderate_solution" src="https://github.com/JesseOli100/VBA-challenge/assets/62526904/9b8216f0-b43b-4a67-8e66-e3921ee90bba">

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

<img width="1038" alt="hard_solution" src="https://github.com/JesseOli100/VBA-challenge/assets/62526904/9cbfb848-fb1c-4695-9e75-11a5526448bd">

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

## NOTE

Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

## Other Considerations

Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

## Submission

All three of the following are uploaded to GitHub/GitLab:

Screenshots of the results (5 points)

Separate VBA script files (5 points)

README file (5 points)

## Sources

•	Help from TA Matthew Werth 

•	Met up w/ other students during study groups to hash out the code through what we have learned so far

•	Used StackOverFlow for issues on the code and/or to explain why certain pieces of the script were not running as intended

•	The majority of how the code works is explained on the .bas file through the comments made next to the code. I did my best to iterate each line with the basics precepts of what each line or set of lines is supposed to be doing

## Notes

What is each set of code supposed to be doing?

The script of code is supposed to loop through all the stocks for one year and it is supposed to output the ticker symbol, yearly change from the opening price to the closing price of the whole year for each ticker, total stock volume. 

How do the formulas work the way they do?

The entirety of the code was written in a single sub. It mainly operates through the iteration of a loop which then gathers the individual ticker (and discards repeats of the same ticker), then gathers the yearly change for each ticker, and finalizes with the stock volume. You technically could’ve also just turned each part of code into a variable and then just called 3 different variables within the sub while the functions ran in a module instead of the worksheet itself. 

Either way, the main point of this exercise was really to get a feel for how loops work and how you can use them to automate an action that would be very time-exhaustive if it had to be done manually. 

Syntax Learned:

Loop syntax 

For x = Variable_1 (a number in this example) To Variable_2

If Variable or Formula Then
		
  code needed
  
End if 

Next x

Interesting Formula for Volume

ws.Cells(Ticker_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(k, 7), ws.Cells(g, 7)))

ws.Cells(Ticker_Count, 12).Value: This line is assigning a value to a specific cell in the worksheet (ws). The cell is identified by its row (Ticker_Count) and column (12). The value being assigned is the result of the subsequent calculation.

WorksheetFunction.Sum(Range(ws.Cells(k, 7), ws.Cells(g, 7))): This part of the code calculates the sum of a range of cells.

ws.Cells(k, 7): This specifies a cell in the worksheet (ws) at the intersection of the row k and column 7.
ws.Cells(g, 7): This specifies another cell in the worksheet at the intersection of the row g and column 7.
Range(...) is used to define a range of cells between the two specified cells.
WorksheetFunction.Sum(...) calculates the sum of the values within the specified range.
The calculated sum is then assigned to the cell specified in the first part of the code (ws.Cells(Ticker_Count, 12).Value).

This code is summing the values in a specified column (column 7) from row k to row g and then storing that sum in a specific cell in column 12 (the value of Ticker_Count row) of the same worksheet.

- - -

This is submitted by Jesse Olivarez for the University of Utah: Data Analytics Bootcamp

