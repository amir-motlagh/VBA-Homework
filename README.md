# VBA-Homework

1) At first I defined my dimensions using 'Dim'operator.

	a) 'Ws' and 'Wb' were defined as worksheets and workbooks.
	b) Then defined 'Ticker' for ticker names, 'Ticker_Total' for total amount of tickers, 'Open_Price' for beginning of the year price,
	  'Close_Price' for end of the eyar price, 'Yearly_Change' for changes in price over the year, 'Yearly_Change_Percentage' for percentage of changes,
	  'Min_Ticker' as smallest amount ticker, 'Max_Ticker' as the biggest amount ticker, ''Min/Max_Percentage' as smallest and greates percentage changes.
	  'Max_Ticker_Volume' as greatest total volume
	c)'Summary_Table_Row' is where I show the results
	d) Then Last row is defined as 'LastRow'

2)I started with the cell where I could get the Opening price. 'Open_Price = Ws.Cells(2, 3).Value'

3) Then I started a loop for commands to execute. 'For i = 2 To LastRow'

   If the ticker in a row is different from the last one, then the command goes through the loop.
   The yearly change is claculated  and if the resut is not zero the percentage is calculated.

4) Then the toal ticker volume from 7th column is added to the initial ticker volume which is zero(0) to keep
   count of the ticker volume. The ticker name and the ticker volume are added to the summary table: 'Ws.Range("I" & Summary_Table_Row).Value = Ticker' and 
   'Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change'. If command is ended.
5) Then, if yearly change is more than Zero, the cell color turns green, If it less than or equal to zero, the cell turns red.

6) Yearly_Change_Percentage and Ticker_Total are then added to columns 'K' and 'L'.

7) Then one row is added to the summary table.

8) To do the bonus part, the code comapres the Yearly_Change_Percentage is compared with Max_Percentage. If Yealy_Change_Percentage is bigger, then 
   it remains the maximum pecentage. If it is smaller then the Max_Perentage becomes the maximum yearly change.

9) Next thing is, Ticker_Total. I have applied the same logic for here too.

10) Both Yearly_Change_Percentage and Ticker_Total are initially zero to add up from there.

11)Finally, the loop goes to the next i.

12) Then the results are printed in the summary table.

13) Next worksheet starts with the same code series.

14) End of the code.

