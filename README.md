# Ranil Joshua Module 2 Challenge, VBA-challenge

New Columns were created for the following:
* Ticker = Each Unique Ticker 
* Yearly Change = '\<Closing Price at End of Year> - \<Opening Price at Start of Year>'
* Percent Change = '-(1-(\<Closing Price at End of Year> / \<Opening Price at Start of Year>))'
* Total Stock Volume = SUM(\<Volume at Start of Year> : \<Volume at End of Year>)
* Greatest Percent Increase = MAX(\<Percent Change>), and corresponding \<Ticker>
* Greatest Percent Decrease = MIN(\<Percent Change>), and corresponding \<Ticker>
* Greatest Total Volume = MAX(\<Total Stock Volume>), and corresponding \<Ticker>
	
Conditional Formatting used for the following Columns:
* Yearly Change - Highlighted positive change in green and negative change in red
* Percent Change - Formatted as Percent
	
 VBA script loops over all sheets in the Worksheet

 NOTE:
One piece of code that was found online was the code for finding the last row in a column of data:
* 'Following solution for automatically finding last Row with data was found here: 'https://www.wallstreetmojo.com/vba-last-row/
* lastRow = Cells(Rows.Count, 1).End(xlUp).Row
