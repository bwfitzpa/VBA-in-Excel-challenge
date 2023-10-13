# VBA-challenge
*Files created for challenge 2/VBA challenge

*Screenshots of the results files

 	2018_results_screenshot.jpg
  
 	2019_results_screenshot.jpg
  
 	2020_results_screenshot.jpg

*VBA script files, the code is saved as a BAS file and is also saved as part of an excel workbook

	text file name, VBA_code_challenge_2.bas
 
	excel file name, VBA_code_challenge_2.xlsm


*Code source/where I received help writing the code
	
 	-I received help from the TA and another student in the class on 
	how to use "For Each ws in Worksheets" and "ws.activate" to run the 
	code across all worksheets in the workbook
	
 	-I modified the code from the "credit card check" in class assignment 
	to use it to create the values in the ticker, volume_total, yearly_change,
	and percent_change columns. Using the following code from the assignment,
	"If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then"
	
 	-I received help through the AskBCS Learning Assistant on how to get the 
	code to loop from one stock to the next in terms of getting the correct 
	values in the yearly_change, percent_change, and total_volume columns 
	specfically when it came to what to set the open_value variable at, how 
	to set the close_value, how to reset the open_value in the loop, and then 
	how to use this to find yearly_change.
	
 	-I learned how to set the column width and how to format cells as a %
	using the macro recorder in excel and looking at the code it created
