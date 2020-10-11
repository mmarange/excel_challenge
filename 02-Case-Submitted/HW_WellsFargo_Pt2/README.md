# Wells Fargo - Part II

* In this second part of the mini project, I combined all of your previous Excel sheets into one massive table on a new sheet called Combined_data.

* **HOW TO RUN SOLUTIONS**

  * To view solution and to run the VBA program, open the [wells_fargo_format_Unsolved.xlsm] and run the available macro. 
  * To view the solution without running the VBA program open the [wells_fargo_format_Solved.xlsm] 

**Method**

* I used for Loops to loop through every worksheet and selected the state contents using `ActiveWorkbook.Sheets(name_sheet).Cells(a, b)` function and subsequently pasted in the Combined_Data sheet.

* Used a combination of `for` loops and `If` function to assign $0.00 values to all blank cells in Combined_data sheet for ease of future data analysis. I assumed blank cells are equivalent to zero deposits.
