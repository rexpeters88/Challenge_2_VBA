# Challenge_2_VBA

In this homework assignment I used VBA scripting to analyze generated stock market data.



Sources 

The vast majority of my code was generated based on the Week 2 course lectures 1, 2, and 3, along with knowledge from classes such as AP Computer Science and Intro to Data Science II at University of California, Santa Barbara. This included the declaration of variables, variable assignments/initializations, formating, iterations, and more.  

   Note: After facing issues enabling the script to run on every worksheet at once (loop through each quarter), I turned to alternative sources such as the Xpert Learning Assistant and similar published Chegg projects in order to ultimately implement the following code:
  
          Dim sheetNames As Variant
          sheetNames = Array("Q1", "Q2", "Q3", "Q4")
          ...
           For Each ws In ThisWorkbook.Sheets(sheetNames)
           
Declaring the sheetNames as a 'Variant' and assigning it to an Array of the four sheet names (as opposed to individual strings) allowed my code to function and properly across all the sheets. 
           
