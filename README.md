# Challenge_2_VBA

In this homework assignment I used VBA scripting to analyze generated stock market data.





Sources 

The vast majority of my code was generated based on the Week 2 course lectures 1, 2, and 3 along with support from Xpert Learning Assistant for debugging. This included the decalration of vairables, variable assignments, variable initializations, formating, looping, if statements and more. 

  Note: After facing issues enabling the script to run on every worksheet (that is, every quarter) at once, I turned to alternative sources on google    such as published Chegg projects and Chat GBT in order to ultimately implement the following code:
          Dim sheetNames As Variant
          sheetNames = Array("Q1", "Q2", "Q3", "Q4")
          ...
           For Each ws In ThisWorkbook.Sheets(sheetNames)
For whatever reason declaring the sheetNames as a Variant and assigniging it to an Array of the four sheet names (as opposed to individual strings) allowed my code to function and properly apply all the steps to each sheet. 
           
