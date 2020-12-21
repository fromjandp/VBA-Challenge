Attribute VB_Name = "Module1"
Sub Wallstreet_Stock_Information():

'   This variable will hold the grand totals for the Stock volume per Stock

    Dim Total_Stock_Volume As LongLong
   
'   Initialize fields for use within the code for calculations and display in the spreadsheet.

    Dim Ticker As String
    Dim Opening_Stock_Value As Double
    Dim Closing_Stock_Value As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    Dim SummaryTable_Location As Long
   
'   execute this True/False statement once, per sheet to place the new column titles into each worksheet
    Dim New_Column_Titles As Boolean
    New_Column_Titles = False

'   Execute this True/False statement once, per stock symbol, to grab the opening stock price
    Dim Got_Opening_Stock_Value As Boolean
    Got_Opening_Stock_Value = False
   
        
    For Each ws In Worksheets
    
    
'    Total_Stock_Volume Set up a Summary Table and counter to keep track of filling in the Stock summary table.
'   Set the summary table to 2 since the first value will be written to the appropriate column in the second row.

    SummaryTable_Location = 2
                
'       Add in the column Titles for 4 new columns on each worksheet

           ws.Range("I1").Value = "Ticker"
           ws.Range("J1").Value = "Yearly Change ($)"
           ws.Range("K1").Value = "Percent Change"
           ws.Range("L1").Value = "Total Stock Volume"
                  
'   This will identify the last row of data in the first column. This routine was shared in a Saturday class with Trevor Taylor
 
       Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'      For Loop to read through the first column and identify when the Stock Symbol changes
'      Starting at 2 because there is a header in the first row, first column (A1)
            
       For i = 2 To Lastrow
         
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     '          Ticker = ws.Cells(i, 1).Value
               Closing_Stock_Value = ws.Cells(i, 6).Value
               Yearly_Change = Closing_Stock_Value - Opening_Stock_Value
               If Yearly_Change > 0 Then
                  ws.Range("J" & SummaryTable_Location).Interior.Color = vbGreen
               ElseIf Yearly_Change < 0 Then
                  ws.Range("J" & SummaryTable_Location).Interior.Color = vbRed
               End If
       
'  Calculate the Percent Change from the opening stock price. Consideration given to avoid a Divide by 0 error.
'  Also rounded answer to 2 decimal
  
               If (Yearly_Change = 0 Or Opening_Stock_Value = 0) Then
                    Percent_Change = 0
               Else
                   Percent_Change = (Yearly_Change / Opening_Stock_Value)
               End If
                                               
' Display the output in Columns I through L
' Formatting functionality was taken from SlackOverflow:   https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage

               Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
               ws.Range("I" & SummaryTable_Location).Value = Ticker
                           
               ws.Range("J" & SummaryTable_Location).Value = Yearly_Change
               ws.Range("J" & SummaryTable_Location).NumberFormat = "0.00"
               
               ws.Range("K" & SummaryTable_Location).Value = Percent_Change
               ws.Range("K" & SummaryTable_Location).NumberFormat = "0.00%"
               
               ws.Range("L" & SummaryTable_Location).Value = Total_Stock_Volume
    
' Reset the Total_Stock_Volume for the next Stock

               Got_Opening_Stock_Value = False
               Closing_Stock_Value = 0
               Yearly_Change = 0
               Percent_Change = 0
               Percent_Change_Rounded = 0
               Total_Stock_Volume = 0
               
' Increment the Summary location for where the next line of summarised Stock information will be written
               
               SummaryTable_Location = SummaryTable_Location + 1
           
            Else
 '             The next record was from the same Stock so continue to process the same ticker
               Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
' Execute grabbing the Opening stock price once per Stock Symbol
               If Got_Opening_Stock_Value = False Then
                  Ticker = ws.Cells(i, 1).Value
                  Opening_Stock_Value = ws.Cells(i, 3)
                  Got_Opening_Stock_Value = True
               End If
             End If
        Next i
        New_Column_Titles = False
        ws.Columns("I:L").AutoFit
       
   Next ws
       
 End Sub

   
