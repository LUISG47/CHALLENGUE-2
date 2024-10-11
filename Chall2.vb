
Sub stocks()
  ' Create a variable to hold the counters and WS var
Dim i As Long
Dim j As Long
Dim k As Long
Dim r As Long
Dim c As Long
Dim ws As Worksheet

    For Each ws In Worksheets
              ' Set all the variables
              
              Dim Ticker As String
              Dim LastRow As Long
              Dim maxValueV As Double
              Dim maxValueP As Double
              Dim minValueP As Double
              Dim tickerinf As String
             
              Dim Total_Stock_V As Double
              Total_Stock_V = 0

              Dim QChangue As Double
              QChangue = 0

              Dim InitialQ As Double
              InitialQ = ws.Cells(2, 3)

              Dim Percentage As Double
              Percentage = 0
            
              Dim Summary_Table_Row As Long
              Summary_Table_Row = 2

              ' Determine the Last Row 
              LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
              
              ' Put the desired headers on the worksheets
              ws.Range("J1") = "Ticker"
              ws.Range("K1") = "Quarterly Changue"
              ws.Range("L1") = "Percent Changue"
              ws.Range("M1") = "Total Stock Volume"
              ws.Range("Q1") = "Ticker"
              ws.Range("R1") = "Value"
              ws.Range("P2") = "Greatest % increase"
              ws.Range("P3") = "Greatest % Decrease"
              ws.Range("P4") = "Greatest Total Volume"

            ' -------------------------------------------------------------
            ' CICLE TO CONSTRUCT THE DESIRED TABLE SUMMARY OUTPUTS
            ' -------------------------------------------------------------
              
              ' Loop through all tickers
                For i = 2 To LastRow

                ' Check if we are still within the same ticker
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                            ' Set the Ticker name
                            Ticker = ws.Cells(i, 1).Value
            
                            ' Add to the Total Stock Volume
                            Total_Stock_V = Total_Stock_V + ws.Cells(i, 7).Value

                            ' Calculate Quarterly Changue
                            QChangue = ws.Cells(i, 6).Value - InitialQ

                            ' Calculate the Percentage
                            Percentage = (ws.Cells(i, 6).Value - InitialQ) / InitialQ
            
                            ' Print the Ticker Name the Summary Table
                            ws.Range("J" & Summary_Table_Row).Value = Ticker
            
                            ' Print the Total Stock Volume to the Summary Table
                            ws.Range("M" & Summary_Table_Row).Value = Total_Stock_V

                             ' Print the quarterly changue
                            ws.Range("K" & Summary_Table_Row).Value = QChangue

                             ' Print the Percentage changue
                            ws.Range("L" & Summary_Table_Row).Value = Percentage
                            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                                
                            ' Add one to the summary table row
                            Summary_Table_Row = Summary_Table_Row + 1
            
                            ' Reset the Totals
                            Total_Stock_V = 0
                            InitialQ = ws.Cells(i + 1, 3)
                          
                ' If the cell immediately following a row is the same ticker
                        Else
                        ' Add to the Total Volume
                        Total_Stock_V = Total_Stock_V + ws.Cells(i, 7).Value
            
                    End If
                Next i
                

                ' -----------------------------------------------------------------
                ' CICLE TO COLOR THE QUARTERLY CHANGUE COLUMN POSITIVE OR NEGATIVE
                ' -----------------------------------------------------------------
                For j = 2 To LastRow
                    If ws.Cells(j, 11).Value > 0 Then
                        ws.Cells(j, 11).Interior.ColorIndex = 4
                    ElseIf ws.Cells(j, 11).Value < 0 Then
                        ws.Cells(j, 11).Interior.ColorIndex = 3
                    End If
                Next j

                ' -------------------------------------------------------------
                ' GET ADITIONAL SUMMARY INFO OF MAX AND MIN VALUES
                ' -------------------------------------------------------------

                maxValueV = Application.WorksheetFunction.Max(ws.Range("M1:M" & LastRow))
                    ws.Range("R4").Value = maxValueV

                maxValueP = Application.WorksheetFunction.Max(ws.Range("L1:L" & LastRow))
                    ws.Range("R2").Value = maxValueP
                    ws.Range("R2").NumberFormat = "0.00%"
                
                minValueP = Application.WorksheetFunction.Min(ws.Range("L1:L" & LastRow))
                    ws.Range("R3").Value = minValueP
                    ws.Range("R3").NumberFormat = "0.00%"


                ' -------------------------------------------------------------
                ' LOOPS TO LOCATE THE TICKER NAME OF MAX AND MIN VALUES
                ' -------------------------------------------------------------
    
                For k = 2 To LastRow
                    If ws.Cells(k, 12).Value = maxValueP Then
                        tickerinf = ws.Cells(k, 10).Value
                        ws.Range("Q2").Value = tickerinf
                    End If
                Next k

                 For r = 2 To LastRow
                    If ws.Cells(r, 12).Value = minValueP Then
                        tickerinf = ws.Cells(r, 10).Value
                        ws.Range("Q3").Value = tickerinf
                    End If
                Next r

                 For c = 2 To LastRow
                    If ws.Cells(c, 13).Value = maxValueV Then
                        tickerinf = ws.Cells(c, 10).Value
                        ws.Range("Q4").Value = tickerinf
                    End If
                Next c
                    
                
        ws.Columns("A:T").AutoFit

    Next ws

End Sub
