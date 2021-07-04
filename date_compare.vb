Rem // setup application reminder event upon opening .xlsm file
Sub Workbook_Open()
  Application.OnTime TimeValue("09:00:00"), Procedure:="ThisWorkbook.at_time"
End Sub

Public Sub at_time()
  Rem // setup recursive application reminder event
  Application.OnTime TimeValue("09:00:00"), Procedure:="ThisWorkbook.at_time"
  Dim int_res As Integer
  Dim TDate As String
  
  TDate = Date
  TDate_p_3 = DateAdd("d", 3, TDate) Rem // today's date plus 3 days
  
  For I = 1 To 1000 Rem // loop through 1000 rows
    Dim val As String
    val1 = Worksheets("Sheet1").Cells(I, 1).Text
    val2 = Worksheets("Sheet2").Cells(I, 2).Text
    
    Rem // StrComp() returns a 0 if the strings are equal
    If StrComp(val1, TDate_p_3) Then
      int_res = 0 Rem // do nothing to fill in code block
    Else
      MsgBox "MEMBER PAYMENTS due 3 days from now: " & TDate_p_3
    End If
    
    If StrComp(val2, TDate) Then
      int_res = 0 Rem // do nothing to fill in code block
    Else
      MsgBox "ALL KIDAZZLER PAYMENTS due today: " & TDate
    End If
    
  Next I
End Sub
