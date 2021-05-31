' 2021-05-31
Public Sub ExtractNsCleanAirSupplyDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\bsdata\nsCleanAirSupply.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Extract data
  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractColumnsData(Sheet1.Range("A2:E100"), 5, myTxt)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub

Sub ExtractColumnsData(range, columnNum, dataTypeString, myTxt)
  Dim row As Integer, column As Integer
  row = 1
  column = 1
  Do While range.Cells(row, 1).Value <> ""
    For column = 1 To columnNum
      myTxt.Write ","
      myTxt.Write range.Cells(row, column).Value
    Next column
    myTxt.Write vbCr
    row = row + 1
  Loop
End Sub
