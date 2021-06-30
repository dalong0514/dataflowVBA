' 2021-06-01
Public Sub ExtractNsCleanAirAllDataToCSV()
  Call ExtractNsCleanAirProjectDataToCSV()
  Call ExtractNsCleanAirSupplyDataToCSV()
  MsgBox "Extract Sucess!"
End Sub

Sub ExtractNsCleanAirProjectDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\nsCleanAirProjectInfo.txt"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True, Unicode:=True))

  myTxt.Write range("C2").Value + "," + range("F2").Value
  myTxt.Write vbCr

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractNsCleanAirSupplyDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\nsCleanAirSupply.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True, Unicode:=True))

  Call ExtractOneRowData(Sheet1.Range("B3:AV3"), myTxt)
  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractColumnsData(Sheet1.Range("B5:AV300"), 47, myTxt)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractColumnsData(range, columnNum, myTxt)
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

Sub ExtractOneRowData(range, myTxt)
  Dim column As Integer
  column = 1
  Do While range.Cells(1, column).Value <> ""
    myTxt.Write ","
    myTxt.Write range.Cells(1, column).Value
    column = column + 1
  Loop
  myTxt.Write vbCr
End Sub