' 2021-06-08
Public Sub ExtractNsCleanAirAllGlobalParamToCSV()
  Call ExtractNsCleanAirGlobalProjectInfoToCSV()
  Call ExtractNsCleanAirGlobalParamToCSV()
  MsgBox "Extract Sucess!"
End Sub

Sub ExtractNsCleanAirGlobalProjectInfoToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\nsCleanAirGlobalProjectInfo.txt"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True, Unicode:=True))

  myTxt.Write range("E2").Value
  myTxt.Write vbCr

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractNsCleanAirGlobalParamToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\nsCleanAirGlobalParam.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True, Unicode:=True))

  Call ExtractOneColumnData(Sheet1.Range("B4:B500"), myTxt, 500)
  myTxt.Write vbCr
  Call ExtractOneColumnData(Sheet1.Range("C4:C10"), myTxt, 62)

  Call ExtractUnitData(Sheet1.Range("C71:E89"), myTxt, 19)
  Call ExtractUnitData(Sheet1.Range("C92:E110"), myTxt, 19)
  Call ExtractUnitData(Sheet1.Range("C113:E131"), myTxt, 19)
  Call ExtractUnitData(Sheet1.Range("C134:E152"), myTxt, 19)
  Call ExtractUnitData(Sheet1.Range("C155:E173"), myTxt, 19)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractOneColumnData(range, myTxt, rowNum)
  Dim row As Integer
  row = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    if range.Cells(row, 1).Value <> "" Then 
      myTxt.Write ","
      myTxt.Write range.Cells(row, 1).Value
    End if
  Next row
End Sub

Sub ExtractUnitData(range, myTxt, rowNum)
  Dim row As Integer, column As Integer
  row = 1
  myTxt.Write ","
  For row = 1 To rowNum
    ' Only extract the row that have the data
    if range.Cells(row, 3).Value <> "" Then 
      myTxt.Write range.Cells(row, 1).Value
      myTxt.Write "#"
    End if
  Next row
End Sub