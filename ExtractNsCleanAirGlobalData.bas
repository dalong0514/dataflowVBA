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
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

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
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  Call ExtractOneColumnData(Sheet1.Range("B4:B100"), myTxt, 150)
  Call ExtractOneColumnData(Sheet1.Range("C4:C100"), myTxt, 150)

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
  myTxt.Write vbCr
End Sub