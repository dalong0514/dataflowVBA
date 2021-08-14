' refactored at 2021-08-14
Public Sub ExtractAllZsBuildingData()
  Dim buildingData, technicalEconomyData, designExplainData As String

  buildingData = "D:\dataflowcad\zsdata\zsBuildingData.txt"
  technicalEconomyData = "D:\dataflowcad\zsdata\zsTechnicalEconomyData.txt"
  designExplainData = "D:\dataflowcad\zsdata\zsDesignExplainData.txt"

  Call ExtractZsZPDataToCSV(buildingData, Sheet1.Range("B5:J500"), 200, 9)
  Call ExtractZsZPDataToCSV(technicalEconomyData, Sheet1.Range("L5:O500"), 200, 5)
  Call ExtractZsZPDataToCSV(designExplainData, Sheet2.Range("A3:B30"), 20, 2)

  MsgBox "Extract Sucess!"

End Sub

' refactored at 2021-08-14
Sub ExtractZsZPDataToCSV(gctFileName, range, rowNum, columnNum)
  Dim fso As Object
  Dim myTxt As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=gctFileName, OverWrite:=True)
  Call ExtractRangeNoNullData(range, rowNum, columnNum, myTxt)
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractRangeNoNullData(range, rowNum, columnNum, myTxt)
  Dim row As Integer, column As Integer
  row = 1
  column = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    ' refactored at 2021-08-13
    if range.Cells(row, 2).Value <> "" Then 
      For column = 1 To columnNum
        myTxt.Write ","
        myTxt.Write range.Cells(row, column).Value
      Next column
      myTxt.Write vbCr
    End if
  Next row
End Sub