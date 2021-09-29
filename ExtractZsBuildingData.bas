' refactored at 2021-08-14
' refactored at 2021-09-29
Public Sub ExtractAllZsBuildingData()
  Dim buildingData, technicalEconomyData, designExplainData As String

  buildingData = "D:\dataflowcad\zsdata\zsBuildingData.json"
  ' technicalEconomyData = "D:\dataflowcad\zsdata\zsTechnicalEconomyData.json"
  ' designExplainData = "D:\dataflowcad\zsdata\zsDesignExplainData.json"

  Call ExtractZsZPDataToJson(buildingData, Sheet1.range("B4:ZZ4"), Sheet1.range("B7:ZZ200"), 200, 20)
  ' Call ExtractZsZPDataToCSV(technicalEconomyData, Sheet1.Range("L5:O500"), 200, 5)
  ' Call ExtractZsZPDataToCSV(designExplainData, Sheet2.Range("A3:B30"), 20, 2)

  MsgBox "Extract Sucess!"

End Sub

' refactored at 2021-08-14
Sub ExtractZsZPDataToJson(gctFileName, keyRange, valueRange, rowNum, columnNum)
  Dim fso As Object
  Dim myTxt As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=gctFileName, OverWrite:=True)
  Call ExtractRangeDataToJsonString(keyRange, valueRange, rowNum, columnNum, myTxt)
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractRangeDataToJsonString(keyRange, valueRange, rowNum, columnNum, myTxt)
  Dim row As Integer, column As Integer, jsonString As String, cellString As String, arr As Variant
  arr = ExtractOneRowDataToArray(keyRange)
  row = 1
  column = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    If valueRange.Cells(row, 2).Value <> "" Then
      myTxt.Write "{"
      jsonString = ""
      For column = 1 To columnNum
      cellString = Replace(valueRange.Cells(row, column).Value, ",", ",")
        ' the cell content may have ", convert it to #, reconvert it before insert the database 2021-07-13
        cellString = Replace(cellString, """", "#")
        cellString = Replace(cellString, ":", ":")
        jsonString = jsonString & ConvertString(arr(column - 1)) & ":" & ConvertString(cellString) & ","
      Next column
      jsonString = Left(jsonString, Len(jsonString) - 1)
      myTxt.Write jsonString
      myTxt.Write "}"
      myTxt.Write vbCr
    End If
  Next row
End Sub

Function ExtractOneRowDataToArray(range)
  Dim column As Integer, csvString As String, arr As Variant
  column = 1
  csvString = ""
  Do While range.Cells(1, column).Value <> ""
    csvString = csvString & range.Cells(1, column).Value & ","
    column = column + 1
  Loop
  arr = Split(csvString, ",")
  ExtractOneRowDataToArray = arr
End Function

Function ConvertString(cellString)
  Dim resultString As String
  resultString = Chr(34) & cellString & Chr(34)
  ConvertString = resultString
End Function