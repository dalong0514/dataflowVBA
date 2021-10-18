' 2021-10-12
Public Sub ExtractKsInstrumentLocationData()
  Dim buildingData As String

  buildingData = "D:\dataflowcad\ksdata\ksInstrumentLocationData.json"

  Call ExtractDataToJson(buildingData, Sheet1.range("B3:L3"), Sheet1.range("B5:L5000"), 5000, 11)

  MsgBox "Extract Sucess!"

End Sub


Public Sub ExtractKsBzInstrumentData()
  Dim buildingData As String

  buildingData = "D:\dataflowcad\ksdata\ksBzInstrumentData.json"
  Call ExtractDataToJson(buildingData, Sheet1.range("B3:J3"), Sheet1.range("B5:J5000"), 5000, 9)
  MsgBox "Extract Sucess!"

End Sub

' 2021-10-12
Sub ExtractDataToJson(gctFileName, keyRange, valueRange, rowNum, columnNum)
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
    If valueRange.Cells(row, 3).Value <> "" Then
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