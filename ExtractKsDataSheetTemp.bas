' 2021-07-04
' refactored at 2021-07-12
Sub ExtractKsPressureSupplyDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\ksDataSheetTemp.json"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractProjectInfoToJsonString(myTxt)
  Call ExtractRangeDataToJsonString(Sheet1.range("B5:AG5"), Sheet1.range("B7:AG500"), 1000, 32, myTxt)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing

  MsgBox "Extract Sucess!"
End Sub

Sub ExtractRangeDataToJsonString(keyRange, valueRange, rowNum, columnNum, myTxt)
  Dim row As Integer, column As Integer, jsonString As String, cellString As String, arr As Variant
  arr = ExtractOneRowDataToArray(keyRange)
  row = 1
  column = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    if valueRange.Cells(row, 1).Value <> "" Then 
      myTxt.Write "{"
      jsonString = ""
      For column = 1 To columnNum
        cellString = Replace(valueRange.Cells(row, column).Value, ",", "，")
        ' the cell content may have ", convert it to #, reconvert it before insert the database 2021-07-13
        cellString = Replace(cellString, """", "#")
        cellString = Replace(cellString, ":", "：")
        jsonString = jsonString & chr(34) & arr(column-1) & chr(34) & ":" & chr(34) & cellString & chr(34) & ","
      Next column
      jsonString = Left(jsonString, Len(jsonString)-1)
      myTxt.Write jsonString
      myTxt.Write "}"
      myTxt.Write vbCr
    End if
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

Sub ExtractProjectInfoToJsonString(myTxt)
  Dim jsonString As String
  myTxt.Write "{"
  jsonString = chr(34) & Sheet1.range("B3") & chr(34) & ":" & chr(34) & Sheet1.range("C2") & chr(34) & "," & chr(34) & Sheet1.range("C3") & chr(34) & ":" & chr(34) & Sheet1.range("E2") & chr(34) & "," & ConvertString("dataClass") & ":" & ConvertString("pressure")
  myTxt.Write jsonString
  myTxt.Write "}"
  myTxt.Write vbCr
End Sub

Function ConvertString(cellString)
  Dim resultString As String
  resultString = chr(34) & cellString & chr(34)
  ConvertString = resultString
End Function