' refactored at 2021-11-29
Public Sub ExtractAllBsGCTTankData()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\bsGCTData.json"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractProjectInfoToJsonString(myTxt)
  Call ExtractTankMainDataToJson(Sheet1.range("B6:AS6"), Sheet1.range("B7:AS200"), 200, 44, myTxt, "MainData")
  Call ExtractTankMainDataToJson(Sheet2.range("B4:M4"), Sheet2.range("B5:M2000"), 2000, 12, myTxt, "NozzleData")
  Call ExtractTankMainDataToJson(Sheet3.range("B4:G4"), Sheet3.range("B5:G200"), 200, 6, myTxt, "PressureElementData")
  Call ExtractTankMainDataToJson(Sheet4.range("B4:H4"), Sheet4.range("B5:H200"), 200, 7, myTxt, "SupportData")
  Call ExtractTankMainDataToJson(Sheet5.range("B4:D3"), Sheet5.range("B5:D200"), 200, 3, myTxt, "StandardData")
  Call ExtractTankMainDataToJson(Sheet6.range("B4:E4"), Sheet6.range("B5:E200"), 200, 4, myTxt, "RequirementData")
  Call ExtractTankMainDataToJson(Sheet7.range("B4:D4"), Sheet7.range("B5:D200"), 200, 3, myTxt, "OtherRequestData")

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing

  MsgBox "Extract Sucess!"

End Sub

Sub ExtractProjectInfoToJsonString(myTxt)
  Dim jsonString As String
  myTxt.Write "{"
  jsonString = ConvertString("PROJECT") & ":" & ConvertString(Sheet1.range("F2")) & "," & ConvertString("UNITNAME") & ":" & ConvertString(Sheet1.range("N2")) & "," & ConvertString("dataClass") & ":" & ConvertString("projectInfo")
  myTxt.Write jsonString
  myTxt.Write "}"
  myTxt.Write vbCr
End Sub

Sub ExtractTankMainDataToJson(keyRange, valueRange, rowNum, columnNum, myTxt, dataType)
  Dim row As Integer, column As Integer, jsonString As String, cellString As String, arr As Variant
  arr = ExtractOneRowDataToArray(keyRange)
  row = 1
  column = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    if valueRange.Cells(row, 1).Value <> "" Then 
      myTxt.Write "{"
      jsonString = ConvertString("dataClass") & ":" & ConvertString(dataType) & ","
      For column = 1 To columnNum
        cellString = Replace(valueRange.Cells(row, column).Value, ",", "，")
        ' the cell content may have ", convert it to #, reconvert it before insert the database 2021-07-13
        cellString = Replace(cellString, """", "#")
        cellString = Replace(cellString, ":", "：")
        jsonString = jsonString & ConvertString(arr(column-1)) & ":" & ConvertString(cellString) & ","
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

Function ConvertString(cellString)
  Dim resultString As String
  resultString = chr(34) & cellString & chr(34)
  ConvertString = resultString
End Function
