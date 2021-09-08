' 2021-08-27
' refactored at 2021-09-01
' refactored at 2021-09-02
Sub ExtractGsLcModuleData()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\gsLcModuleData.json"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractProjectInfoToJsonString(myTxt)
  Call ExtractRangeDataToJsonStringByDataType(Sheet1.range("B4:L4"), Sheet1.range("B6:L100"), 100, 3, myTxt, "moduleBuild")
  Call ExtractRangeDataToJsonStringByDataType(Sheet2.range("B2:L2"), Sheet2.range("B4:L100"), 100, 6, myTxt, "moduleEquip")
  Call ExtractRangeDataToJsonStringByDataType(Sheet3.range("C2:S2"), Sheet3.range("C4:S500"), 500, 18, myTxt, "moduleCorrespond")

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing

  MsgBox "Extract Sucess!"
End Sub

Sub ExtractRangeDataToJsonStringByDataType(keyRange, valueRange, rowNum, columnNum, myTxt, dataType)
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

Sub ExtractProjectInfoToJsonString(myTxt)
  Dim jsonString As String
  myTxt.Write "{"
  jsonString = ConvertString("dataClass") & ":" & ConvertString("projectInfo") & "," & ConvertString("PROJECT") & ":" & ConvertString(Sheet1.range("C1")) & "," & ConvertString("UNITNAME") & ":" & ConvertString(Sheet1.range("C2")) & "," & ConvertString("ProjM") & ":" & ConvertString(Sheet1.range("I1")) & "," & ConvertString("SpeciM") & ":" & ConvertString(Sheet1.range("I2")) & "," & ConvertString("Made") & ":" & ConvertString(Sheet1.range("K1")) & "," & ConvertString("Chkd") & ":" & ConvertString(Sheet1.range("K2")) & "," & ConvertString("Appr") & ":" & ConvertString(Sheet1.range("M1")) & "," & ConvertString("AuthD") & ":" & ConvertString(Sheet1.range("M2")) & "," & ConvertString("projectNum") & ":" & ConvertString(Sheet1.range("G1")) & "," & ConvertString("monomertNum") & ":" & ConvertString(Sheet1.range("G2"))
  myTxt.Write jsonString
  myTxt.Write "}"
  myTxt.Write vbCr
End Sub
