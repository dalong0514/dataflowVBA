Sub ExtractToxicitySubtanceData()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\tempdata\gsToxicitySubstance.txt"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True, Unicode:=True)

  myTxt.Write "#" + range("B2").Value
  myTxt.Write vbCr
  Call ExtractOneColumnData(Sheet1.range("B4:B500"), myTxt, 500)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing

  MsgBox "Extract Sucess!"
End Sub

Sub ExtractOneColumnData(range, myTxt, rowNum)
  Dim row As Integer
  row = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    If range.Cells(row, 1).Value <> "" Then
      myTxt.Write "#"
      myTxt.Write range.Cells(row, 1).Value
    End If
  Next row
End Sub