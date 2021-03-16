Public Sub CreateTxtFile()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim i As Integer
  Dim nowDate As String
  Dim sht As Worksheet
  
  nowDate = CDate(Now())
  nowDate = Replace(nowDate, ":", "") 

  MyFName = "E:\2.txt" 
  Set fso = CreateObject("Scripting.FileSystemObject") 
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True) 

  For Each sht In ThisWorkbook.Worksheets 
    myTxt.Write "dalong,\\n"
  Next
  
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub
