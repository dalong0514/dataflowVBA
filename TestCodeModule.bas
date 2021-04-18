Attribute VB_Name = "TestCodeModule"
With ThisDrawing
  .Layers.Add "Disclaimer"
  Dim objMText As AcadMText
  Dim insPt(2) As Double
  insPt(0) = 0.25: insPt(1) = 1.75: insPt(2) = 0
  Set objMText = .ModelSpace.AddMText(insPt, 15, _
    "Confidential: This drawing is for use by internal" & _
    "employees and approved vendors only")
  objMText.Layer = "Disclaimer"
  objMText.Height = 0.5
  If .ActiveSpace = acPaperSpace Then
    .ActiveSpace = acModelSpace
    End If
  .Save
End With

