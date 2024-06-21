Attribute VB_Name = "Module1"
Sub ToggleVisibility()
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes("Group 9242")

    If shp.Visible = msoTrue Then
        shp.Visible = msoFalse
    Else
        shp.Visible = msoTrue
    End If
End Sub

