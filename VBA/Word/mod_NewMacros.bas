Attribute VB_Name = "mod_NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.mod_NewMacros.Macro1"
' Macro1 Macro
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.EscapeKey
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.mod_NewMacros.Macro2"
' Macro2 Macro
    Selection.InlineShapes(1).Height = 170.1
End Sub


' http://www.techsupportforum.com/forums/f57/solved-creating-a-macro-in-word-to-resize-an-image-342723.html
Sub ReformatPictures()
    On Error Resume Next
    Dim oShp As Shape
    Dim iShp As InlineShape
    Dim ShpScale As Double
    With ActiveDocument
      For Each oShp In .Shapes
        With oShp
          If .Type = msoPicture Or msoLinkedPicture Then
            ShpScale = CentimetersToPoints(11) / .Width
            .Width = .Width * ShpScale
            If .LockAspectRatio = False Then .Height = .Height * ShpScale
          End If
        End With
      Next oShp
      For Each iShp In .InlineShapes
        With iShp
          If .Type = wdInlineShapePicture Or wdInlineShapeLinkedPicture Then
            ShpScale = CentimetersToPoints(11) / .Width
            .Width = .Width * ShpScale
            If .LockAspectRatio = False Then .Height = .Height * ShpScale
          End If
        End With
      Next iShp
    End With
    MsgBox "Finished Reformatting."
End Sub


' http://superuser.com/questions/369977/word-resize-image-by-percent-macro
Sub PicResize()
     Dim PecentSize As Integer

     PercentSize = 75

     If Selection.InlineShapes.Count > 0 Then
         Selection.InlineShapes(1).ScaleHeight = PercentSize
         Selection.InlineShapes(1).ScaleWidth = PercentSize
     Else
         Selection.ShapeRange.ScaleHeight Factor:=(PercentSize / 100), _
           RelativeToOriginalSize:=msoCTrue
         Selection.ShapeRange.ScaleWidth Factor:=(PercentSize / 100), _
           RelativeToOriginalSize:=msoCTrue
     End If
End Sub


' http://stackoverflow.com/questions/6407194/finding-image-reference-in-ms-word-using-vba
'With ActiveDocument.InlineShapes(ActiveDocument.InlineShapes.Count)
'    .Height = 314.95 ' or whatever
'End With


Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.mod_NewMacros.Macro3"
' Macro3 Macro
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.InlineShapes(1).Height = 170.1
End Sub
