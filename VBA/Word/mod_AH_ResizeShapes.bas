Attribute VB_Name = "mod_AH_ResizeShapes"
'https://groups.google.com/forum/?fromgroups=#!topic/microsoft.public.word.vba.general/YZFwpNk4oaA

Sub ResizePicture_AH()

    Set image = Selection.ShapeRange
    'Set image = Selection.InlineShapes
    With image
        .Height = 100
        .Width = 100
    End With
    
End Sub


Sub ResizePicture()

    ' change these numbers to the maximum width and height
    ' (in inches) to make the inserted pictures
       Const PicWidth = 1.9
       Const PicHeight = 2.25

       Dim Photo As InlineShape

        Set Photo = .InlineShapes.AddPicture(FileName:=FName, LinkToFile:=False, SaveWithDocument:=True, Range:=PicRg)
        'Set Photo = .Shapes.AddPicture(FileName:=FName, LinkToFile:=False, SaveWithDocument:=True, Range:=PicRg)
        With Photo
            RatioW = CSng(InchesToPoints(PicWidth)) / .Width
            RatioH = CSng(InchesToPoints(PicHeight)) / .Height
            
            ' choose the smaller ratio
            If RatioW < RatioH Then
                RatioUse = RatioW
            Else
                RatioUse = RatioH
            End If
            
            ' size the picture to fit the cell
            .Height = .Height * RatioUse
            .Width = .Width * RatioUse
        End With

End Sub


'http://yuriy-okhmat.blogspot.co.uk/2011/07/how-to-resize-all-images-in-word.html
Sub AllPictSize()
    Dim targetWidth As Integer
    Dim oShp As Shape
    Dim oILShp As InlineShape
 
    targetWidth = 16
 
    For Each oShp In ActiveDocument.Shapes
        With oShp
            .Height = AspectHt(.Width, .Height, CentimetersToPoints(targetWidth))
            .Width = CentimetersToPoints(targetWidth)
        End With
    Next
 
    For Each oILShp In ActiveDocument.InlineShapes
        With oILShp
            .Height = AspectHt(.Width, .Height, CentimetersToPoints(targetWidth))
            .Width = CentimetersToPoints(targetWidth)
        End With
    Next
End Sub
 
Private Function AspectHt(ByVal origWd As Long, ByVal origHt As Long, ByVal newWd As Long) As Long
    If origWd <> 0 Then
        AspectHt = (CSng(origHt) / CSng(origWd)) * newWd
    Else
        AspectHt = 0
    End If
End Function

