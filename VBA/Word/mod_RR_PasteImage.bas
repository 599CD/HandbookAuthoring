Attribute VB_Name = "mod_RR_PasteImage"
'
' PasteTrainingGraphic Macro
' Macro recorded 09/06/99 by Richard Rost
'
Sub PasteTrainingGraphic()

    Selection.PasteSpecial Link:=False, DataType:= _
        wdPasteDeviceIndependentBitmap, Placement:=wdInLine, DisplayAsIcon:=False
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
    getpicselection
    setborder
    
    ' move back right
    Selection.MoveRight Unit:=wdCharacter, Count:=4
    
End Sub

'
' PasteTrainingGraphic2 Macro
' Macro recorded 10/22/2009 by Richard Rost
'
Sub PasteTrainingGraphic2()

    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Paste
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
End Sub



'Probably the second one.
'Also found this:

'
' Macro1 Macro
' Macro recorded 10/22/2009 by Richard Rost
'
Sub Macro1()

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Paste
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub

