Attribute VB_Name = "modChangeFont"
' Word Options | Customer Ribbon
' Keyboard Shorcuts Customise
'
' Keyboard Shortcut - Alt + C
Sub ChangeFont()
    
    Selection.Font.Name = "Courier New"
    
'    With ActiveWindow.Selection.TextRange.Font
'        .Name = "Courier New"
'        .Size = 10
'        .Bold = msoFalse
'        .Italic = msoFalse
'        .Underline = msoFalse
'        .Shadow = msoFalse
'        .Emboss = msoFalse
'        .BaselineOffset = 0.3
'        .AutorotateNumbers = msoFalse
'        .Color.SchemeColor = ppForeground
'    End With

End Sub

' Keyboard Shortcut - Alt + A
Sub ChangeColourAccess()
    Selection.Font.ColorIndex = wdDarkRed 'wdRed 'wdViolet
End Sub

' Keyboard Shortcut - Alt + W
Sub ChangeColourWord()
    Selection.Font.ColorIndex = wdBlue
End Sub

' Keyboard Shortcut - Alt + O
Sub ChangeColourOutlook()
    Selection.Font.ColorIndex = wdDarkYellow
End Sub

' Keyboard Shortcut - Alt + U
Sub Capitalise()
    Selection.Range.Case = wdTitleWord
End Sub
