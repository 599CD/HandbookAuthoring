Attribute VB_Name = "mod_JW_Resize"

'Assign a Keyboard Shortcut to a Macro
' File | Options | Customise Ribbon
' Keyboard Shortcuts - Customise ...
' Scroll down in Categories to Macros
' Choose the Macro name on the right hand side
' Type your Shortcut
' Press Assign.

' Keyboard Shortcut - Alt + S
' Jamie Waite
Sub ShrinkImg()
    'Dim pastedImage As InlineShape  'Decalre Image
    'Set pastedImage = ThisDocument.InlineShapes(ThisDocument.InlineShapes.Count)  'Gets The Last Image you pasted into the Documnet
    Dim SelectedImage As InlineShape  'Decalre Image
    
    'Get Selected Images
    Selection.Find.Execute Replace:=2
    Selection.Expand wdParagraph
    
    Set SelectedImage = Selection.InlineShapes(1)
    SelectedImage.Height = "170"    'Change Height
    'SelectedImage.Width = "#"      'Change Width
    
     Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    'BOOM!
End Sub

Sub PastedImage()
    Dim PastedImage As InlineShape  'Decalre Image
    Set PastedImage = ThisDocument.InlineShapes(ThisDocument.InlineShapes.Count)  'Gets The Last Image you pasted into the Documnet
    PastedImage.Height = "1200"  'Change Height
    PastedImage.Width = "450"    'Change Width
    'BOOM!!"
End Sub
