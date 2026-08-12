Attribute VB_Name = "Module1"
Sub SATIN_ALMA_UYGULAMASINI_AÇ()
Attribute SATIN_ALMA_UYGULAMASINI_AÇ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SATIN_ALMA_UYGULAMASINI_AÇ Makro
'

'
    Selection.Characters.Text = "SATIN ALMA UYGULAMASINI AÇ"
    With Selection.Characters(Start:=1, Length:=26).Font
        .Name = "Aptos Narrow"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        .ColorIndex = 1
    End With
    ActiveCell.Offset(13, -3).Range("A1").Select
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    ActiveCell.Offset(1, 2).Range("A1").Select
    ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
End Sub
