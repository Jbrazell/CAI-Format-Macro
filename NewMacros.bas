Attribute VB_Name = "NewMacros"

Sub caiFormat()
'
' caiFormat Macro
'
'
    Dim doc As Document
    Dim rng As Range
    Set doc = ActiveDocument
    Set rng = doc.Paragraphs(1).Range

    rng.WholeStory
    rng.Select

    Selection.WholeStory
    Selection.ClearFormatting
    Selection.ParagraphFormat.SpaceBefore = 0
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt15
    Selection.ParagraphFormat.IndentFirstLineCharWidth (2)
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.Font.ColorIndex = wdBlack

    ActiveDocument.Footnotes.Convert
    ActiveDocument.Endnotes.Location = wdEndOfDocument
    ActiveDocument.Endnotes.NumberStyle = wdNoteNumberStyleArabic

    Set rng = doc.Endnotes(1).Range
        rng.WholeStory
        rng.Select
        Selection.Font.Name = "Times New Roman"

    With ActiveDocument
    Set rng = doc.Range.Sentences(1)
        rng.Select
        Selection.InsertParagraphAfter
        Selection.Font.Bold = True
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

    With ActiveDocument
        Selection.EndKey Unit:=wdStory
        Selection.InsertParagraph
        Selection.InsertAfter Text:="Word Count: " & .ComputeStatistics(wdStatisticWords)
    End With

    Selection.HomeKey Unit:=wdStory

End Sub
