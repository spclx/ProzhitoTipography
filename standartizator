Attribute VB_Name = "Module1"

Sub Стандартизатор()
Selection.WholeStory

' Меняем  шрифт
With Selection.Font
    .Name = "Calibri"
    .Size = 12
    .ColorIndex = wdBlack
End With

' Меняем параметры абзаца
With Selection.ParagraphFormat
    .LeftIndent = CentimetersToPoints(0)
    .RightIndent = CentimetersToPoints(0)
    .SpaceBefore = 0
    .SpaceAfter = 0
    .LineSpacingRule = wdLineSpaceSingle
    .FirstLineIndent = CentimetersToPoints(0)
End With

' Замена абзацев

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Замена цитат
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = ">"
        .Replacement.Text = ">"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "^p>^p"
        .Replacement.Text = "^p   ^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Замена заголовков
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "###"
        .Replacement.Text = "### "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = "###  "
        .Replacement.Text = "### "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' Замена тире и прочей фигни
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
    End With
    With Selection.Find
        .Text = " - "
        .Replacement.Text = " — "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
    End With
    With Selection.Find
        .Text = "–"
        .Replacement.Text = "—"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
' сообщение, если есть сноски  нестандартного вида
massage = "Обработка закончена" + Chr(13)

Set MyRange = ActiveDocument.Content
With MyRange.Find
    .Font.Bold = True
    .Execute FindText:="^f"
    If MyRange.Find.Found = True Then
        massage = massage + "ATTENTION! Сноски жирные" + Chr(13)
    End If
End With

Set MyRange = ActiveDocument.Content
With MyRange.Find
    .Font.Italic = True
    .Execute FindText:="^f"
    If MyRange.Find.Found = True Then
        massage = massage + "ATTENTION! Сноски курсивные" + Chr(13)
    End If
End With

Set MyRange = ActiveDocument.Content
With MyRange.Find
    .Font.Underline = True
    .Execute FindText:="^f"
    If MyRange.Find.Found = True Then
        massage = massage + "ATTENTION! Сноски подчёркнутые" + Chr(13)
    End If
End With

Set MyRange = ActiveDocument.Content
With MyRange.Find
    .Font.StrikeThrough = True
    .Execute FindText:="^f"
    If MyRange.Find.Found = True Then
        massage = massage + "ATTENTION! Сноски зачёркнутые" + Chr(13)
    End If
End With
        
Set MyRange = ActiveDocument.Content
With MyRange.Find
    .Execute FindText:="- "
    If MyRange.Find.Found = True Then
        massage = massage + "ATTENTION! тире-пробел" + Chr(13)
    End If
End With
    
signal = MsgBox(massage, vbInformation, "Обработка текстов")


End Sub
