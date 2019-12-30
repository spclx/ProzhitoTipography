Sub replaceWizzard(originalText, replasedText)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .StrikeThrough = False
    End With
    With Selection.Find
        .Text = originalText
        .Replacement.Text = replasedText
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
End Sub

Sub PracticeHelper()

    Set MyRange = ActiveDocument.Content

    ' Первоначальное форматирование текста
    ' Меняем  шрифт
    Selection.WholeStory
    With Selection.Font
    .Name = "Calibri"
    .Size = 12
    End With

    ' Меняем параметры абзаца
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .FirstLineIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 1
    End With

    ' Замена абзацев
    ' "^l" -> "^p"
    replaceWizzard "^l", "^p"
    ' "^b" -> "^p"
    replaceWizzard "^b", "^p"
    ' "^m" -> "^p"
    replaceWizzard "^m", "^p"
    ' удаление всех двойных абзацев
    replaceWizzard "^p^p^p", "^p^p"
   
    ' убираются специальные ненужные знаки
    ' удаление неразрывных пробелов
    replaceWizzard "^s", " "
    replaceWizzard "^-", ""
    ' удаление табуляции
    replaceWizzard "^t", " "
    
    ' дефис -> длинное тире
    replaceWizzard " - ", " — "
    ' короткое тире -> длинное тире
    replaceWizzard "–", "—"
    replaceWizzard "^p-", "^p—"
    ' удаление лишних пробелов в частицах и некоторых предлогах
    particles = Array("то", "таки", "нибудь", "ка", "за", "под")
    For Each Particle In particles
        oT = "- " + Particle
        rT = "-" + Particle
        replaceWizzard oT, rT
    Next
    ' выделение тире-пробелов
    Options.DefaultHighlightColorIndex = wdYellow
    With ActiveDocument.Range.Find
        .Text = "- "
        .Replacement.Highlight = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With

    Response = MsgBox("Обработка закончена")

End Sub

