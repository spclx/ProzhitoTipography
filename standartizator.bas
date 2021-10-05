Sub replaceWizzard(originalText, replasedText, wildcard)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    'With Selection.Find.Replacement.Font
    '    .Bold = False
    '    .Italic = False
    '    .Underline = False
    '    .StrikeThrough = False
    'End With
    With Selection.Find
        .Text = originalText
        .Replacement.Text = replasedText
        .Highlight = False
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .matchCase = False
        .MatchWholeWord = False
        .MatchWildcards = wildcard
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub highlighting(originalText)
    Options.DefaultHighlightColorIndex = wdYellow
    With ActiveDocument.Range.Find
        .Text = originalText
        .MatchWildcards = True
        .Replacement.Highlight = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub Стандартизатор2()

    Set MyRange = ActiveDocument.Content
    
    ' Первоначальное форматирование текста
    ' Меняем  шрифт
    Selection.WholeStory
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
    End With

    ' Преобразование нумерованных и маркированных списков в текст
    For Each li In ActiveDocument.Lists
        li.ConvertNumbersToText
    Next li

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
    
    ' Избавление от разрывов абзацев
    replaceWizzard "^b", "^p", False
    replaceWizzard "^m", "^p", False
    ' Замена абзацев
    ' "^l" -> "^p"
    replaceWizzard "^l", "^p", False
    ' " ^p" -> "^p"
    replaceWizzard " ^p", "^p", False
    ' форматирование всех концов абзаца
    replaceWizzard "^p", "^p", False
    ' удаление всех двойных абзацев
    replaceWizzard "^p^p^p", "^p^p", False
   
    ' убираются специальные ненужные знаки
    ' удаление неразрывных пробелов
    replaceWizzard "^s", " ", False
    replaceWizzard "^-", "", False
    ' удаление табуляции
    replaceWizzard "^t", " ", False
  
    ' удаление всех двойных пробелов
    flag = True
    While flag = True
        Set MyRange = ActiveDocument.Content
        replaceWizzard "  ", " ", False
        MyRange.Find.Execute FindText:="  "
        If MyRange.Find.Found = False Then flag = False
    Wend
    
    ' "^p " -> "^p"
    replaceWizzard "^p ", "^p", False
    replaceWizzard "^p-", "^p—", False

    ' Замена цитат
    replaceWizzard "^p>", "^p", False

    ' Замена тире и прочей фигни
    ' короткое тире между цифрами
    replaceWizzard "([0-9])-([0-9])", "\1–\2", True
    replaceWizzard "([0-9]) - ([0-9])", "\1–\2", True
    replaceWizzard "([0-9]) – ([0-9])", "\1–\2", True
    replaceWizzard "([0-9])—([0-9])", "\1–\2", True
    replaceWizzard "([0-9]) — ([0-9])", "\1–\2", True
    replaceWizzard "([IVX])-([0-9])", "\1–\2", True
    replaceWizzard "([IVX])—([0-9])", "\1–\2", True
    replaceWizzard "([IVX]) - ([0-9])", "\1–\2", True
    replaceWizzard "([IVX]) — ([0-9])", "\1–\2", True
    replaceWizzard "([A-zА-я])—([0-9])", "\1-\2", True
    
    ' исправление макроса для дат 0000-00-00
    replaceWizzard "([0-9]{4})–([0-9]{2})–([0-9]{2})", "\1-\2-\3", True

    ' удаление лишнего пробела между инициалами
    replaceWizzard "([А-Я].)([А-Я][a-я])", "\1 \2", True
    replaceWizzard "([А-Я].) ([А-Я].)", "\1\2", True
    ' добавление пробела между цифрой и некоторыми буквами и выделение этих мест для проверки
    replaceWizzard "([0-9])([гмч])", "\1 \2", True
    highlighting "[0-9] [гмч]"
    ' пробел-дефис-пробел -> длинное тире
    replaceWizzard " - ", " — ", False
    ' пробел-короткое тире-пробел -> длинное тире
    replaceWizzard " – ", " — ", False
    ' Замена на длинное тире в конце и начале абзаца
    replaceWizzard "-^p", "—^p", False
    replaceWizzard "–^p", "—^p", False
    replaceWizzard "^p-", "^p—", False
    replaceWizzard "^p–", "^p—", False

    ' удаление лишних пробелов в тегах
    replaceWizzard "[" & Chr(34) & "»”][>] ", "" & Chr(34) & ">", True
    replaceWizzard " </", "</", False
    replaceWizzard "=[" & Chr(34) & "«“] ([0-9]{1;10})", "=" & Chr(34) & "\1", True
    replaceWizzard "([0-9]{1;10}) [" & Chr(34) & "»”] ", "\1" & Chr(34) & " ", True
    replaceWizzard "([A-я])<персона ", "\1 <персона ", True
    replaceWizzard "персона[>]([A-я])", "персона> \1", True

    ' [пробел][знак препинания] -> [знак препинания]
    'Dim punctuationMark1, Mark;
    punctuationMark1 = Array(".", ",", ":", ";", ")", "]", "!", "?")
    For Each Mark In punctuationMark1
        oT = " " + Mark
        rT = Mark
        replaceWizzard oT, rT, False
    Next
    ' [знак препинания][пробел] -> [знак препинания]
    punctuationMark2 = Array("(", "[")
    For Each Mark In punctuationMark2
        oT = Mark + " "
        rT = Mark
        replaceWizzard oT, rT, False
    Next

    ' выделение в тексте всех нестандартных случаев тире и дефисов
    highlighting "[A-я0-9]- [A-я]" ' дефис
    highlighting "[A-я0-9] -[A-я]"
    highlighting "[A-я0-9]-[A-я]"
    highlighting "[A-я0-9]– [A-я]" ' короткое тире
    highlighting "[A-я0-9] –[A-я]"
    highlighting "[A-я0-9]–[A-я]"
    highlighting "[A-я0-9]— [A-я]"  ' длинное тире
    highlighting "[A-я0-9] —[A-я]"
    highlighting "[A-я0-9]—[A-я]"
    highlighting "[ A-zА-я][ A-zА-я]^0013[A-zА-я]"

    ' удаление лишних пробелов в частицах и некоторых предлогах
    particles = Array("то", "таки", "нибудь", "ка", "за", "под")
    For Each Particle In particles
        oT = "- " + Particle
        rT = "-" + Particle
        replaceWizzard oT, rT, False
    Next

End Sub
