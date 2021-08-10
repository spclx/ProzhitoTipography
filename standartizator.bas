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

Sub Стандартизатор2()

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
    
    For Each li In ActiveDocument.Lists
        li.ConvertNumbersToText
    Next li

    'ActiveDocument.Lists(1).ConvertNumbersToText
    
    replaceWizzard "^b", "^p"
    replaceWizzard "^m", "^p"
    ' Замена абзацев
    ' "^l" -> "^p"
    replaceWizzard "^l", "^p"
    ' " ^p" -> "^p"
    replaceWizzard " ^p", "^p"
    ' форматирование всех концов абзаца
    replaceWizzard "^p", "^p"
    ' удаление всех двойных абзацев
    replaceWizzard "^p^p^p", "^p^p"
   
    ' убираются специальные ненужные знаки
    ' удаление неразрывных пробелов
    replaceWizzard "^s", " "
    replaceWizzard "^-", ""
    ' удаление табуляции
    replaceWizzard "^t", " "
    
    ' форматирование всех пробелов
    replaceWizzard " ", " "
    ' Замена заголовков
    replaceWizzard "###", "### "
  
    ' удаление всех двойных пробелов
    flag = True
    While flag = True
        Set MyRange = ActiveDocument.Content
        replaceWizzard "  ", " "
        MyRange.Find.Execute FindText:="  "
        If MyRange.Find.Found = False Then flag = False
    Wend
    ' "^p " -> "^p"
    replaceWizzard "^p ", "^p"

    ' Замена цитат
    replaceWizzard ">", ">"
    replaceWizzard "^p>^p", "^p     ^p"

    ' Замена тире и прочей фигни
    ' дефис -> длинное тире
    replaceWizzard " - ", " — "
    ' короткое тире -> длинное тире
    replaceWizzard "–", "—"
     replaceWizzard "^p-", "^p—"
    ' [пробел][знак препинания] -> [знак препинания]
    'Dim punctuationMark1, Mark;
    punctuationMark1 = Array(".", ",", ":", ";", ")", "]", "!", "?")
    For Each Mark In punctuationMark1
        oT = " " + Mark
        rT = Mark
        replaceWizzard oT, rT
    Next
    ' [знак препинания][пробел] -> [знак препинания]
    punctuationMark2 = Array("(", "[")
    For Each Mark In punctuationMark2
        oT = Mark + " "
        rT = Mark
        replaceWizzard oT, rT
    Next

    ' удаление лишних пробелов в частицах и некоторых предлогах
    particles = Array("то", "таки", "нибудь", "ка", "за", "под")
    For Each Particle In particles
        oT = "- " + Particle
        rT = "-" + Particle
        replaceWizzard oT, rT
    Next
    ' оставльные случаи дефис-пробел будут предупреждаться в Message Box

    ' сообщение, если есть сноски  нестандартного вида
    Message = "Обработка закончена" + Chr(13)

    Set MyRange = ActiveDocument.Content
    With MyRange.Find
        .Font.Bold = True
        .Execute FindText:="^f"
        If MyRange.Find.Found = True Then
            Message = Message + "ATTENTION! Сноски жирные" + Chr(13)
        End If
    End With

    Set MyRange = ActiveDocument.Content
    With MyRange.Find
        .Font.Italic = True
        .Execute FindText:="^f"
        If MyRange.Find.Found = True Then
            Message = Message + "ATTENTION! Сноски курсивные" + Chr(13)
        End If
    End With

    Set MyRange = ActiveDocument.Content
    With MyRange.Find
        .Font.Underline = True
        .Execute FindText:="^f"
        If MyRange.Find.Found = True Then
            Message = Message + "ATTENTION! Сноски подчёркнутые" + Chr(13)
        End If
    End With

    Set MyRange = ActiveDocument.Content
    With MyRange.Find
        .Font.StrikeThrough = True
        .Execute FindText:="^f"
        If MyRange.Find.Found = True Then
            Message = Message + "ATTENTION! Сноски зачёркнутые" + Chr(13)
        End If
    End With

    Set MyRange = ActiveDocument.Content
    With MyRange.Find
        .Execute FindText:="- "
        If MyRange.Find.Found = True Then
            Message = Message + "ATTENTION! тире-пробел" + Chr(13)
        End If
    End With
        
    signal = MsgBox(Message, vbInformation, "Обработка текстов")

End Sub



