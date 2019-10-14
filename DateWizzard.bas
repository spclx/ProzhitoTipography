Sub DateWizzard(OriginalText As String, ReplacedText As String)
              
' Создание массива с названиями месяцев

Dim Month, Counter, MonthRoman
Month = Array("января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", _
                        "сентября", "октября", "ноября", "декабря")
MonthRoman = Array("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII")

Counter = 1 ' Счётчик месяцев

' Кейс № 1 : 1.XII.2019
    'If Allert = vbYes Then
        While Counter <= 12
            
            CurrentMonth = Month(Counter - 1)
            CurrentMonthRoman = MonthRoman(Counter - 1)
            
            OriginalText = "(^0013)([0-9]{1;2}).(" & CurrentMonthRoman & ").([0-9]{1;4})"
            ReplacedText = "\1\2 " & CurrentMonth & ". \2.\3.\4"
            
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            
            With ActiveDocument.Content.Find
                .Execute FindText:=OriginalText, MatchWildcards:=True, ReplaceWith:=ReplacedText, Replace:=wdReplaceAll
                
            End With
            
            
            Counter = Counter + 1
        Wend
End Sub


Sub Датизатор()
'
' Датизатор Макрос

Allert = MsgBox("Перед началом выполнения макроса поставьте курсор в начале документа", vbYesNo, "ВНИМАНИЕ!")
If Allert = vbYes Then
    DateWizzard "(^0013)([0-9]{1;2}).(" & CurrentMonthRoman & ").([0-9]{1;4})", "\1\2 " & CurrentMonth & ". \2.\3.\4"
End If
End Sub
