' Процедура, которая заменяет названия месяцев
Sub DateWizzard(CaseCode As Integer)
              
' Создание массива с названиями месяцев

  Dim Month, Counter, MonthRoman, MonthArabic, InnormalMonth 
' Названия месяцев для арабских чисел
  Month = Array("декабря", "ноября", "октября", "января", "февраля", "марта", "апреля", _
                "мая", "июня", "июля", "августа", "сентября")
' Названия месяцев для римских чисел
  InnormalMonth = Array("декабря", "ноября", "октября", "августа", "июля", "июня", _
                "мая", "сентября", "апреля", "марта", "февраля", "января")

' Массивы с римскими и арабскими числами соответственно
  MonthRoman = Array("XII", "XI", "X", "VIII", "VII", "VI", "V", "IX", "IV", "III", "II", "I")
  MonthArabic = Array(12, 11, 10, 1, 2, 3, 4, 5, 6, 7, 8, 9)

  Counter = 1 ' Счётчик шагов
   
  While Counter <= 12
    ' Вычисление объектов поиска и замены
    CurrentMonth = Month(Counter - 1)
    CurrentInnormalMonth = InnormalMonth(Counter - 1)
    CurrentMonthRoman = MonthRoman(Counter - 1)
    CurrentMonthArabic = MonthArabic(Counter - 1)
            
    Dim OriginalText As String, ReplacedText As String
            
    ' Создание полей для поиска (OriginalText) и замены (ReplacedText) в зависимости от кейса
    Select Case CaseCode
      Case 1 '1.I.1900
        OriginalText = "(^0013)([0-9]{1;2}).(" & CurrentMonthRoman & ").([0-9]{1;4})"
        ReplacedText = "\1\2 " & CurrentInnormalMonth & ". \2.\3.\4"
      Case 2 '1.I
        OriginalText = "(^0013)([0-9]{1;2}).(" & CurrentMonthRoman & ")"
        ReplacedText = "\1\2 " & CurrentInnormalMonth & ". \2.\3"
      Case 3 'Case3: 1.1.2000
        OriginalText = "(^0013)([0-9]{1;2}).(" & CurrentMonthArabic & ").([0-9]{1;4})"
        ReplacedText = "\1\2 " & CurrentMonth & ". \2.\3.\4"
      Case 4 'Case4: 1/I
        OriginalText = "(^0013)([0-9]{1;2})/(" & CurrentMonthRoman & ")"
        ReplacedText = "\1\2 " & CurrentInnormalMonth & ". \2/\3"
      Case 5 'Case5: 1.1
        OriginalText = "(^0013)([0-9]{1;2}).(" & CurrentMonthArabic & ")"
        ReplacedText = "\1\2 " & CurrentMonth & ". \2.\3"
      Case 6 'Case6: 1/I-00
        OriginalText = "(^0013)(" & CurrentMonthRoman & ")/(<[A-Z]@>)-([0-9]{1;4})"
        ReplacedText = "\1\2 " & CurrentInnormalMonth & ". \2/\3-\4"
    End Select
            
    ' Осуществление замены по указанным выше шаблонам
    With ActiveDocument.Content.Find
      .Execute FindText:=OriginalText, MatchWildcards:=True, ReplaceWith:=ReplacedText, Replace:=wdReplaceAll
    End With
                        
    Counter = Counter + 1
  
  Wend

End Sub


Sub Датизатор()
'
' Датизатор Макрос — основной макрос, который ищет случаи с особым написанием дат
   
  Dim CaseCounter As Integer
' Счётчик случаев
  CaseCounter = 1

  While CaseCounter <= 6
    Set MyRange = ActiveDocument.Content
    
    ' Поиск особых случаев
    Dim TextForFind As String
    Select Case CaseCounter
      Case 1 'Case 1: 1.I.2000
        TextForFind = "(^0013)([0-9]{1;2}).(<[A-Z]@>).([0-9]{1;4})"
      Case 2 'Case2: 1.I
        TextForFind = "(^0013)([0-9]{1;2}).(<[A-Z]@>)"
      Case 3 'Case3: 1.1.2000
        TextForFind = "(^0013)([0-9]{1;2}).([0-9]{1;2}).([0-9]{1;4})"
      Case 4 'Case4: 1/I
        TextForFind = "(^0013)([0-9]{1;2})/(<[A-Z]@>)"
      Case 5 'Case5: 1.1
        TextForFind = "(^0013)([0-9]{1;2}).([0-9]{1;2})"
      Case 6 'Case6: 1/I-00
        TextForFind = "(^0013)([0-9]{1;2})/(<[A-Z]@>)-([0-9]{1;4})"
    End Select
    
    ' Ссылка на замену особых случаев, если условие выполняется
    With MyRange.Find
      .Execute FindText:=TextForFind, MatchWildcards:=True
      If MyRange.Find.Found = True Then
        DateWizzard CaseCounter
      End If
    End With
    
    CaseCounter = CaseCounter + 1
 
  Wend

End Sub
