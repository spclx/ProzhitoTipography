
' Процедура, которая заменяет названия месяцев
Sub DateWizzard2()

' Модель датизатора лежит здесь: https://docs.google.com/spreadsheets/d/1t3m6UfmC_jD2kmaUirafpR1-hYx6Gz5Fl5cEO1VjgOg/edit?usp=sharing
' Создание массива с названиями месяцев

' Объектная модель Scripting не работает на Mac, поэтому вводится проверка системы
  #If Mac = False Then
    Dim d
    Set d = CreateObject("Scripting.Dictionary")

  #Else
' На Mac нет объекта Scripting, поэтому используется сторонняя реализация класса
' Dictionary: https://sysmod.wordpress.com/2011/11/02/dictionary-class-in-vba-instead-of-scripting-dictionary/
' Дополнительно установить ещё один класс: http://www.sysmod.com/KeyValuePair.cls
    Dim d As Dictionary
    Set d = New Dictionary
  #End If

    d.Add "XII", "декабря"
    d.Add "XI", "ноября"
    d.Add "X", "октября"
    d.Add "VIII", "августа"
    d.Add "VII", "июля"
    d.Add "VI", "июня"
    d.Add "V", "мая"
    d.Add "IX", "сентября"
    d.Add "IХ", "сентября"
    d.Add "IV", "апреля"
    d.Add "III", "марта"
    d.Add "II", "февраля"
    d.Add "I", "января"
    d.Add "ХII", "декабря"
    d.Add "ХI", "ноября"
    d.Add "Х", "октября"
    d.Add "12", "декабря"
    d.Add "11", "ноября"
    d.Add "10", "октября"
    d.Add "09", "сентября"
    d.Add "08", "августа"
    d.Add "07", "июля"
    d.Add "06", "июня"
    d.Add "05", "мая"
    d.Add "04", "апреля"
    d.Add "03", "марта"
    d.Add "02", "февраля"
    d.Add "01", "января"
    d.Add "9", "сентября"
    d.Add "8", "августа"
    d.Add "7", "июля"
    d.Add "6", "июня"
    d.Add "5", "мая"
    d.Add "4", "апреля"
    d.Add "3", "марта"
    d.Add "2", "февраля"
    d.Add "1", "января"
  
  #If Mac = False Then
    For Each i In d
        originalText = "(^0013)([0-9]{1;2}).(" & i.Key & ").([0-9]{1;4})"
        ReplacedText = "\1\2 " & i.Item '& ". " ' \2.\3.\4"
      ' Осуществление замены по указанным выше шаблонам
      With ActiveDocument.Content.Find
          .Execute FindText:=originalText, MatchWildcards:=True, ReplaceWith:=ReplacedText, Replace:=wdReplaceAll
      End With
    Next
  #Else
    
    Dim i As KeyValuePair
  
    For Each i In d.KeyValuePairs
      originalText = "(^0013)([0-9]{1;2}).(" & i.Key & ").([0-9]{1;4})"
      ReplacedText = "\1\2 " & i.Value '& ". " ' \2.\3.\4"
      ' Осуществление замены по указанным выше шаблонам
      With ActiveDocument.Content.Find
          .Execute FindText:=originalText, MatchWildcards:=True, ReplaceWith:=ReplacedText, Replace:=wdReplaceAll
      End With
    Next
    
  #End If
  
End Sub

