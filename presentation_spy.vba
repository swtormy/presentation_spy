Sub start()
    On Error Resume Next
    Set SI_ = CreateObject("ADSystemInfo")
    Set Un_ = GetObject("LDAP://" & SI_.userName)
    
    'Определение имени пользователя
    userName = Un_.DisplayName
    
    'Определение номера недели
    'numWeek = DatePart("ww", Now(), vbMonday, vbFirstJan1) - 3
    
    'Определение адреса и названия активного файла презентации
    activePres = ActivePresentation.Path & "\" & ActivePresentation.Name
    
    'Определение необходимого адреса и названия файла
    mostPath = "С:\Program files\Files\<filename>"
    
    'Поиск номера недели в строке плюс количество символов в номере недели
    'для определения количества символов в строке
    'endStringActive = InStr(activePres, numWeek) + Len(numWeek)
    'endStringMost = InStr(mostPath, numWeek) + Len(numWeek)
    
    'Получение адреса и имени файла с учетом количества необходимых символов
    'activePresN = Left(activePres, endStringActive - 1)
    'mostPathN = Left(mostPath, endStringMost - 1)
    
    'Условия работы скрипта
    'Проверка на открытие в режиме чтения
    If ActivePresentation.ReadOnly Then
    'если необходимы действия при открытии в режиме чтения, то код писать сюда
    Else
        'Если файл открыт в режиме редактирования
        If Split(activePres, ".")(1) = "pptx" Then
            'Сверка адреса и имени открытого файла с необходимым
            If Left(activePres, 49) = mostPath Then
                'Открытие файла с логами для получения всех строк в переменную s
                Open "С:\Program files\Files\logs.txt" For Input As #1
                Dim s As String
                Input #1, s
                Close #1
                
                'Открытие файла с логами для записи нового лога
                Open "С:\Program files\Files\logs.txt" For Output As #2
                Write #2, userName & " открыл файл " & ActivePresentation.Path & "\" & ActivePresentation.Name & " в " & Now() & "/" & vbNewLine & s
                Close #2
            End If
        End If
    End If
    
    

End Sub
