Attribute VB_Name = "Module1"
Sub Save_Object_As_Picture()
    ' Объявление переменных
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String, productName As String
    Dim topLeftCell As Range
 
    ' Установка пути для сохранения изображений
    sImagesPath = ActiveWorkbook.Path & "\images\" ' Папка для сохранения изображений в текущем каталоге книги
    
    ' Создание папки, если она не существует
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath ' Создать папку для изображений, если её нет
    End If
    
    ' Отключение обновления экрана и предупреждений, чтобы ускорить выполнение
    On Error Resume Next ' Игнорировать ошибки
    Application.ScreenUpdating = False ' Отключить обновление экрана
    Application.DisplayAlerts = False ' Отключить предупреждения

    ' Установка текущего листа и создание временного листа
    Set wsSh = ActiveSheet ' Установка активного листа
    Set wsTmpSh = ActiveWorkbook.Sheets.Add ' Добавление временного листа для работы с графиком

    ' Перебор всех объектов на активном листе
    For Each oObj In wsSh.Shapes
        ' Проверка, является ли объект изображением
        If oObj.Type = 13 Then ' Тип 13 — это изображения
            li = li + 1 ' Счетчик для имен изображений
            
            ' Получаем ячейку, где находится верхний левый угол объекта
            Set topLeftCell = oObj.topLeftCell
            
            ' Получаем имя продукта из первого столбца (колонка A) той же строки, что и изображение
            productName = wsSh.Cells(topLeftCell.Row, 1).Value ' Наименование из столбца A
            
            ' Удаляем недопустимые символы из имени файла
            productName = Replace(productName, "/", "_") ' Заменяем слэши
            productName = Replace(productName, "\", "_") ' Заменяем обратные слэши
            productName = Replace(productName, ":", "_") ' Заменяем двоеточия
            productName = Replace(productName, "*", "_") ' Заменяем звездочки
            productName = Replace(productName, "?", "_") ' Заменяем вопросительные знаки
            productName = Replace(productName, """", "_") ' Заменяем кавычки
            productName = Replace(productName, "<", "_") ' Заменяем меньшие знаки
            productName = Replace(productName, ">", "_") ' Заменяем большие знаки
            productName = Replace(productName, "|", "_") ' Заменяем вертикальные линии
            
            ' Если имя продукта пустое, используем стандартное имя
            If productName = "" Then
                productName = "img" & li
            End If
            
            ' Копируем изображение
            oObj.Copy

            ' Использование временного графика для экспорта изображения
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0 ' Убираем границы графика
                .Parent.Select ' Выбираем график
                .Paste ' Вставляем изображение в график
                .Export Filename:=sImagesPath & productName & ".jpg", FilterName:="JPG" ' Экспортируем изображение как файл JPG
                .Parent.Delete ' Удаляем временный график после сохранения изображения
            End With
            
            ' Записываем имя файла в ячейку, где находилось изображение
            oObj.topLeftCell.Value = productName ' Записываем имя файла в ячейку
        End If
    Next oObj

    ' Освобождение памяти
    Set oObj = Nothing
    Set wsSh = Nothing
    wsTmpSh.Delete ' Удаление временного листа

    ' Включаем обратно обновление экрана и предупреждения
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Сообщение о завершении процесса
    MsgBox "Объекты сохранены в папке: " & sImagesPath, vbInformation, "Успех"
End Sub

