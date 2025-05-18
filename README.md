Sub ExtractImagesFast()
    Dim savePath As String
    Dim htmlPath As String
    Dim doc As Document

    ' Укажите путь для сохранения изображений
    savePath = "C:\макросы\картинки\" ' Замените на нужный путь

    ' Сохраняем документ в формате HTML
    htmlPath = savePath & "temp.html"
    ActiveDocument.SaveAs2 FileName:=htmlPath, FileFormat:=wdFormatHTML

    ' Сообщаем пользователю, где искать изображения
    MsgBox "Изображения извлечены и сохранены в папке: " & savePath & "temp_files", vbInformation

    ' Удаляем временный HTML файл, но оставляем папку с изображениями
    Kill htmlPath
End Sub
