Public Class AI

    'lpApplicationName - Раздел-имя,заключенное в квадратные скобки [] и группирующее ключи и значения. 
    'lpKeyName - Значение ключа.Ключ должен быть уникальным только внутри своего раздела. 
    'lpDefault -Возвращаемое значение, если правильное(допустимое) значение не может читаться. 
    'lpReturnedString - Строка фиксированной длины, получаемая при чтении любой строки файла или lpDefault. 
    'nSize - Длина в символах переменной lpReturnedString. 
    'lpFileName - Имя INI-файла для чтения.

    Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                         ByVal lpDefault As String, ByVal lpReturnedString As String, _
                         ByVal nSize As UInteger, ByVal lpFileName As String) As UInt32

    'lpApplicationName - Значение раздела INI-файла. 
    'lpKeyName - Значение ключа. 
    'lpString - устанавлимое строковое значение. 
    'lpFileName - Имя INI-файла.

    Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
                    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                     ByVal lpString As String, ByVal lpFileName As String) As Long


    Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
                    (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                     ByVal nDefault As Integer, ByVal lpFileName As String) As UInt32


    ' Получение строкового параметра из секции
    Shared Function GetIniStringParam(ByVal section As String, ByVal key As String, ByVal fPath As String, Optional ByVal def As String = "<Unknown>") As String
        Dim max_len As UInt32 = 255
        Dim svalue = Space(max_len)  ' обеспечиваем достаточно места для функции, чтобы поместить значение в буфер
        ' читаем файл, slength длина получаемой строки
        Dim slength As UInt32 = GetPrivateProfileString(section, key, def, svalue, max_len, fPath)
        Return Left(svalue, slength) ' извлекаем нужную строчку из буфера
    End Function

    ' Получение числового параметра из секции
    Shared Function GetIniParam(ByVal section As String, ByRef key As String, ByRef fPath As String, Optional ByRef def As Integer = -1) As Integer
        ' читаем в файл
        Return GetPrivateProfileInt(section, key, def, fPath)
    End Function

    ' Запись параметра или секции в AI (возвращает 0 при ошибке в выполнении и 1 при успешном выполнении)
    Shared Function PutIniParam(ByVal section As String, ByRef key As String, ByRef value As Integer, ByRef fPath As String) As Long
        ' запись в файл
        Return WritePrivateProfileString(section, key, value.ToString, fPath)
    End Function

    ' получает массив всех имен AIPACKET
    Shared Function GetAll_AIPacket(ByRef Path As String) As String(,)
        Dim massive(,) As String = Nothing
        Dim file() As String = IO.File.ReadAllLines(Path)
        Dim i As Integer
        For n As Integer = 0 To file.Length - 1
            If file(n).StartsWith("[") Then
                ReDim Preserve massive(1, i)
                massive(0, i) = file(n).Substring(1, (file(n).LastIndexOf("]") - 1))
                massive(1, i) = n
                i += 1
            End If
        Next
        Return massive
    End Function

End Class
