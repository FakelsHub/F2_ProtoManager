Imports System.IO

Friend Class AI

    Friend Const AIFILE As String = "\data\AI.txt"

    Friend Const Unknown As String = "<NotSpecified>"
    Friend Const endPackedID As String = "EOFLINE"

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
    Shared Function GetIniStringParam(ByVal section As String, ByVal key As String, ByVal fPath As String, Optional ByVal def As String = Unknown) As String
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
    Shared Function PutIniParam(ByVal section As String, ByRef key As String, ByVal value As String, ByVal fPath As String) As Long
        ' запись в файл
        Return WritePrivateProfileString(section, key, value, fPath)
    End Function

    ' получает массив всех имен AIPACKET
    Shared Function GetAll_AIPacket(ByRef Path As String) As Dictionary(Of String, Integer)
        Dim packet As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()
        Dim fileData As String() = File.ReadAllLines(Path)

        For n = 0 To fileData.Length - 1
            If fileData(n).StartsWith("[") Then
                packet.Add(fileData(n).TrimEnd().Substring(1, (fileData(n).LastIndexOf("]") - 1)), n)
            End If
        Next
        packet.Add(endPackedID, fileData.Length)

        Return packet
    End Function

    Shared Function GetPacketName(ByVal packet As String) As String
        Dim n As String = packet.IndexOf("|")
        If n <> -1 Then
            Return packet.Remove(0, n + 2)
        End If

        Return packet
    End Function

    ' получает данные имен и номеров из AIPACKET  
    Shared Function GetAllAIPacketNumber(ByRef Path As String) As SortedList(Of String, Integer)
        Dim packet As SortedList(Of String, Integer) = New SortedList(Of String, Integer)
        Dim fileData As String() = File.ReadAllLines(Path)
        Dim debris As Char() = {"[", "]", " "}

        For Each line In fileData
            If line.StartsWith("[") Then
                Dim name As String = line.Trim(debris) '.Substring(1, (line.LastIndexOf("]") - 1))
                Dim num As Integer = GetIniParam(name, "packet_num", Path)
                packet.Add(String.Format("{0} ({1})", name, num), num)
            End If
        Next
        Return packet
    End Function
End Class
