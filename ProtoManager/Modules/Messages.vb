﻿Imports System.IO
Imports System.Text

Friend Module Messages

    Private Const TEXT_GAME As String = "\Text\English\Game\"

    'Содержимое считанного msg файла
    Friend MSG_DATATEXT() As String

    Friend Sub SaveMSGFile(ByVal msgFile As String)
        Dim savePath As String = SaveMOD_Path & TEXT_GAME

        If Not Directory.Exists(savePath) Then Directory.CreateDirectory(savePath)
        File.WriteAllLines(savePath & msgFile, MSG_DATATEXT, Settings.MsgEncoding)

        'Log
        Main.PrintLog("Update file: " & savePath & msgFile)
    End Sub

    'Считывает содержимое из msg файла в массив с соответствующей выбранной кодировкой
    Friend Sub GetMsgData(ByRef msgFile As String, Optional ByVal ToLV As Boolean = True)
        MSG_DATATEXT = File.ReadAllLines(DatFiles.CheckFile(TEXT_GAME & msgFile), Settings.MsgEncoding)
        If txtLvCp And ToLV Then EncodingLevCorp()
    End Sub

    'Возвращает Имя криттера или его описание
    Friend Function GetNameObject(ByRef NameID As Integer) As String
        If NameID = 0 Then Return String.Empty
        Dim strLine As String, sNameID As String = CStr(NameID)

        'Ищем строку с номером NameID
        For i As Integer = 0 To UBound(MSG_DATATEXT)
            strLine = MSG_DATATEXT(i)
            If GetParamMsg(strLine) = sNameID Then Return GetParamMsg(strLine, True)
        Next

        Return String.Empty
    End Function

    'Возвращает параметры из строки формата Msg  
    Friend Function GetParamMsg(ByRef str As String, Optional ByRef strValue As Boolean = False) As String
        If str.Length < 2 Then Return Nothing
        Dim n As Integer = str.IndexOf("}", 1)
        If n = -1 Then Return Nothing

        'Извлекаем
        If strValue = True Then
            n = str.IndexOf("{", n + 2)
            Return str.Substring(n + 1, str.Length - (n + 2))
        Else
            Return str.Substring(1, n - 1)
        End If
    End Function

    'Возвращает номер строки массива msg-файла
    Friend Function GetMSGLine(ByRef NameID As Integer) As Integer
        Dim strLine As String, sNameID As String = CStr(NameID)

        'Ищем строку с номером NameID
        For i As Integer = 0 To UBound(MSG_DATATEXT)
            strLine = MSG_DATATEXT(i)
            If GetParamMsg(strLine) = sNameID Then Return i
        Next

        Return -1
    End Function

    'Добавление или измнение значения в MSG файле
    Friend Function AddTextMSG(ByVal str As String, ByVal ID As Integer, ByVal desc As Boolean) As Boolean
        Dim nline As Integer = GetMSGLine(ID)

        If nline = -1 And desc Then Return True
        If nline = -1 Then
            ReDim Preserve MSG_DATATEXT(UBound(MSG_DATATEXT) + 1)
            nline = UBound(MSG_DATATEXT)
        End If

        If desc Then
            ID += 1
            nline += 1
            ' desc line
            If GetMSGLine(ID) = -1 Then
                ReDim Preserve MSG_DATATEXT(UBound(MSG_DATATEXT) + 1)
                For n As Integer = UBound(MSG_DATATEXT) To nline Step -1
                    MSG_DATATEXT(n) = MSG_DATATEXT(n - 1)
                Next
            End If
        End If

        If txtLvCp Then CodingToLevCorp(str)
        MSG_DATATEXT(nline) = "{" & ID & "}{}{" & str & "}"

        Return False
    End Function

    Friend Sub EncodingLevCorp()
        Dim data() As Byte

        For n As Integer = 0 To UBound(MSG_DATATEXT)
            If MSG_DATATEXT(n).StartsWith("{") Then
                data = Encoding.GetEncoding("cp866").GetBytes(MSG_DATATEXT(n))
                For m As Integer = 0 To UBound(data)
                    If data(m) >= &HB0 And data(m) <= &HBF Then
                        data(m) = data(m) + &H30
                    End If
                Next
                MSG_DATATEXT(n) = Encoding.GetEncoding("cp866").GetString(data)
            End If
        Next
    End Sub

    ' Конвертирует передаваемую строку в кодировку LevCorp
    Private Sub CodingToLevCorp(ByRef str As String)
        Dim data() As Byte = Encoding.GetEncoding("cp866").GetBytes(str)

        For m As Integer = 0 To UBound(data)
            If data(m) >= &HE0 And data(m) <= &HEF Then
                data(m) = data(m) - &H30
            End If
        Next
        str = Encoding.GetEncoding("cp866").GetString(data)
    End Sub

End Module
