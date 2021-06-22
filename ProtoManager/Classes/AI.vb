﻿Imports System.IO

Friend NotInheritable Class AI

    Friend Const AIFILE As String = "\data\AI.txt"

    Friend Const Unknown As String = "<NotSpecified>"
    Friend Const endPackedID As String = "EOFLINE"


    ' получает массив всех имен AIPACKET
    Shared Function GetAllAIPackets(ByRef Path As String) As Dictionary(Of String, Integer)
        Dim packet As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()
        Dim fileData As String() = File.ReadAllLines(Path)

        For n = 0 To fileData.Length - 1
            If fileData(n).StartsWith("[") Then
                Dim name = fileData(n).TrimEnd().Substring(1, (fileData(n).IndexOf("]") - 1))
                If packet.ContainsKey(name) Then
                    MessageBox.Show("The AI.txt file contains duplicate name of AI package: " + name + vbLf _
                                    + "You must correct the value of this name to another unique name.", "Warning: Duplicate name")
                Else
                    packet.Add(name, n)
                End If
            End If
        Next
        packet.Add(endPackedID, fileData.Length)

        Return packet
    End Function

    Shared Function GetPacketName(ByVal packet As String) As String
        Dim n = packet.IndexOf("|")
        If n <> -1 Then
            Return packet.Remove(0, n + 2)
        End If

        Return packet
    End Function

    ' получает данные имен и номеров из AIPACKET  
    Shared Function GetAllAIPacketNumber(ByRef Path As String) As SortedList(Of String, Integer)
        Dim packet As SortedList(Of String, Integer) = New SortedList(Of String, Integer)
        Dim fileData As String() = File.ReadAllLines(Path)

        For Each line In fileData
            If line.StartsWith("[") Then
                Dim name As String = line.Substring(1, (line.IndexOf("]") - 1))
                Dim num As Integer = INIFile.GetInt(name, "packet_num", Path)
                name = String.Format("{0} ({1})", name, num)
                If Not (packet.ContainsKey(name)) Then packet.Add(name, num)
            End If
        Next
        Return packet
    End Function
End Class
