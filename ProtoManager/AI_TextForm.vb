﻿Imports System.IO

Public Class AI_TextForm

    Private pathAI As String
    Private aiCustom As Boolean

    Private buffer As List(Of String)
    Private sPacketID As Integer
    Private ePacketID As Integer
    Private change As Boolean
    Private ownerSaveButton As Boolean

    Friend Sub New(ByVal sPacketID As Integer, ByVal ePacketID As Integer, ByRef path As String, ByVal ownerSaveButton As Boolean, ByVal aiCustom As Boolean)
        InitializeComponent()

        Me.sPacketID = sPacketID
        Me.ePacketID = ePacketID
        Me.ownerSaveButton = ownerSaveButton
        Me.aiCustom = aiCustom
        Me.pathAI = path

        buffer = File.ReadAllLines(path).ToList
        Dim str As String = String.Empty
        For n = sPacketID To ePacketID - 1
            str &= vbCrLf & buffer(n)
        Next
        TextBox1.Text = str.Remove(0, 2)
    End Sub

    Private Sub AI_TextForm_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        FormControl(Me.Owner)
    End Sub

    Private Sub AI_TextForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormControl(Me.Owner, True)
    End Sub

    Private Sub Close_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        If change Then
            AI_Form.ReloadFile(CType(Me.Owner, AI_Form))
            ownerSaveButton = False
        End If
        Me.Close()
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        buffer.RemoveRange(sPacketID, (ePacketID - sPacketID))
        buffer.InsertRange(sPacketID, Split(TextBox1.Text, vbCrLf).ToList)

        ' save to file
        Dim sFile As String = If(aiCustom, pathAI, SaveMOD_Path & AI.AIFILE)

        Dim dir = Path.GetDirectoryName(sFile)
        If (Directory.Exists(dir) = False) Then Directory.CreateDirectory(dir)

        File.WriteAllLines(sFile, buffer)

        change = True
        If (Main.PacketAI IsNot Nothing) Then Main.PacketAI.Clear()
        Main.PrintLog("Update AI: " & sFile)
    End Sub

    ' Рекурсивный перебор контролов класса формы
    Private Sub FormControl(сntr As Control, Optional def As Boolean = False)
        For Each _control As Control In сntr.Controls
            If (TypeOf _control Is GroupBox) Then
                FormControl(_control, def)
            ElseIf Not (TypeOf _control Is Label) Then
                If _control.Name = "SaveButton" Then
                    _control.Enabled = ownerSaveButton
                Else
                    _control.Enabled = def
                End If
            End If
        Next
    End Sub

End Class