Imports System.IO

Public Class AI_TextForm

    Private buffer As List(Of String)
    Private sPacketID As Integer
    Private ePacketID As Integer
    Private change As Boolean
    Private ownerSaveButton As Boolean

    Friend Sub New(ByVal sPacketID As Integer, ByVal ePacketID As Integer, ByRef path As String, ByVal ownerSaveButton As Boolean)
        InitializeComponent()

        Me.sPacketID = sPacketID
        Me.ePacketID = ePacketID
        Me.ownerSaveButton = ownerSaveButton

        buffer = File.ReadAllLines(path).ToList
        Dim str As String = String.Empty
        For n = sPacketID To ePacketID - 1
            str &= vbLf & buffer(n)
        Next
        RichTextBox1.Text = str.Remove(0, 1)
    End Sub

    Private Sub AI_TextForm_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        FormControl(Me.Owner)
    End Sub

    Private Sub AI_TextForm_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormControl(Me.Owner, True)
    End Sub

    Private Sub Close_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        If change Then
            AI_Form.ReloadFile(Me.Owner)
            ownerSaveButton = False
        End If
        Me.Close()
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        buffer.RemoveRange(sPacketID, (ePacketID - sPacketID))
        buffer.InsertRange(sPacketID, Split(RichTextBox1.Text, vbLf).ToList)
        ' save to file
        File.WriteAllLines(SaveMOD_Path & AI.AIFILE, buffer)
        change = True
    End Sub

    ' Рекурсивный перебор контролов класса формы
    Private Sub FormControl(ByRef сntr As Control, Optional ByRef def As Boolean = False)
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