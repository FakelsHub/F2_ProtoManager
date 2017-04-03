Public Class AI_TextForm

    Friend Sub Initialize(ByVal sPacketID As Integer, ByVal ePacketID As Integer, ByRef path As String)
        Dim buffer() As String = IO.File.ReadAllLines(path)
        For n As Integer = sPacketID To ePacketID - 1
            RichTextBox1.Text &= buffer(n) & vbCrLf
        Next
    End Sub

    Private Sub AI_TextForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormControl(Me.Owner)
    End Sub

    Private Sub AI_TextForm_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        FormControl(Me.Owner, True)
    End Sub

    ' Рекурсивный перебор контролов класса формы
    Private Sub FormControl(ByRef сntr As Control, Optional ByRef def As Boolean = False)
        For Each _control As Control In сntr.Controls
            If (TypeOf _control Is GroupBox) Then
                FormControl(_control, def)
            ElseIf Not (TypeOf _control Is Label) Then
                _control.Enabled = def
            End If
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class