Public Class AI_Form
    Private Const _width As Integer = 287

    Private Sub AI_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width -= _width
    End Sub

    Private Sub Taunts_Show(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If sender.tag = 0 Then
            Me.Width += _width
            sender.tag = 1
            sender.Image = My.Resources.LeftArrow
        Else
            Me.Width -= _width
            sender.tag = 0
            sender.Image = My.Resources.RightArrow
        End If
    End Sub

End Class