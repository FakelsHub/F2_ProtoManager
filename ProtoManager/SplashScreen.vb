Public NotInheritable Class SplashScreen

    Private Sub SplashScreen_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Label3.Text &= My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        Get_Config()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Label1.Visible = True
        ProgressBar1.Visible = True
        Application.DoEvents()
        Main_Module.Main()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If cCache Then Clear_Cache()
        Application.Exit()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.TopMost = False
        Setting_Form.ShowDialog()
        Me.TopMost = True
    End Sub
End Class
