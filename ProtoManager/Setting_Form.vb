Imports System.Windows.Forms
Imports System.IO
Imports System.Text

Friend Class Setting_Form

    'Private Const SC_CLOSE As Int32 = &HF060
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    ' Нажата кнопка Close.
    '    If m.WParam.ToInt32() = SC_CLOSE Then SettingExt = True
    '    MyBase.WndProc(m)
    'End Sub

    Friend settingExit, fRun As Boolean
    Private pathChange As Boolean = False

    Private Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OK_Button.Click
        Dim restart As Boolean = False

        Settings.txtWin = RadioButton1.Checked
        Settings.txtLvCp = RadioButton3.Checked
        Settings.proRO = CheckBox2.Checked
        Settings.cCache = CheckBox3.Checked
        Settings.ExtractBack = CheckBox1.Checked

        If TextBox1.Text <> String.Empty AndAlso (File.Exists(TextBox1.Text & MasterDAT) OrElse Directory.Exists(TextBox1.Text & MasterDAT)) Then
            TextBox2.Text = TextBox2.Text.Trim
            If TextBox2.Text = String.Empty Then
                TextBox2.Text = TextBox1.Text & DIR_DATA
            End If
            If fRun Then
                Game_Path = TextBox1.Text
                GameDATA_Path = Game_Path & DIR_DATA
                SaveMOD_Path = TextBox2.Text
            ElseIf pathChange Then
                File.Delete(Cache_Patch & "\cache.id")
                If MsgBox("New paths will take effect after restarting editor." & vbLf & "Restart the editor now?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    restart = True
                End If
            End If
            settingExit = False
            Settings.gPath = TextBox1.Text
            Settings.sPath = TextBox2.Text
            Settings.Save_Config()
            pathChange = False
            Me.Hide()
            If restart Then
                Main_Form.Dispose()
                Application.Exit()
                Application.Restart()
            End If
        Else
            MsgBox("The master.dat file or folder can not be found. Set the path to folder game Fallout 2.")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        FolderBrowserDialog1.Description = "Select the folder where the files are located and Master.dat Critter.dat"
        FolderBrowserDialog1.ShowDialog()
        If FolderBrowserDialog1.SelectedPath <> String.Empty Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
            If TextBox2.Text.Trim = String.Empty And TextBox1.Text <> String.Empty Then
                TextBox2.Text = TextBox1.Text & DIR_DATA
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        FolderBrowserDialog1.Description = "Select the folder in which the editor will save the edited file"
        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        If FolderBrowserDialog1.SelectedPath <> String.Empty Then TextBox2.Text = FolderBrowserDialog1.SelectedPath
        FolderBrowserDialog1.ShowNewFolderButton = False
    End Sub

    Private Sub Setting_Form_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        Main_Form.Focus()
        'Me.Dispose()
    End Sub

    Private Sub Setting_Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        If TextBox1.Text = String.Empty Then TextBox1.Text = Game_Path
        If TextBox2.Text = String.Empty Then TextBox2.Text = SaveMOD_Path
        If HEX_Path <> String.Empty Then
            TextBox3.Enabled = True
            TextBox3.Text = HEX_Path
        End If

        RadioButton2.Checked = Not txtWin
        RadioButton3.Checked = txtLvCp
        CheckBox2.Checked = proRO
        CheckBox3.Checked = cCache
        CheckBox1.Checked = ExtractBack
        pathChange = False
    End Sub

    Private Sub Setting_Form_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        'If e.CloseReason = CloseReason.UserClosing Then SettingExt = True
        e.Cancel = True
        Me.Hide()

        If pathChange Then
            TextBox2.Text = SaveMOD_Path
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Clear_Art_Cache()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If TextBox1.Text.Trim.ToLower <> Game_Path.ToLower OrElse TextBox2.Text.Trim.ToLower <> SaveMOD_Path.ToLower Then pathChange = True
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        OpenFileDialog1.InitialDirectory = WorkAppDIR
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> String.Empty Then
            TextBox3.Text = OpenFileDialog1.FileName
            HEX_Path = OpenFileDialog1.FileName
            TextBox3.Enabled = True
        End If
    End Sub

End Class
