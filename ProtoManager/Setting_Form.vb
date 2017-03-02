Imports System.Windows.Forms

Public Class Setting_Form

    'Private Const SC_CLOSE As Int32 = &HF060
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    ' Нажата кнопка Close.
    '    If m.WParam.ToInt32() = SC_CLOSE Then SettingExt = True
    '    MyBase.WndProc(m)
    'End Sub

    Private pChange As Boolean

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        txtWin = RadioButton1.Checked
        txtLvCp = RadioButton3.Checked
        proRO = CheckBox2.Checked
        cCache = CheckBox3.Checked
        ExtractBack = CheckBox1.Checked
        If Trim(TextBox1.Text) <> Nothing AndAlso (My.Computer.FileSystem.FileExists(TextBox1.Text & "\master.dat") = True OrElse My.Computer.FileSystem.DirectoryExists(TextBox1.Text & "\master.dat") = True) Then
            If TextBox2.Text = Nothing Then TextBox2.Text = TextBox1.Text & "\Data" : TextBox2.ReadOnly = True
            If fRun Then
                Game_Path = TextBox1.Text
                GameDATA_Path = Game_Path & "\Data" '"\MyTestData"
                If TextBox2.ReadOnly Then
                    SaveMOD_Path = GameDATA_Path
                Else
                    SaveMOD_Path = TextBox2.Text
                End If
            ElseIf pChange Then
                My.Computer.FileSystem.DeleteFile(Cache_Patch & "\cache.id")
                MsgBox("New installation paths will take effect after restarting the editor.")
            End If
            Store_Ini()
            SettingExt = False
            Main_Form.Cp886ToolStripMenuItem.Checked = Not (txtWin) And Not (txtLvCp)
            Main_Form.ClearToolStripMenuItem2.Checked = cArtCache
            Main_Form.AttrReadOnlyToolStripMenuItem.Checked = proRO
            Main_Form.LinkLabel1.Text = GameDATA_Path
            Main_Form.LinkLabel2.Text = SaveMOD_Path
            Me.Hide() 'Me.Dispose()
        Else
            MsgBox("Not found master.dat file. Set the path to the folder Fallout.")
        End If
    End Sub

    'Store to ini
    Friend Sub Store_Ini()
        Dim AppSetting(12) As String
        AppSetting(0) = "[Path]"
        AppSetting(1) = "CommonPath=" & TextBox1.Text
        If TextBox2.ReadOnly Then AppSetting(2) = "ModPath=" Else AppSetting(2) = "ModPath=" & TextBox2.Text
        If Not (TextBox3.Enabled) Then AppSetting(3) = "HexPath=" Else AppSetting(3) = "HexPath=" & TextBox3.Text
        AppSetting(4) = Nothing
        AppSetting(5) = "[Option]"
        AppSetting(6) = "ReadOnly=" & proRO
        AppSetting(7) = "MsgWIN=" & txtWin
        AppSetting(8) = "MsgLC=" & txtLvCp
        AppSetting(9) = "ClearCache=" & cCache
        AppSetting(10) = "ClearArtCache=" & cArtCache
        AppSetting(11) = "Background=" & ExtractBack
        If Main_Form.Visible Then AppSetting(12) = "SplitSize=" & Main_Form.SplitContainer1.SplitterDistance
        '
        System.IO.File.WriteAllLines(Application.StartupPath & "\config.ini", AppSetting)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FolderBrowserDialog1.Description = "Select the folder where the files are located and Master.dat Critter.dat"
        FolderBrowserDialog1.ShowDialog()
        If FolderBrowserDialog1.SelectedPath <> "" Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
            If TextBox2.Text = Nothing And TextBox1.Text <> Nothing Then TextBox2.Text = TextBox1.Text & "\Data" : TextBox2.ReadOnly = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FolderBrowserDialog1.Description = "Select the folder in which the editor will save the edited file"
        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        If FolderBrowserDialog1.SelectedPath <> "" Then
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
            TextBox2.ReadOnly = False
        End If
        FolderBrowserDialog1.ShowNewFolderButton = False
    End Sub

    Private Sub Setting_Form_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Main_Form.Focus()
        'Me.Dispose()
    End Sub

    Friend Sub Setting_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If TextBox1.Text = Nothing Then TextBox1.Text = Game_Path
        If TextBox2.Text = Nothing Then TextBox2.Text = SaveMOD_Path
        If TextBox3.Enabled Then TextBox3.Text = HEX_Path
        RadioButton2.Checked = Not txtWin
        RadioButton3.Checked = txtLvCp
        CheckBox2.Checked = proRO
        CheckBox3.Checked = cCache
        CheckBox1.Checked = ExtractBack
        pChange = False
    End Sub

    Private Sub Setting_Form_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'If e.CloseReason = CloseReason.UserClosing Then SettingExt = True
        e.Cancel = True
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Clear_Art_Cache()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If TextBox1.Text <> Game_Path Or TextBox2.Text <> SaveMOD_Path Then pChange = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1.InitialDirectory = Application.StartupPath
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "" Then
            TextBox3.Text = OpenFileDialog1.FileName
            HEX_Path = OpenFileDialog1.FileName
            TextBox3.Enabled = True
        End If
    End Sub

End Class
