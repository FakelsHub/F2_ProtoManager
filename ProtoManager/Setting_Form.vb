Imports System.IO

Friend Class Setting_Form

    'Private Const SC_CLOSE As Int32 = &HF060
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    ' Нажата кнопка Close.
    '    If m.WParam.ToInt32() = SC_CLOSE Then SettingExt = True
    '    MyBase.WndProc(m)
    'End Sub

    Friend settingExit, firstRun, updateList As Boolean

    Private Sub Setting_Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        tbMainPath.Text = Settings.Game_Config
        tbModDataPath.Text = Settings.SaveMOD_Path

        If HEX_Path <> String.Empty Then
            TextBox3.Enabled = True
            TextBox3.Text = HEX_Path
        Else
            TextBox3.Text = defaultHEX
        End If

        RadioButton2.Checked = Not Settings.txtWin
        RadioButton3.Checked = Settings.txtLvCp
        CheckBox2.Checked = Settings.proRO
        CheckBox3.Checked = Settings.cCache
        'CheckBox1.Checked = ExtractBack
        tbMsgLang.Text = Settings.languagePath

        lstCheckMods.Items.Clear()
        For Each el As ExtraModData In DatFiles.extraMods
            lstCheckMods.Items.Add(el.filePath, el.isEnabled)
        Next

        '' обновление списка
        'Dim folder = Path.GetDirectoryName(tbMainPath.Text) + "\"
        'If (File.Exists(folder)) Then
        '    GameConfig.SearchDatFiles(folder, DatFiles.MasterDAT, DatFiles.CritterDAT)
        '    For Each el As ListMods In GameConfig.listMods
        '        For Each item As ListMods In DatFiles.mods
        '            If el.modPath.Equals(item.modPath, StringComparison.OrdinalIgnoreCase) Then
        '                el.isEnabled = item.isEnabled
        '                Exit For
        '            End If
        '        Next
        '    Next
        '    lstCheckMods.Items.Clear()
        '    For Each el As ListMods In GameConfig.listMods
        '        lstCheckMods.Items.Add(el.modPath, el.isEnabled)
        '    Next
        'End If
    End Sub

    Private Sub Setting_Form_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        'Main_Form.Focus()
        updateList = False 'Me.Dispose()
    End Sub

    Private Sub ApplySetting(ByVal sender As Object, ByVal e As EventArgs) Handles btnOK.Click
        Dim restart As Boolean = False

        Dim config = tbMainPath.Text.Trim
        Dim savePath = tbModDataPath.Text.Trim
        Dim pathChange = (config.ToLower <> Game_Config.ToLower) OrElse (savePath.ToLower <> SaveMOD_Path.ToLower)

        UpdateExtraMods()

        Settings.txtWin = RadioButton1.Checked
        Settings.txtLvCp = RadioButton3.Checked
        Settings.proRO = CheckBox2.Checked
        Settings.cCache = CheckBox3.Checked
        'Settings.ExtractBack = CheckBox1.Checked

        If (pathChange AndAlso config <> String.Empty) Then
            Dim master = GameConfig.gameFolder + GameConfig.masterDatPath
            If (File.Exists(master) OrElse Directory.Exists(master)) Then

                If savePath = String.Empty Then
                    savePath = GameConfig.gameFolder & GameConfig.m_patches
                End If

                If firstRun Then
                    Game_Config = config
                    Game_Path = GameConfig.gameFolder
                    GameDATA_Path = GameConfig.gameFolder & GameConfig.m_patches
                    SaveMOD_Path = savePath
                ElseIf (pathChange) Then
                    File.Delete(Cache_Patch & "\cache.id")
                    If MsgBox("The new set path will take effect after the program is restarted." & vbLf & "Restart the program now?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        restart = True
                    End If
                End If

                settingExit = False

                Settings.languagePath = GameConfig.GameLanguage
                Settings.sConfigPath = config
                Settings.sSaveFolderPath = savePath
                Settings.SaveConfigFile()

                If (restart) Then
                    Main_Form.Dispose()
                    Application.Exit()
                    Application.Restart()
                Else
                    Me.Close()
                End If
            Else
                MessageBox.Show(String.Format("The {0} file or folder cannot be found. Set the correct path to the game folder or file.", GameConfig.masterDatPath))
            End If
        Else
            Settings.languagePath = tbMsgLang.Text.Trim
            Settings.SaveConfigFile()
            Messages.SetMessageLangPath(Settings.languagePath)
            Me.Close()
        End If
    End Sub

    Private Sub UpdateExtraMods()
        If (updateList) Then DatFiles.extraMods.Clear()
        If (DatFiles.extraMods.Count = 0) Then
            For Each el As ExtraModData In GameConfig.gcExtraMods
                DatFiles.AddExtraMod(el)
            Next
        End If

        For i As Integer = 0 To lstCheckMods.Items.Count - 1
            If (DatFiles.extraMods(i).filePath.Equals(lstCheckMods.Items(i).ToString, StringComparison.OrdinalIgnoreCase) = False) Then
                ' найти положение sfall 
                For n = 0 To DatFiles.extraMods.Count - 1
                    If (DatFiles.extraMods(i).filePath.Equals(lstCheckMods.Items(n).ToString, StringComparison.OrdinalIgnoreCase)) Then
                        Dim temp = DatFiles.extraMods(i)
                        temp.isEnabled = lstCheckMods.GetItemChecked(n)
                        DatFiles.extraMods.RemoveAt(i)
                        DatFiles.extraMods.Insert(n, temp)
                        Exit For
                    End If
                Next
            Else
                DatFiles.extraMods(i).isEnabled = lstCheckMods.GetItemChecked(i)
            End If
        Next
    End Sub

    Private Sub OpenConfig(ByVal sender As Object, ByVal e As EventArgs) Handles btnConfig.Click
        OpenFileDialog1.Title = "Select the folder and the game configuration file."
        OpenFileDialog1.Filter = "Config file|*.cfg"

        If (tbMainPath.Text <> String.Empty) Then OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(tbMainPath.Text) + "\"

        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            Dim cfgFile = OpenFileDialog1.FileName
            tbMainPath.Text = cfgFile

            GameConfig.ReadGameConfig(cfgFile)
            GameConfig.SearchExtraModFiles(Settings.Game_Path, DatFiles.MasterDAT, DatFiles.CritterDAT)
            tbMsgLang.Text = GameConfig.GameLanguage()

            If tbModDataPath.Text.Trim = String.Empty Then
                tbModDataPath.Text = GameConfig.gameFolder + GameConfig.m_patches
            End If

            lstCheckMods.Items.Clear()
            For Each item In GameConfig.gcExtraMods
                lstCheckMods.Items.Add(item.filePath)
            Next
        End If
    End Sub

    Private Sub ChangeSaveFolder(ByVal sender As Object, ByVal e As EventArgs) Handles btnSaveFolder.Click
        'FolderBrowserDialog1.Description = "Select the folder where the editor will save the edited files."
        If (tbModDataPath.Text <> String.Empty) Then FolderBrowserDialog1.SelectedPath = tbModDataPath.Text

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            If FolderBrowserDialog1.SelectedPath <> String.Empty Then tbModDataPath.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub ClearArtCache(ByVal sender As Object, ByVal e As EventArgs) Handles btnClear.Click
        Clear_Art_Cache()
    End Sub

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        Dim i = lstCheckMods.SelectedIndex
        If (i = 0) Then Return

        Dim isCheck = lstCheckMods.GetItemChecked(i)
        lstCheckMods.Items.Insert(i - 1, lstCheckMods.Items(i))
        lstCheckMods.Items.RemoveAt(i + 1)
        i -= 1
        If (isCheck) Then lstCheckMods.SetItemChecked(i, True)
        lstCheckMods.SetSelected(i, True)
    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click
        Dim i = lstCheckMods.SelectedIndex
        If (i = lstCheckMods.Items.Count - 1) Then Return

        Dim isCheck = lstCheckMods.GetItemChecked(i)
        lstCheckMods.Items.Insert(i + 2, lstCheckMods.Items(i))
        lstCheckMods.Items.RemoveAt(i)
        i += 1
        If (isCheck) Then lstCheckMods.SetItemChecked(i, True)
        lstCheckMods.SetSelected(i, True)
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        GameConfig.ReadGameConfig(tbMainPath.Text.Trim)
        GameConfig.SearchExtraModFiles(Settings.Game_Path, DatFiles.MasterDAT, DatFiles.CritterDAT)
        lstCheckMods.Items.Clear()
        For Each item In GameConfig.gcExtraMods
            lstCheckMods.Items.Add(item.filePath)
        Next
        updateList = True
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        OpenFileDialog1.Title = ""
        OpenFileDialog1.Filter = "Execute file|*.exe"
        OpenFileDialog1.InitialDirectory = WorkAppDIR

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox3.Text = OpenFileDialog1.FileName
            HEX_Path = OpenFileDialog1.FileName
            TextBox3.Enabled = True
        End If
    End Sub

End Class
