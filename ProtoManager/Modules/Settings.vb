Imports System.IO
Imports System.Text

Friend Module Settings

    Friend ReadOnly WorkAppDIR As String = Application.StartupPath
    Friend ReadOnly Cache_Patch As String = WorkAppDIR & "\Cache"

    Friend Game_Path As String 'Папка игры
    Friend GameDATA_Path As String 'Папка DATA игры
    Friend SaveMOD_Path As String 'Папка в которую сохраняются отредактированные файлы.
    Friend HEX_Path As String

    Friend ReadOnly defaultHEX As String = WorkAppDIR & "\hex\frhed.exe"

    Friend gPath, sPath As String

    ' Program sets
    Friend SplitSize As Integer = 950 'default size
    Friend txtWin As Boolean
    Friend txtLvCp As Boolean
    Friend proRO As Boolean
    Friend cCache As Boolean
    Friend cArtCache As Boolean
    Friend ExtractBack As Boolean
    Friend HoverSelect As Boolean

    Friend ShowAIPacket As Boolean = True
    Friend SortedAIPacket As Boolean

    Friend MsgEncoding As Encoding

    Friend Sub SetDoubleBuffered(ByVal control As Control)
        Dim doubleBufferPropertyInfo = control.GetType().GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        doubleBufferPropertyInfo.SetValue(control, True, Nothing)
    End Sub

    Friend Sub SetEncoding()
        If txtWin Then
            MsgEncoding = Encoding.Default
        Else
            MsgEncoding = Encoding.GetEncoding("cp866")
        End If
    End Sub

    'Load from ini
    Friend Sub Get_Config()
        Dim strIni As String = String.Empty

        Dim ifile As StreamReader = File.OpenText(WorkAppDIR & "\config.ini")
        Try
            Do Until ifile.EndOfStream
                strIni = ifile.ReadLine
                Select Case strIni
                    Case "[Path]"
                        Game_Path = ifile.ReadLine.Substring(11).Trim
                        SaveMOD_Path = ifile.ReadLine.Substring(8).Trim
                        HEX_Path = ifile.ReadLine.Substring(8).Trim
                    Case "[Option]"
                        proRO = CBool(ifile.ReadLine.Substring(9))
                        txtWin = CBool(ifile.ReadLine.Substring(7))
                        txtLvCp = CBool(ifile.ReadLine.Substring(6))
                        cCache = CBool(ifile.ReadLine.Substring(11))
                        cArtCache = CBool(ifile.ReadLine.Substring(14))
                        ExtractBack = CBool(ifile.ReadLine.Substring(11)) 'Background=
                        HoverSelect = CBool(ifile.ReadLine.Substring(12))
                        SplitSize = CInt(ifile.ReadLine.Substring(10))
                End Select
            Loop
        Catch ex As Exception
            GoTo SetDefConf
        Finally
            ifile.Close()
        End Try
        
        If Game_Path = String.Empty Then GoTo SetDefConf
        GameDATA_Path = Game_Path & DIR_DATA
        If SaveMOD_Path = String.Empty Then SaveMOD_Path = GameDATA_Path
        gPath = Game_Path
        sPath = SaveMOD_Path
        Exit Sub

SetDefConf:
        proRO = True : txtWin = True : cCache = True : ExtractBack = True
        Setting_Form.fRun = True
        Setting_Form.settingExit = True
        SplashScreen.TopMost = False
        Setting_Form.ShowDialog()
        SplashScreen.TopMost = True
    End Sub

    'Save to ini
    Friend Sub Save_Config()
        Dim AppSetting As New List(Of String)
        AppSetting.Add("[Path]")
        AppSetting.Add("CommonPath=" & gPath)
        AppSetting.Add("ModPath=" & sPath)
        AppSetting.Add("HexPath=" & HEX_Path)
        AppSetting.Add(String.Empty)
        AppSetting.Add("[Option]")
        AppSetting.Add("ReadOnly=" & proRO)
        AppSetting.Add("MsgWIN=" & txtWin)
        AppSetting.Add("MsgLC=" & txtLvCp)
        AppSetting.Add("ClearCache=" & cCache)
        AppSetting.Add("ClearArtCache=" & cArtCache)
        AppSetting.Add("Background=" & ExtractBack)
        AppSetting.Add("HoverSelect=" & HoverSelect)
        If Main_Form.WindowState = FormWindowState.Maximized Then
            AppSetting.Add("SplitSize=" & Main_Form.SplitContainer1.SplitterDistance)
        Else
            AppSetting.Add("SplitSize=" & SplitSize)
        End If
        '
        File.WriteAllLines(WorkAppDIR & "\config.ini", AppSetting)
    End Sub

    Friend Sub Clear_Cache()
        On Error Resume Next
        Directory.Delete(Cache_Patch & "\proto", True)
        Directory.Delete(Cache_Patch & "\data", True)
        Directory.Delete(Cache_Patch & "\scripts", True)
        Directory.Delete(Cache_Patch & "\text", True)
        File.Delete(Cache_Patch & "\cache.id")
        Clear_Art_Cache()
    End Sub

    Friend Sub Clear_Art_Cache()
        On Error Resume Next
        Directory.Delete(Cache_Patch & "\art", True)
    End Sub

End Module
