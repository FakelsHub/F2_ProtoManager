Imports System.IO
Imports System.Text

Friend Module Settings

    Private configMap As Dictionary(Of String, String)

    Friend ReadOnly WorkAppDIR As String = Application.StartupPath
    Friend ReadOnly Cache_Patch As String = WorkAppDIR & "\Cache"

    Friend Game_Path As String 'Папка игры
    Friend GameDATA_Path As String 'Папка DATA игры
    Friend SaveMOD_Path As String 'Папка в которую сохраняются отредактированные файлы.
    Friend HEX_Path As String

    Friend ReadOnly defaultHEX As String = WorkAppDIR & "\hex\frhed.exe"

    Friend gPath, sPath As String
    Friend msgLangPath As String = "english"

    ' Program sets
    Friend SplitSize As Integer = -1 'default size
    Friend txtWin As Boolean = True
    Friend txtLvCp As Boolean = False
    Friend proRO As Boolean = True
    Friend cCache As Boolean = True
    Friend cArtCache As Boolean = False
    Friend HoverSelect As Boolean = True

    Friend ShowAIPacket As Boolean = True
    Friend SortedAIPacket As Boolean

    Friend ColumnItemSize(4) As Integer
    Friend ColumnCritterSize(4) As Integer

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
        Dim temp() As String
        Dim ifile As StreamReader = File.OpenText(WorkAppDIR & "\config.ini")
        configMap = New Dictionary(Of String, String)
        Try
            Do Until ifile.EndOfStream
                temp = Split(ifile.ReadLine, "=")
                If temp.Length > 1 Then configMap.Add(temp(0).Trim, temp(1).Trim)
            Loop
        Catch ex As Exception
            GoTo SetDefConf
        Finally
            ifile.Close()
        End Try

        Dim strIni As String = String.Empty
        If (configMap.TryGetValue("CommonPath", strIni) = False) Then GoTo SetDefConf
        Game_Path = strIni
        If Game_Path = String.Empty Then GoTo SetDefConf
        GameDATA_Path = Game_Path & DIR_DATA

        If (configMap.TryGetValue("ModPath", strIni) = False) Then GoTo SetDefConf
        SaveMOD_Path = strIni
        If SaveMOD_Path = String.Empty Then SaveMOD_Path = GameDATA_Path

        If (configMap.TryGetValue("HexPath", strIni)) Then HEX_Path = strIni
        If (configMap.TryGetValue("LangPath", strIni)) Then msgLangPath = strIni

        If (configMap.TryGetValue("ReadOnly", strIni)) Then proRO = CBool(strIni)
        If (configMap.TryGetValue("MsgWIN", strIni)) Then txtWin = CBool(strIni)
        If (configMap.TryGetValue("MsgLC", strIni)) Then txtLvCp = CBool(strIni)
        If (configMap.TryGetValue("ClearCache", strIni)) Then cCache = CBool(strIni)
        If (configMap.TryGetValue("ClearArtCache", strIni)) Then cArtCache = CBool(strIni)
        If (configMap.TryGetValue("HoverSelect", strIni)) Then HoverSelect = CBool(strIni)
        If (configMap.TryGetValue("StatFormula", strIni)) Then CalcStats.SetFormula(Convert.ToInt32(strIni))

        If (configMap.TryGetValue("SplitSize", strIni)) Then SplitSize = CInt(strIni)
        Dim i As Integer = 0
        If (configMap.TryGetValue("ColumnIt", strIni)) Then
            temp = Split(strIni, ",")
            For Each size As String In temp
                ColumnItemSize(i) = CInt(size.Trim())
                i += 1
                If i > 4 Then Exit For
            Next
        End If
        If (configMap.TryGetValue("ColumnCr", strIni)) Then
            temp = Split(strIni, ",")
            i = 0
            For Each size As String In temp
                ColumnCritterSize(i) = CInt(size.Trim())
                i += 1
                If i > 4 Then Exit For
            Next
        End If

        gPath = Game_Path
        sPath = SaveMOD_Path

        Messages.SetMessageLangPath()
        Exit Sub

SetDefConf:
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
        AppSetting.Add("LangPath=" & msgLangPath)
        AppSetting.Add(String.Empty)
        AppSetting.Add("[Option]")
        AppSetting.Add("ReadOnly=" & proRO)
        AppSetting.Add("MsgWIN=" & txtWin)
        AppSetting.Add("MsgLC=" & txtLvCp)
        AppSetting.Add("ClearCache=" & cCache)
        AppSetting.Add("ClearArtCache=" & cArtCache)
        AppSetting.Add("Background=")
        AppSetting.Add("HoverSelect=" & HoverSelect)
        AppSetting.Add("StatFormula=" & CalcStats.GetFormula().ToString)
        AppSetting.Add(String.Empty)
        AppSetting.Add("[Size]")
        If Main_Form.WindowState = FormWindowState.Maximized Then
            AppSetting.Add("SplitSize=" & Main_Form.SplitContainer1.SplitterDistance)
        Else
            AppSetting.Add("SplitSize=" & SplitSize)
        End If
        '
        For i As Integer = 0 To Main_Form.ListView1.Columns.Count - 1
            ColumnCritterSize(i) = Main_Form.ListView1.Columns(i).Width
        Next
        For i As Integer = 0 To Main_Form.ListView2.Columns.Count - 1
            ColumnItemSize(i) = Main_Form.ListView2.Columns(i).Width
        Next
        AppSetting.Add("ColumnIt=" & String.Join(",", ColumnItemSize))
        AppSetting.Add("ColumnCr=" & String.Join(",", ColumnCritterSize))
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
