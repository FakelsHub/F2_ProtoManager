Option Explicit On
Imports System.IO
Imports System.Runtime.InteropServices
Imports Prototypes

Module Main_Module

    Friend Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer

    Friend Game_Path As String 'Папка игры
    Friend GameDATA_Path As String 'Папка DATA игры 
    Friend SaveMOD_Path As String 'Папка в которую сохраняются отредактированные файлы.
    '
    Const MasterDAT As String = "\master.dat"
    Const CritterDAT As String = "\critter.dat"
    '
    Friend Cache_Patch As String = Application.StartupPath & "\Cache"
    Friend HEX_Path As String = Application.StartupPath & "\hex\frhed.exe"

    Friend Current_Path As String
    '
    Friend Critter_LST() As String
    Friend Critter_NAME() As String
    Private Critters_FRM() As String
    Private Teams() As String
    Private AI() As String

    'общая структура куда записываются данный из про файла.
    Friend CritterPro As CritPro

    Friend Items_LST(0, 0) As String
    Friend Items_NAME() As String
    Private Items_FRM() As String
    Friend Iven_FRM() As String

    Friend AmmoPID() As Integer
    Friend AmmoNAME() As String
    Friend CaliberNAME() As String
    '
    Friend Misc_LST() As String
    Friend Misc_NAME() As String
    Friend Perk_NAME() As String
    '
    Private Scripts_Lst() As String

    Friend MSG_DATATEXT() As String

    Friend DmgType() As String = {"Normal", "Laser", "Fire", "Plasma", "Electrical", "EMP", "Explode"}

    Friend SplitSize As Integer = 600 'default size
    Friend txtWin As Boolean
    Friend txtLvCp As Boolean
    Friend proRO As Boolean
    Friend cCache As Boolean
    Friend cArtCache As Boolean
    Friend ExtractBack As Boolean

    Friend SettingExt, fRun As Boolean

    'Declare zlib functions "Compress" and "Uncompress" for compressing Byte Arrays
    '    <DllImport("zlib.DLL", EntryPoint:="compress")> _
    '    Private Function CompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    ' Leave function empty - DLLImport attribute forwards calls to CompressByteArray to compress in zlib.dLL
    '    End Function
    '    <DllImport("zlib.DLL", EntryPoint:="uncompress")> _
    '    Private Function UncompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    ' Leave function empty - DLLImport attribute forwards calls to UnCompressByteArray to Uncompress in zlib.dLL
    '   End Function

    'Инициализация...
    Friend Sub Main()
        Dim n As Short, noRe As Boolean, noId As Boolean
        Dim timeout As Integer = 10000

        'Get_Config()
        SplashScreen.ProgressBar1.Value = 10
        Application.DoEvents()
        If SettingExt = True Then GoTo ExitProg
        If Not (My.Computer.FileSystem.FileExists(Application.StartupPath & "\cache\cache.id")) Then noId = True
        If cCache Or fRun Or noId Then
            If cCache Or noId Then Clear_Cache()
            '
            If ExtractBack Then timeout = 1
            Current_Path = Check_File("proto\items\items.lst")
            Dim pLST() As String = IO.File.ReadAllLines(Current_Path & "\proto\items\items.lst")
            For n = UBound(pLST) To 0 Step -1
                pLST(n) = pLST(n).Trim
                If pLST(n).Length > 0 Then
                    pLST(n) = "proto\items\" & pLST(n)
                    noRe = True
                ElseIf noRe = False Then
                    ReDim Preserve pLST(n - 1)
                End If
            Next
            noRe = False
            IO.File.WriteAllLines("iProto.lst", pLST)
            SplashScreen.Label1.Text = "Loading: Extracting items Pro-files..."
            SplashScreen.ProgressBar1.Value = 20
            Application.DoEvents()
            Shell(Application.StartupPath & "\dat2.exe x -d cache """ & Game_Path & "\master.dat"" " & "@iProto.lst", AppWinStyle.Hide, True, 10000)
            '
            Current_Path = Check_File("proto\critters\critters.lst")
            pLST = IO.File.ReadAllLines(Current_Path & "\proto\critters\critters.lst")
            For n = 0 To UBound(pLST)
                pLST(n) = pLST(n).Trim
                If pLST(n).Length > 0 Then
                    pLST(n) = "proto\critters\" & pLST(n)
                    noRe = True
                ElseIf noRe = False Then
                    ReDim Preserve pLST(n - 1)
                End If
            Next
            IO.File.WriteAllLines("cProto.lst", pLST)
            SplashScreen.Label1.Text = "Loading: Extracting critter Pro-files..."
            SplashScreen.ProgressBar1.Value = 40
            Application.DoEvents()
            Shell(Application.StartupPath & "\dat2.exe x -d cache """ & Game_Path & "\master.dat"" " & "@cProto.lst", AppWinStyle.Hide, True, timeout)
            IO.File.Create(Application.StartupPath & "\cache\cache.id")
        End If

        SplashScreen.ProgressBar1.Value = 60
        CreateItemsList()
        If Not (cCache) And Not (ExtractBack) Then
            SplashScreen.ProgressBar1.Value = 80
            CreateCritterList()
        End If
        GetScriptLst()

        AI = IO.File.ReadAllLines(Application.StartupPath & "\ai.def")
        Teams = IO.File.ReadAllLines(Application.StartupPath & "\teams.def")

        SplashScreen.ProgressBar1.Value = 100
        Main_Form.Show()
        If fRun Then AboutBox.ShowDialog() : fRun = False
        Exit Sub
ExitProg:
        Application.Exit()
    End Sub

    Friend Sub GetScriptLst()
        Dim splt() As String
        Current_Path = Check_File("scripts\scripts.lst")
        Scripts_Lst = IO.File.ReadAllLines(Current_Path & "\scripts\scripts.lst")
        For n As Integer = 0 To UBound(Scripts_Lst)
            splt = Scripts_Lst(n).Split("#")
            splt = splt(0).Split(";")
            Scripts_Lst(n) = splt(0) & " [" & (n + 1) & "]      -  " & splt(1).Trim
        Next
    End Sub

    Friend Sub CreateCritterList()
        Dim n As Integer
        SetParent(Progress_Form.Handle.ToInt32, Main_Form.Handle.ToInt32)
        Progress_Form.SetDesktopLocation(Main_Form.Width / 4, Main_Form.Height / 2.25)
        Progress_Form.Show()
        Application.DoEvents()

        Current_Path = Check_File("proto\critters\critters.lst")
        Critter_LST = IO.File.ReadAllLines(Current_Path & "\proto\critters\critters.lst")
        ClearTrimLine(Critter_LST)
        ClearEmptyLine(Critter_LST)
        n = UBound(Critter_LST)
        Progress_Form.ProgressBar1.Maximum = n

        GetMsgData("pro_crit.msg")
        ReDim Critter_NAME(n)
        For n = 0 To UBound(Critter_LST)
            Critter_NAME(n) = GetNameCritter(GetProCritNameID(Critter_LST(n)))
            Main_Form.ListView1.Items.Add(Critter_NAME(n))
            Main_Form.ListView1.Items(n).SubItems.Add(Critter_LST(n))
            If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & Critter_LST(n)) = True Then
                Main_Form.ListView1.Items(n).Text = "* " & Critter_NAME(n)
                If (IO.File.GetAttributes(SaveMOD_Path & "\proto\critters\" & Critter_LST(n)) And &H1) = FileAttributes.ReadOnly Then
                    Main_Form.ListView1.Items(n).SubItems.Add("R/O")
                Else
                    Main_Form.ListView1.Items(n).SubItems.Add(vbNullString)
                End If
            Else
                Main_Form.ListView1.Items(n).SubItems.Add(vbNullString)
            End If
            Main_Form.ListView1.Items(n).SubItems.Add(&H1000001 + n)
            Progress_Form.ProgressBar1.Value = n
            Application.DoEvents()
        Next

        Current_Path = Check_File("art\critters\critters.lst", True)
        Critters_FRM = IO.File.ReadAllLines(Current_Path & "\art\critters\critters.lst")
        ClearTrimLine(Critters_FRM)
        ClearEmptyLine(Critters_FRM)

        Main_Form.ListView1.Visible = True
        Progress_Form.Close()
    End Sub

    Friend Sub CreateItemsList()
        Dim n As Integer, x As Integer

        'Progress_Form.MdiParent = Main_Form Main_Form.SplitContainer1.Panel1.Handle.ToInt32
        SetParent(Progress_Form.Handle.ToInt32, Main_Form.Handle.ToInt32)
        Progress_Form.SetDesktopLocation(Main_Form.Width / 4, Main_Form.Height / 2.25)
        Progress_Form.Show()
        Application.DoEvents()

        Current_Path = Check_File("art\critters\critters.lst", True)
        On Error GoTo CrtLstBadFile
        Critters_FRM = IO.File.ReadAllLines(Current_Path & "\art\critters\critters.lst")
        ClearTrimLine(Critters_FRM)
        ClearEmptyLine(Critters_FRM)
        On Error GoTo 0

        Current_Path = Check_File("art\items\items.lst")
        'On Error GoTo ItmsLstBadFile
        Items_FRM = IO.File.ReadAllLines(Current_Path & "\art\items\items.lst")
        ClearTrimLine(Items_FRM)
        ClearEmptyLine(Items_FRM)
        'On Error GoTo 0

        Current_Path = Check_File("art\inven\inven.lst")
        'On Error GoTo InvnLstBadFile
        Iven_FRM = IO.File.ReadAllLines(Current_Path & "\art\inven\inven.lst")
        ClearTrimLine(Iven_FRM)
        ClearEmptyLine(Iven_FRM)
        'On Error GoTo 0

        GetMsgData("perk.msg")
        For n = 0 To UBound(MSG_DATATEXT)
            If GetParamMsg(MSG_DATATEXT(n)) = "1101" Then Exit For
            If MSG_DATATEXT(n).StartsWith("{") Then
                ReDim Preserve Perk_NAME(x)
                Perk_NAME(x) = GetParamMsg(MSG_DATATEXT(n), True)
                x += 1
            End If
        Next
        '
        GetMsgData("pro_item.msg")
        Current_Path = Check_File("proto\items\items.lst")
        Dim crtfile() As String = IO.File.ReadAllLines(Current_Path & "\proto\items\items.lst")
        ClearTrimLine(crtfile)
        ClearEmptyLine(crtfile)
        ReDim Items_LST(UBound(crtfile), 1)
        For n = 0 To UBound(crtfile)
            Items_LST(n, 0) = crtfile(n)
        Next
        '
        x = 0
        Progress_Form.ProgressBar1.Maximum = n
        ReDim Items_NAME(n - 1)
        For n = 0 To UBound(Items_LST)
            Items_NAME(n) = GetNameCritter(GetProItemsNameID(Items_LST(n, 0), n))
            If Items_NAME(n) = "" Then Items_NAME(n) = "<NoName>"
            Main_Form.ListView2.Items.Add(Items_NAME(n))
            Main_Form.ListView2.Items(n).SubItems.Add(Items_LST(n, 0))
            Main_Form.ListView2.Items(n).SubItems.Add(Items_LST(n, 1))
            CheckProFileRO(n)
            Main_Form.ListView2.Items(n).Tag = n 'запись индекса(pid) итема в item.lst
            If Items_LST(n, 1) = "Ammo" Then
                ReDim Preserve AmmoPID(x)
                ReDim Preserve AmmoNAME(x)
                AmmoPID(x) = n + 1
                AmmoNAME(x) = Items_NAME(n)
                x += 1
            End If
            Progress_Form.ProgressBar1.Value = n
            Application.DoEvents()
        Next
        '
        Current_Path = Check_File("proto\misc\misc.lst")
        Misc_LST = IO.File.ReadAllLines(Current_Path & "\proto\misc\misc.lst")
        ClearTrimLine(Misc_LST)
        ClearEmptyLine(Misc_LST)
        '
        GetMsgData("pro_misc.msg")
        ReDim Misc_NAME(UBound(Misc_LST))
        For n = 0 To UBound(Misc_LST)
            Misc_NAME(n) = GetNameCritter((n + 1) * 100)
        Next
        '
        GetMsgData("proto.msg")
        x = 0
        For n = GetMSGLine(300) To UBound(MSG_DATATEXT)
            If GetParamMsg(MSG_DATATEXT(n)) = "350" Then Exit For
            If MSG_DATATEXT(n).StartsWith("{") Then
                ReDim Preserve CaliberNAME(x)
                CaliberNAME(x) = GetParamMsg(MSG_DATATEXT(n), True)
                x += 1
            End If
        Next
        '
        Main_Form.ListView2.Visible = True
        Progress_Form.Close()
        Exit Sub

CrtLstBadFile:
        MsgBox("Cannot open the required file: art\critter\critter.lst", MsgBoxStyle.Critical, "File Missing")
        SplashScreen.TopMost = False
        Setting_Form.ShowDialog()
        Application.Exit()
    End Sub

    Friend Sub FilterCreateItemsList()
        Dim n As Integer = 0
        Dim filter As String = Nothing
        Dim x As UShort = 0

        If Main_Form.fWeaponToolStripMenuItem3.Checked = True Then filter = "Weapon"
        If Main_Form.fAmmoToolStripMenuItem2.Checked = True Then filter = "Ammo"
        If Main_Form.fArmorToolStripMenuItem2.Checked = True Then filter = "Armor"
        If Main_Form.fDrugToolStripMenuItem3.Checked = True Then filter = "Drugs"
        If Main_Form.fMiscToolStripMenuItem2.Checked = True Then filter = "Misc"
        If Main_Form.fContainerToolStripMenuItem.Checked = True Then filter = "Container"

        For n = 0 To UBound(Items_LST)
            'Items_NAME(n) = GetNameCritter(GetProItemsNameID(Items_LST(n, 0), n))
            If filter <> Nothing Then
                If Items_LST(n, 1) = filter Or (filter = "Misc" And Items_LST(n, 1) = "Key") Then
                    Main_Form.ListView2.Items.Add(Items_NAME(n))
                    Main_Form.ListView2.Items(x).SubItems.Add(Items_LST(n, 0))
                    Main_Form.ListView2.Items(x).SubItems.Add(Items_LST(n, 1))
                    If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & Items_LST(n, 0)) = True Then
                        Main_Form.ListView2.Items(x).Text = "* " & Items_NAME(n)
                        If (IO.File.GetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(n, 0)) And &H1) = FileAttributes.ReadOnly Then
                            Main_Form.ListView2.Items(x).SubItems.Add("R/O")
                        Else
                            Main_Form.ListView2.Items(x).SubItems.Add(vbNullString)
                        End If
                    Else
                        Main_Form.ListView2.Items(x).SubItems.Add(vbNullString)
                    End If
                    Main_Form.ListView2.Items(x).Tag = n 'запись индекса(pid) итема в item.lst
                    x += 1
                End If
            Else
                Main_Form.ListView2.Items.Add(Items_NAME(n))
                Main_Form.ListView2.Items(n).SubItems.Add(Items_LST(n, 0))
                Main_Form.ListView2.Items(n).SubItems.Add(Items_LST(n, 1))
                CheckProFileRO(n)
                Main_Form.ListView2.Items(n).Tag = n 'запись индекса(pid) итема в item.lst
            End If
        Next
        Main_Form.ListView2.Visible = True
    End Sub

    Friend Sub ClearTrimLine(ByRef Massive() As String)
        Dim n As Integer
        For n = 0 To UBound(Massive)
            Massive(n) = Massive(n).Trim
        Next
    End Sub

    Friend Sub ClearEmptyLine(ByRef Massive() As String)
        For n = UBound(Massive) To 0 Step -1
            If Massive(n).Length > 0 Then Exit Sub
            ReDim Preserve Massive(n - 1)
        Next
    End Sub

    'Поиск индекса предмета в списке ListView
    Friend Function LW_SearhItemIndex(ByRef indx As UShort, ByRef LW As ListView)
        For Each Item As ListViewItem In LW.Items
            If Item.Tag = indx Then
                Return Item.Index
            End If
        Next
        Return Nothing
    End Function

    ' Проверяет профайл итема на атрибут чтения и ставит соответствующие метки в листе 
    Private Sub CheckProFileRO(ByRef n As Integer)
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & Items_LST(n, 0)) = True Then
            Main_Form.ListView2.Items(n).Text = "* " & Items_NAME(n)
            If (IO.File.GetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(n, 0)) And &H1) = FileAttributes.ReadOnly Then
                Main_Form.ListView2.Items(n).SubItems.Add("R/O")
            Else
                Main_Form.ListView2.Items(n).SubItems.Add(vbNullString)
            End If
        Else
            Main_Form.ListView2.Items(n).SubItems.Add(vbNullString)
        End If
    End Sub

    ' Получает Description ID из про-файла предмета и определяет его тип
    Friend Function GetProItemsNameID(ByRef ProFile As String, ByRef n As Integer) As Integer
        Dim NameID, TypeID As Integer
        Current_Path = Check_File("proto\items\" & ProFile)
        Dim rFile As New BinaryReader(File.Open(Current_Path & "\proto\items\" & ProFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        rFile.BaseStream.Seek(4, SeekOrigin.Begin)
        NameID = rFile.ReadInt32()
        rFile.BaseStream.Seek(&H20, SeekOrigin.Begin)
        TypeID = rFile.ReadInt32()
        rFile.Close()
        'определяем тип предмета
        'Dim fInfo As IO.FileInfo = New IO.FileInfo(Current_Path & "\proto\items\" & ProFile)
        Select Case ReverseBytes(TypeID) 'fInfo.Length
            Case 5 '69 'Misc
                Items_LST(n, 1) = "Misc"
            Case 1 '65 'Контейнер
                Items_LST(n, 1) = "Container"
            Case 3 '122 'Оружие 
                Items_LST(n, 1) = "Weapon"
            Case 2 '125 'Наркотик
                Items_LST(n, 1) = "Drugs"
            Case 4 '81 'Патрон
                Items_LST(n, 1) = "Ammo"
            Case 0 '129 'Броня
                Items_LST(n, 1) = "Armor"
            Case 6 '61 'Key
                Items_LST(n, 1) = "Key"
            Case Else
                Items_LST(n, 1) = "Unknown"
        End Select
        Return ReverseBytes(NameID)
    End Function

    Friend Sub Get_Config()
        Dim file As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\config.ini")
        Dim strIni As String = Nothing
        On Error GoTo SetDefConf
        Do Until file.EndOfStream
            strIni = file.ReadLine
            Select Case strIni
                Case "[Path]"
                    Game_Path = file.ReadLine.Substring(11)
                    SaveMOD_Path = file.ReadLine.Substring(8)
                    Dim HEXPath As String = file.ReadLine.Substring(8)
                    If HEXPath <> Nothing Then HEX_Path = HEXPath : Setting_Form.TextBox3.Enabled = True
                Case "[Option]"
                    proRO = file.ReadLine.Substring(9)
                    txtWin = file.ReadLine.Substring(7)
                    txtLvCp = file.ReadLine.Substring(6)
                    cCache = file.ReadLine.Substring(11)
                    cArtCache = file.ReadLine.Substring(14)
                    ExtractBack = file.ReadLine.Substring(11) 'Background=
                    SplitSize = CInt(file.ReadLine.Substring(10))
            End Select
        Loop
        file.Close()
        If Game_Path = Nothing Then GoTo SetDefConf
        If SaveMOD_Path = Nothing Then SaveMOD_Path = Game_Path & "\Data" : Setting_Form.TextBox2.ReadOnly = True
        GameDATA_Path = Game_Path & "\Data" '"\MyTestData" '
        Main_Form.LinkLabel1.Text = GameDATA_Path
        Main_Form.LinkLabel2.Text = SaveMOD_Path
        Exit Sub
SetDefConf:
        'On Error Resume Next
        file.Close()
        proRO = True : txtWin = True : cCache = True : ExtractBack = True : fRun = True
        SettingExt = True
        SplashScreen.TopMost = False
        Setting_Form.ShowDialog()
        SplashScreen.TopMost = True
    End Sub

    'Проверить наличее файла если его нет то извлеч из Dat
    Friend Function Check_File(ByRef cFile As String, Optional ByRef сDat As Boolean = False, Optional ByRef unpack As Boolean = True) As String
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\" & cFile) = True Then
            Return SaveMOD_Path
        ElseIf My.Computer.FileSystem.FileExists(GameDATA_Path & "\" & cFile) = True Then
            Return GameDATA_Path
        ElseIf My.Computer.FileSystem.FileExists(Game_Path & MasterDAT & "\" & cFile) = True OrElse My.Computer.FileSystem.FileExists(Game_Path & CritterDAT & "\" & cFile) = True Then
            If сDat Then
                Return Game_Path & CritterDAT 'папка
            Else
                Return Game_Path & MasterDAT  'папка
            End If
        Else
            If My.Computer.FileSystem.FileExists(Cache_Patch & "\" & cFile) = False And unpack Then
                'Извлечь требуемый файл в кэш
                'Main_Form.ToolStripStatusLabel1.Text = "Извлечение: " & cFile
                Main_Form.TextBox1.Text = "Extraction: " & cFile & vbCrLf & Main_Form.TextBox1.Text
                If сDat Then
                    Shell(Application.StartupPath & "\dat2.exe x -d cache """ & Game_Path & "\critter.dat"" " & cFile, AppWinStyle.Hide, True)
                Else
                    Shell(Application.StartupPath & "\dat2.exe x -d cache """ & Game_Path & "\master.dat"" " & cFile, AppWinStyle.Hide, True)
                End If
            End If
            Return Cache_Patch
        End If
    End Function

    Friend Function UnDat_File(ByRef pFile As String) As Boolean
        If FileLen(Cache_Patch & "\" & pFile) = 416 Then Return True
        'Извлечь требуемый файл в кэш
        Shell(Application.StartupPath & "\dat2.exe x -d cache """ & Game_Path & "\master.dat"" " & pFile, AppWinStyle.Hide, True)
        If My.Computer.FileSystem.FileExists(Cache_Patch & "\" & pFile) Then Return True
        Return False
    End Function

    'Cоздает и открывает новую форму для редактирования криттера
    Friend Sub Create_CritterForm(ByVal Lw_Index As UShort)
        Dim CrttFrm As New Critter_Form
        'CrttFrm.MdiParent = Main_Form
        SetParent(CrttFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        CrttFrm.Ini_CritterForm(Lw_Index)
        On Error GoTo EndProc
        CrttFrm.ComboBox1.Items.AddRange(Critters_FRM)
        CrttFrm.ComboBox2.Items.AddRange(AI)
        CrttFrm.ComboBox3.Items.AddRange(Teams)
        CrttFrm.ComboBox9.Items.AddRange(Scripts_Lst)
        CrttFrm.Show()
EndProc:
    End Sub

    'Cоздает и открывает новую форму для редактирования предметов
    Friend Sub Create_ItemsForm(ByVal LST_Index As UShort)
        Dim n As UShort
        Dim ItmsFrm As New Items_Form
        'ItmsFrm.MdiParent = Main_Form
        SetParent(ItmsFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        ItmsFrm.ComboBox10.Items.AddRange(Misc_NAME)
        ItmsFrm.ComboBox11.Items.AddRange(Perk_NAME)
        ItmsFrm.ComboBox12.Items.AddRange(CaliberNAME)
        For n = 0 To UBound(AmmoPID)
            ItmsFrm.ComboBox13.Items.Add(AmmoNAME(n))
        Next
        ItmsFrm.ComboBox16.Items.AddRange(Critters_FRM)
        ItmsFrm.ComboBox17.Items.AddRange(Critters_FRM)
        ItmsFrm.ComboBox18.Items.AddRange(Perk_NAME)
        ItmsFrm.ComboBox22.Items.AddRange(Perk_NAME)
        ItmsFrm.ComboBox23.Items.AddRange(CaliberNAME)
        For n = 0 To UBound(AmmoPID)
            ItmsFrm.ComboBox24.Items.Add(AmmoNAME(n))
        Next
        ItmsFrm.ComboBox25.Items.AddRange(CaliberNAME)
        'тип предмета
        Select Case Items_LST(LST_Index, 1)
            Case "Weapon"
                ItmsFrm.TabControl1.TabPages.RemoveAt(3)
                ItmsFrm.TabControl1.TabPages.RemoveAt(2)
                ItmsFrm.TabControl1.TabPages.RemoveAt(1)
            Case "Armor"
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(3))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(2))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(0))
            Case "Drugs"
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(3))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(1))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(0))
            Case "Key", "Container"
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(3))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(2))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(1))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(0))
            Case Else
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(2))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(1))
                ItmsFrm.TabControl1.TabPages.Remove(ItmsFrm.TabControl1.TabPages.Item(0))
        End Select
        ItmsFrm.ComboBox1.Items.AddRange(Items_FRM)
        ItmsFrm.ComboBox2.Items.AddRange(Iven_FRM)
        ItmsFrm.ComboBox9.Items.AddRange(Scripts_Lst)
        '
        ItmsFrm.Ini_ItemsForm(LST_Index)
        ItmsFrm.Show()
    End Sub

    Friend Sub Create_TxtEditForm(ByRef Lw_Index As UShort, ByRef Type As Byte)
        Dim TxtFrm As New TxtEdit_Form
        SetParent(TxtFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        If Type = 0 Then TxtFrm.Text &= Critter_LST(Lw_Index) Else TxtFrm.Text &= Items_LST(Lw_Index, 0)
        TxtFrm.Text &= "]"
        TxtFrm.Ini_Form(Lw_Index, Type)
    End Sub

    Friend Sub Create_AIEditForm(Optional ByRef AIPacket As Integer = -1)
        Dim AIFrm As New AI_Form
        SetParent(AIFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        AIFrm.Initialize(AIPacket)
    End Sub

    ' Получает Description ID из про-файла криттера
    Friend Function GetProCritNameID(ByRef ProFile As String) As Integer
        Dim NameID As Integer
        Dim fFile As Byte = FreeFile()
        Current_Path = Check_File("proto\critters\" & ProFile)
        FileOpen(fFile, Current_Path & "\proto\critters\" & ProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        FileGet(fFile, NameID, 5)
        FileClose(fFile)
        Return ReverseBytes(NameID)
    End Function

    'Инвертирует значение в Big-endian и обратно
    Friend Function ReverseBytes(ByRef Value As Integer) As Integer
        Dim bt1() As Byte = BitConverter.GetBytes(Value)
        Array.Reverse(bt1)
        Return BitConverter.ToInt32(bt1, 0)
        'Return (Value And &HFF) << 24 Or (Value And &HFF00) << 8 Or (Value And &HFF0000) >> 8 Or (Value And &HFF000000) >> 24
    End Function

    'из структуры в байты
    Friend Function fnStructToBytes(ByVal MyStruct As CritPro) As Integer()
        Dim size As Integer = Marshal.SizeOf(MyStruct)
        Dim Buff(size - 1) As Integer
        Dim Ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(MyStruct))
        Marshal.StructureToPtr(MyStruct, Ptr, False)
        Marshal.Copy(Ptr, Buff, 0, size)
        Marshal.FreeHGlobal(Ptr)
        Return Buff
    End Function

    'Из байтов в структуру
    Friend Function fnBytesToStruct(ByVal Buff() As Integer, ByVal MyType As System.Type) As Object
        Dim MyGC As GCHandle = GCHandle.Alloc(Buff, GCHandleType.Pinned)
        Dim Obj As Object = Marshal.PtrToStructure(MyGC.AddrOfPinnedObject, MyType)
        MyGC.Free()
        Return Obj
    End Function

    'Считывает содержимое из msg файла в массив с соответствующей выбранной кодировкой
    Friend Sub GetMsgData(ByRef msgFile As String)
        Current_Path = Check_File("Text\English\Game\" & msgFile) '"pro_item.msg"
        If txtWin Then MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\" & msgFile, System.Text.Encoding.Default) _
        Else MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\" & msgFile, System.Text.Encoding.GetEncoding("cp866"))
        If txtLvCp Then EncodingLevCorp()
    End Sub

    'Возвращает Имя криттера или его описание
    Friend Function GetNameCritter(ByRef NameID As Integer) As String
        Dim a As Integer
        Dim strLine As String = Nothing
        Dim sNameID As String = CStr(NameID)
        'Ищем строку с номером NameID
        For a = 0 To UBound(MSG_DATATEXT)
            strLine = MSG_DATATEXT(a)
            If GetParamMsg(strLine) = sNameID Then Return GetParamMsg(strLine, True)
        Next
        Return "" 'Nothing
    End Function

    'Возвращает параметры из строки формата Msg  
    Friend Function GetParamMsg(ByRef str As String, Optional ByRef strValue As Boolean = False) As String
        If str.Length < 2 Then Return Nothing 'On Error GoTo erret
        Dim n As Integer = str.IndexOf("}", 1) 'InStr(2, str, "}", CompareMethod.Text)
        If n = -1 Then Return Nothing
        'Извлекаем
        If strValue = True Then
            n = str.IndexOf("{", n + 2) 'n = InStr(n + 2, str, "{", CompareMethod.Text)
            Return str.Substring(n + 1, str.Length - (n + 2))
        Else
            Return str.Substring(1, n - 1) 'Mid(str, 2, n - 2)
        End If
erret:
        Return Nothing
    End Function

    'Возвращает номер строки массива msg-файла
    Friend Function GetMSGLine(ByRef NameID As Integer) As Integer
        Dim a As Integer
        Dim strLine As String = Nothing
        Dim sNameID As String = CStr(NameID)
        'Ищем строку с номером NameID
        For a = 0 To UBound(MSG_DATATEXT)
            strLine = MSG_DATATEXT(a)
            If GetParamMsg(strLine) = sNameID Then Return a
        Next
        Return -1
    End Function

    'Добавление или измнение значения в MSG файле
    Friend Function Add_ProMSG(ByVal str As String, ByVal ID As Integer, ByVal desc As Boolean) As Boolean
        Dim a As Integer = GetMSGLine(ID)
        If a = -1 And desc Then Return True
        If a = -1 Then
            ReDim Preserve MSG_DATATEXT(UBound(MSG_DATATEXT) + 1)
            a = UBound(MSG_DATATEXT)
        End If
        If desc = True Then
            ID += 1 : a += 1
            Dim b As Integer = GetMSGLine(ID)
            If b = -1 Then
                Dim n As Integer
                ReDim Preserve MSG_DATATEXT(UBound(MSG_DATATEXT) + 1)
                For n = UBound(MSG_DATATEXT) To a Step -1
                    MSG_DATATEXT(n) = MSG_DATATEXT(n - 1)
                Next
            End If
        End If
        MSG_DATATEXT(a) = "{" & ID & "}{}{" & str & "}"
        Return False
    End Function

    Friend Sub EncodingLevCorp()
        Dim data() As Byte
        For n As Integer = 0 To UBound(MSG_DATATEXT)
            If MSG_DATATEXT(n).StartsWith("{") Then
                data = System.Text.Encoding.GetEncoding("cp866").GetBytes(MSG_DATATEXT(n))
                For m As Integer = 0 To UBound(data)
                    If data(m) >= &HB0 And data(m) <= &HBF Then
                        data(m) = data(m) + &H30
                    End If
                Next
                MSG_DATATEXT(n) = System.Text.Encoding.GetEncoding("cp866").GetString(data)
            End If
        Next
    End Sub

    Friend Sub CodingToLevCorp(ByRef str As String)
        Dim data() As Byte = System.Text.Encoding.GetEncoding("cp866").GetBytes(str)
        For m As Integer = 0 To UBound(data)
            If data(m) >= &HE0 And data(m) <= &HEF Then
                data(m) = data(m) - &H30
            End If
        Next
        str = System.Text.Encoding.GetEncoding("cp866").GetString(data)
    End Sub

    'for critters
    Friend Sub Frm_to_Gif(ByRef FrmFile As String)
        On Error Resume Next
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\art\critters\" & FrmFile & "aa.frm") = True Then
            My.Computer.FileSystem.CopyFile(SaveMOD_Path & "\art\critters\" & FrmFile & "aa.frm", Application.StartupPath & "\art\critters\" & FrmFile & "aa.frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\critters\" & FrmFile & "aa.frm", AppWinStyle.Hide, True)
        ElseIf My.Computer.FileSystem.FileExists(GameDATA_Path & "\art\critters\" & FrmFile & "aa.frm") = True Then
            My.Computer.FileSystem.CopyFile(GameDATA_Path & "\art\critters\" & FrmFile & "aa.frm", Application.StartupPath & "\art\critters\" & FrmFile & "aa.frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\critters\" & FrmFile & "aa.frm", AppWinStyle.Hide, True)
        ElseIf My.Computer.FileSystem.FileExists(Game_Path & CritterDAT & "\art\critters\" & FrmFile & "aa.frm") = True Then
            My.Computer.FileSystem.CopyFile(Game_Path & CritterDAT & "\art\critters\" & FrmFile & "aa.frm", Application.StartupPath & "\art\critters\" & FrmFile & "aa.frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\critters\" & FrmFile & "aa.frm", AppWinStyle.Hide, True)
        Else 'Извлечь требуемый файл
            Shell(Application.StartupPath & "\frm2gif.exe -d -f """ & Game_Path & "\critter.dat"" -p color.pal " & FrmFile & "aa.frm", AppWinStyle.Hide, True)
        End If
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\art\critters\" & FrmFile & "aa.gif") Then
            My.Computer.FileSystem.MoveFile(Application.StartupPath & "\art\critters\" & FrmFile & "aa.gif", Cache_Patch & "\art\critters\" & FrmFile & "aa.gif")
        Else
            My.Computer.FileSystem.MoveFile(Application.StartupPath & "\art\critters\" & FrmFile & "aa_sw.gif", Cache_Patch & "\art\critters\" & FrmFile & "aa.gif")
        End If
        My.Computer.FileSystem.DeleteDirectory(Application.StartupPath & "\art", FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

    Friend Sub ItemFrmGif(ByVal Path As String, ByVal FrmFile As String)
        On Error Resume Next
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\art\" & Path & FrmFile & ".frm") = True Then
            My.Computer.FileSystem.CopyFile(SaveMOD_Path & "\art\" & Path & FrmFile & ".frm", Application.StartupPath & "\art\" & Path & FrmFile & ".frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\" & Path & FrmFile & ".frm", AppWinStyle.Hide, True)
        ElseIf My.Computer.FileSystem.FileExists(GameDATA_Path & "\art\" & Path & FrmFile & ".frm") = True Then
            My.Computer.FileSystem.CopyFile(GameDATA_Path & "\art\" & Path & FrmFile & ".frm", Application.StartupPath & "\art\" & Path & FrmFile & ".frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\" & Path & FrmFile & ".frm", AppWinStyle.Hide, True)
        ElseIf My.Computer.FileSystem.FileExists(Game_Path & MasterDAT & "\art\" & Path & FrmFile & ".frm") = True Then
            My.Computer.FileSystem.CopyFile(Game_Path & MasterDAT & "\art\" & Path & FrmFile & ".frm", Application.StartupPath & "\art\" & Path & FrmFile & ".frm")
            Shell(Application.StartupPath & "\frm2gif.exe -p color.pal " & ".\art\" & Path & FrmFile & ".frm", AppWinStyle.Hide, True)
        Else 'Извлечь требуемый файл
            Shell(Application.StartupPath & "\frm2gif.exe -d -f """ & Game_Path & "\master.dat"" -p color.pal " & FrmFile & ".frm", AppWinStyle.Hide, True)
        End If
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\art\" & Path & FrmFile & ".gif") Then
            My.Computer.FileSystem.MoveFile(Application.StartupPath & "\art\" & Path & FrmFile & ".gif", Cache_Patch & "\art\" & Path & FrmFile & ".gif")
        Else
            My.Computer.FileSystem.MoveFile(Application.StartupPath & "\art\" & Path & FrmFile & "_ne.gif", Cache_Patch & "\art\" & Path & FrmFile & ".gif")
        End If
        My.Computer.FileSystem.DeleteDirectory(Application.StartupPath & "\art", FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

    Friend Sub Clear_Cache()
        On Error Resume Next
        My.Computer.FileSystem.DeleteDirectory(Cache_Patch & "\proto", FileIO.DeleteDirectoryOption.DeleteAllContents)
        My.Computer.FileSystem.DeleteDirectory(Cache_Patch & "\scripts", FileIO.DeleteDirectoryOption.DeleteAllContents)
        My.Computer.FileSystem.DeleteDirectory(Cache_Patch & "\text", FileIO.DeleteDirectoryOption.DeleteAllContents)
        My.Computer.FileSystem.DeleteFile(Cache_Patch & "\cache.id")
        Clear_Art_Cache()
    End Sub

    Friend Sub Clear_Art_Cache()
        On Error Resume Next
        My.Computer.FileSystem.DeleteDirectory(Cache_Patch & "\art", FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub

End Module
