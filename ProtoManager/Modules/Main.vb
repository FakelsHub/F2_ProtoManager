Option Explicit On
Imports System.IO
Imports Prototypes
Imports System.Drawing

Friend Module Main

    Structure CrittersLst
        Friend proFile As String
        Friend crtName As String
        Friend crtHP As Integer
    End Structure

    Structure ItemsLst
        Friend proFile As String
        Friend itemType As Integer
        Friend itemName As String
    End Structure

    Friend Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer

    Friend Critter_LST As CrittersLst()
    Friend Items_LST As ItemsLst()
    '
    Friend Critters_FRM As String()
    Friend Items_FRM As String()
    Friend Iven_FRM As String()
    '
    Private Misc_LST As String()
    Friend Misc_NAME As String()
    '
    Friend Scripts_Lst As String()
    '
    Friend AmmoPID As Integer()
    Friend AmmoNAME As String()
    Friend CaliberNAME As String()
    Friend Perk_NAME As String()
    '
    Private Teams As List(Of String) = New List(Of String)
    Friend PacketAI As SortedList(Of String, Integer)

    'Initialization...
    Friend Sub Main()
        Dim noRe, noId, wait As Boolean

        SplashScreen.ProgressBar1.Value = 10
        Application.DoEvents()

        If Not (File.Exists(Cache_Patch & "\cache.id")) Then noId = True
        If cCache OrElse Setting_Form.fRun OrElse noId Then
            If cCache OrElse noId Then Settings.Clear_Cache()
            '
            If Not (ExtractBack) Then wait = True
            Dim pLST() As String = File.ReadAllLines(DatFiles.CheckFile(itemsLstPath))
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
            File.WriteAllLines("iProto.lst", pLST)

            SplashScreen.Label1.Text = "Loading: Extracting items Pro-files..."
            SplashScreen.ProgressBar1.Value = 20
            Application.DoEvents()

            Shell(WorkAppDIR & "\dat2.exe x -d cache """ & Game_Path & MasterDAT & """ " & "@iProto.lst", AppWinStyle.Hide, True, 60000)

            pLST = File.ReadAllLines(DatFiles.CheckFile(crittersLstPath))
            For n = 0 To UBound(pLST)
                pLST(n) = pLST(n).Trim
                If pLST(n).Length > 0 Then
                    pLST(n) = "proto\critters\" & pLST(n)
                    noRe = True
                ElseIf noRe = False Then
                    ReDim Preserve pLST(n - 1)
                End If
            Next
            File.WriteAllLines("cProto.lst", pLST)

            SplashScreen.Label1.Text = "Loading: Extracting critter Pro-files..."
            SplashScreen.ProgressBar1.Value = 40
            Application.DoEvents()

            Shell(WorkAppDIR & "\dat2.exe x -d cache """ & Game_Path & MasterDAT & """ " & "@cProto.lst", AppWinStyle.Hide, wait, 60000)
            File.Create(Cache_Patch & "\cache.id").Close()
        End If

        SplashScreen.ProgressBar1.Value = 60

        Settings.SetEncoding()

        GetCrittersLstFRM()
        CreateItemsList()
        If Not (cCache) And Not (ExtractBack) Then
            SplashScreen.ProgressBar1.Value = 80
            CreateCritterList()
        End If

        SplashScreen.ProgressBar1.Value = 100
        Main_Form.Show()

        If Setting_Form.fRun Then
            Setting_Form.fRun = False
            AboutBox.ShowDialog()
        End If
    End Sub

    Friend Sub GetScriptLst()
        If Scripts_Lst IsNot Nothing Then Return

        Dim splt() As String

        Scripts_Lst = File.ReadAllLines(DatFiles.CheckFile(scriptsLstPath), System.Text.Encoding.Default)
        For n As Integer = 0 To UBound(Scripts_Lst)
            splt = Scripts_Lst(n).Split("#"c)
            splt = splt(0).Split(";"c)
            Scripts_Lst(n) = String.Format("{0} {1} - {2}", splt(0).TrimEnd.PadRight(14),
                                           ("[" & (n + 1) & "]").PadRight(6), splt(1).Trim)
        Next
    End Sub

    Private Function GetCrittersLstFRM() As Boolean

        Try
            Critters_FRM = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(artCrittersLstPath, , True)))
        Catch ex As DirectoryNotFoundException
            MsgBox("Cannot open required file: \art\critter\critter.lst", MsgBoxStyle.Critical, "File Missing")
            Return True
        End Try

        For i As Integer = 0 To Critters_FRM.Length - 1
            Dim frm As String = Critters_FRM(i)
            Dim z As Integer = frm.IndexOf(","c)
            If z > 0 Then frm = frm.Remove(z)
            Critters_FRM(i) = frm.ToUpper
        Next

        Return False
    End Function

    Friend Sub GetItemsLstFRM()
        If Items_FRM Is Nothing Then
            Items_FRM = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(artItemsLstPath)))
        End If

        If Iven_FRM IsNot Nothing Then Return
        Iven_FRM = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(artInvenLstPath)))
    End Sub

    Friend Sub CreateCritterList()
        Progress_Form.ShowProgressBar(0)

        Dim lstfile() As String = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(crittersLstPath)))
        Dim cCount As Integer = UBound(lstfile)

        Progress_Form.ProgressBar1.Maximum = cCount
        With Main_Form
            .ListView1.BeginUpdate()
            .ListView1.Items.Clear()

            Messages.GetMsgData("pro_crit.msg")
            ReDim Critter_LST(cCount)
            For n = 0 To cCount
                Critter_LST(n).proFile = lstfile(n)
                Critter_LST(n).crtName = Messages.GetNameObject(ProFiles.GetProCritNameID(Critter_LST(n).proFile))
                If Critter_LST(n).crtName = String.Empty Then Critter_LST(n).crtName = "<NoName>"
                Dim proIsEdit As Boolean = False
                Dim rOnly As String = CheckProFileRO(proIsEdit, (PROTO_CRITTERS & Critter_LST(n).proFile))
                If (rOnly = String.Empty AndAlso proIsEdit) Then rOnly = "*"
                .ListView1.Items.Add(New ListViewItem({Critter_LST(n).crtName, Critter_LST(n).proFile, rOnly, (&H1000001 + n)}))
                .ListView1.Items(n).Tag = n 'запись индекса(pid) криттера в critters.lst
                If proIsEdit Then .ListView1.Items(n).ForeColor = Color.DarkBlue
                Progress_Form.ProgressBar1.Value = n
            Next
            .ListView1.EndUpdate()
        End With

        Progress_Form.Close()
    End Sub

    Friend Sub CreateItemsList()
        Dim tempList0 As List(Of String) = New List(Of String)()
        Dim tempList1 As List(Of Integer) = New List(Of Integer)()
        Dim n As Integer

        Progress_Form.ShowProgressBar(0)

        Messages.GetMsgData("pro_item.msg")
        Dim lstfile() As String = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(itemsLstPath)))
        ReDim Items_LST(UBound(lstfile))
        For n = 0 To UBound(lstfile)
            Items_LST(n).proFile = lstfile(n)
        Next

        Progress_Form.ProgressBar1.Maximum = n
        Application.DoEvents()

        With Main_Form
            .ListView2.BeginUpdate()
            .ListView2.Items.Clear()

            For n = 0 To UBound(Items_LST)
                Items_LST(n).itemName = Messages.GetNameObject(ProFiles.GetProItemsNameID(Items_LST(n).proFile, n))
                If Items_LST(n).itemName = String.Empty Then Items_LST(n).itemName = "<NoName>"
                Dim proIsEdit As Boolean = False
                Dim rOnly As String = CheckProFileRO(proIsEdit, (PROTO_ITEMS & Items_LST(n).proFile))
                If (rOnly = String.Empty AndAlso proIsEdit) Then rOnly = "*"
                .ListView2.Items.Add(New ListViewItem({Items_LST(n).itemName, Items_LST(n).proFile, ItemTypesName(Items_LST(n).itemType), rOnly}))
                .ListView2.Items(n).Tag = n 'запись индекса(pid) итема из item.lst
                If proIsEdit Then .ListView2.Items(n).ForeColor = Color.DarkBlue
                If Items_LST(n).itemType = ItemType.Ammo Then
                    tempList0.Add(Items_LST(n).itemName)
                    tempList1.Add(n + 1)
                End If
                Progress_Form.ProgressBar1.Value = n
            Next
            ReDim AmmoNAME(tempList0.Count - 1)
            ReDim AmmoPID(tempList1.Count - 1)
            tempList0.CopyTo(AmmoNAME)
            tempList1.CopyTo(AmmoPID)

            .ListView2.Visible = True
            .ListView2.EndUpdate()
        End With

        Progress_Form.Close()
    End Sub

    Private Sub GetItemsData()
        If Misc_NAME IsNot Nothing Then Return

        Dim tempList0 As List(Of String) = New List(Of String)()

        Misc_LST = ProFiles.ClearEmptyLines(File.ReadAllLines(DatFiles.CheckFile(miscLstPath)))
        Messages.GetMsgData("pro_misc.msg")
        ReDim Misc_NAME(UBound(Misc_LST))
        For n As Integer = 0 To UBound(Misc_LST)
            Misc_NAME(n) = Messages.GetNameObject((n + 1) * 100)
        Next

        Messages.GetMsgData("proto.msg")
        For n As Integer = Messages.GetMSGLine(300) To UBound(MSG_DATATEXT)
            If Messages.GetParamMsg(MSG_DATATEXT(n)) = "350" Then Exit For
            If MSG_DATATEXT(n).StartsWith("{") Then
                tempList0.Add(Messages.GetParamMsg(MSG_DATATEXT(n), True))
            End If
        Next
        ReDim CaliberNAME(tempList0.Count - 1)
        tempList0.CopyTo(CaliberNAME)
        tempList0.Clear()

        Messages.GetMsgData("perk.msg")
        For n As Integer = 0 To UBound(MSG_DATATEXT)
            If Messages.GetParamMsg(MSG_DATATEXT(n)) = "1101" Then Exit For
            If MSG_DATATEXT(n).StartsWith("{") Then
                tempList0.Add(Messages.GetParamMsg(MSG_DATATEXT(n), True))
            End If
        Next
        ReDim Perk_NAME(tempList0.Count - 1)
        tempList0.CopyTo(Perk_NAME)

    End Sub

    Friend Sub FilterCreateItemsList(ByVal filter As Integer)
        Dim x As Integer

        With Main_Form
            .ListView2.BeginUpdate()
            .ListView2.Items.Clear()
            For n As Integer = 0 To UBound(Items_LST)
                Dim proIsEdit As Boolean = False
                Dim rOnly As String = CheckProFileRO(proIsEdit, (PROTO_ITEMS & Items_LST(n).proFile))
                If (rOnly = String.Empty AndAlso proIsEdit) Then rOnly = "*"
                If filter <> ItemType.Unknown Then
                    If Items_LST(n).itemType = filter OrElse (filter = ItemType.Misc And Items_LST(n).itemType = ItemType.Key) Then
                        .ListView2.Items.Add(New ListViewItem({Items_LST(n).itemName, Items_LST(n).proFile, ItemTypesName(Items_LST(n).itemType), rOnly}))
                        .ListView2.Items(x).Tag = n 'указатель индекса(pid) итема в item.lst
                        If proIsEdit Then .ListView2.Items(x).ForeColor = Color.DarkBlue
                        x += 1
                    End If
                Else
                    .ListView2.Items.Add(New ListViewItem({Items_LST(n).itemName, Items_LST(n).proFile, ItemTypesName(Items_LST(n).itemType), rOnly}))
                    .ListView2.Items(n).Tag = n 'указатель индекса(pid) итема в item.lst
                    If proIsEdit Then .ListView2.Items(n).ForeColor = Color.Blue
                End If
            Next
            .ListView2.EndUpdate()
        End With
    End Sub

    ' Проверяет профайл итема на атрибут чтения и ставит соответствующие метки в листе 
    Private Function CheckProFileRO(ByRef exists As Boolean, ByVal pFile As String) As String
        Dim path As String = SaveMOD_Path & pFile

        If File.Exists(path) Then
            exists = True
            If (File.GetAttributes(path) And &H1) = FileAttributes.ReadOnly Then
                Return "R/O"
            End If
        End If

        Return String.Empty
    End Function

    'Поиск индекса предмета в списке ListView
    Friend Function LW_SearhItemIndex(ByVal indx As Integer, ByVal LW As ListView) As Integer
        For Each Item As ListViewItem In LW.Items
            If CInt(Item.Tag) = indx Then
                Return Item.Index
            End If
        Next

        Return Nothing
    End Function

    'Cоздает и открывает новую форму для редактирования криттера
    Friend Sub Create_CritterForm(ByVal cLST_Index As Integer)
        'Check...
        If GetCrittersLstFRM() Then Return
        GetScriptLst()
        GetTeams()
        If PacketAI Is Nothing Then PacketAI = AI.GetAllAIPacketNumber(DatFiles.CheckFile(AI.AIFILE))

        Dim CrttFrm As New Critter_Form(cLST_Index)
        With CrttFrm
            SetParent(.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32) 'CrttFrm.MdiParent = Main_Form
            .ComboBox1.Items.AddRange(Critters_FRM)
            .ComboBox2.Items.AddRange(PacketAI.Keys.ToArray)
            .ComboBox3.Items.AddRange(Teams.ToArray)
            .ComboBox9.Items.AddRange(Scripts_Lst)
            If .LoadProData() Then
                .Dispose()
            Else
                .Show()
            End If
        End With
    End Sub

    'Cоздает и открывает новую форму для редактирования предметов
    Friend Sub Create_ItemsForm(ByVal iLST_Index As Integer)
        If Items_LST(iLST_Index).itemType >= ItemType.Unknown Then
            MsgBox("This object has an unknown type." & vbLf & "The prototype has not the correct format or the file is corrupted.", MsgBoxStyle.Critical, "Error Item Type")
            Exit Sub
        End If

        If (Items_LST(iLST_Index).itemType = ItemType.Armor) AndAlso GetCrittersLstFRM() Then Return
        GetItemsData()
        GetItemsLstFRM()
        GetScriptLst()

        Dim ItmsFrm As New Items_Form(iLST_Index)
        SetParent(ItmsFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32) 'ItmsFrm.MdiParent = Main_Form
        ItmsFrm.IniItemsForm()
    End Sub

    Friend Sub Create_TxtEditForm(ByRef Lw_Index As Integer, ByRef Type As Byte)
        Dim TxtFrm As New TxtEdit_Form(Lw_Index, Type)

        SetParent(TxtFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        If Type = 0 Then
            TxtFrm.Text &= Critter_LST(Lw_Index).proFile
        Else
            TxtFrm.Text &= Items_LST(Lw_Index).proFile
        End If

        TxtFrm.Text &= "]"
        TxtFrm.Init_Data()
    End Sub

    Friend Sub Create_AIEditForm(Optional ByRef AIPacket As Integer = -1)
        Dim AIFrm As New AI_Form

        SetParent(AIFrm.Handle.ToInt32, Main_Form.SplitContainer1.Handle.ToInt32)
        AIFrm.Initialize(AIPacket)
    End Sub

    Private Sub GetTeams()
        If Teams.Count > 0 Then Return

        Dim tData As String() = File.ReadAllLines(WorkAppDIR & "\teams.h")
        For Each t In tData
            Dim line As String = t.Trim("/"c, " "c)
            If line.ToLower.StartsWith("#define ") Then
                Dim fSpace As Integer = line.IndexOf(" ", 9)
                If fSpace <= 0 Then Continue For
                Dim tName As String = line.Remove(fSpace).Remove(0, 8)
                fSpace = line.IndexOf("(", fSpace) + 1
                If fSpace <= 0 Then Continue For
                Dim tNum As String = line.Substring(fSpace, line.LastIndexOf(")") - fSpace).Trim
                Teams.Add(String.Format("{0} ({1})", Strings.RSet(tNum, 3), tName))
            End If
        Next
        Teams.Sort()
    End Sub

    Friend Sub PrintLog(ByVal textLog As String)
        Main_Form.TextBox1.AppendText(textLog & vbLf)
        Main_Form.TextBox1.ScrollToCaret()
    End Sub

End Module
