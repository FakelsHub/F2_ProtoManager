Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports System.Windows
Imports Prototypes

Friend Class Main_Form

    'Private ClickXClose As Boolean = False
    'Private Const SC_CLOSE As Int32 = &HF060
    '
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    ' Нажата кнопка Close.
    '    If m.WParam.ToInt32() = SC_CLOSE Then ClickXClose = True
    '    MyBase.WndProc(m)
    'End Sub

    Friend Sub SetFormSettings()
        If Not (Settings.HoverSelect) Then DontHoverSelectToolStripMenuItem.PerformClick()
        Cp886ToolStripMenuItem.Checked = Not (Settings.txtWin) And Not (Settings.txtLvCp)
        ClearToolStripMenuItem2.Checked = Settings.cArtCache
        AttrReadOnlyToolStripMenuItem.Checked = Settings.proRO
    End Sub

    Private Sub Main_Form_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        Me.Text &= My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build & " - by " & My.Application.Info.Copyright
        LinkLabel1.Text = Settings.GameDATA_Path
        LinkLabel2.Text = Settings.SaveMOD_Path
        SetFormSettings()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles Timer1.Tick
        SplashScreen.Hide()
        Timer1.Stop()
        Me.Focus()
    End Sub

    Private Sub Main_Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        TabControl1.Visible = True
        SplitContainer1.SplitterDistance = SplitSize
    End Sub

    Private Sub Main_Form_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If cCache Then
                Clear_Cache()
            ElseIf cArtCache Then
                Clear_Art_Cache()
            End If
            Settings.Save_Config()
            File.Delete("iProto.lst")
            File.Delete("cProto.lst")
            Application.Exit()
        End If
    End Sub

    Private Sub AddCritterPro()
        Dim pCont As Integer = ListView1.Items.Count
        Dim pName As String = StrDup(8 - (pCont + 1).ToString.Length, "0") & (pCont + 1).ToString & ".pro"
        Dim ffile As Integer = FreeFile()

        FileOpen(ffile, "template", OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
        FilePut(ffile, ReverseBytes(&H1000001I + pCont))
        FilePut(ffile, ReverseBytes((pCont + 1) * 100))
        FileClose(ffile)

        ProFiles.CreateProFile(PROTO_CRITTERS, pName)

        Array.Resize(Critter_LST, pCont + 1)
        Critter_LST(pCont).proFile = pName
        Dim lst(pCont) As String
        For n = 0 To pCont
            lst(n) = Critter_LST(n).proFile
        Next
        File.WriteAllLines(SaveMOD_Path & crittersLstPath, lst)

        'Log 
        TextBox1.Text = "Update: " & SaveMOD_Path & crittersLstPath & vbCrLf & TextBox1.Text

        ListView1.Items.Add("<NoName>")
        ListView1.Items(pCont).SubItems.Add(pName)
        ListView1.Items(pCont).SubItems.Add("N/A")
        ListView1.Items(pCont).SubItems.Add(&H1000001 + pCont)
        ListView1.Items(pCont).Selected = True
        ListView1.EnsureVisible(pCont)

        Main.Create_CritterForm(pCont)
    End Sub

    Private Sub AddItemPro(ByVal iType As Integer)
        Dim pCont As Integer = Items_LST.Length
        Dim pName As String = StrDup(8 - (pCont + 1).ToString.Length, "0") & (pCont + 1).ToString & ".pro"
        Dim ffile As Integer = FreeFile()

        FileOpen(ffile, "template", OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
        FilePut(ffile, ReverseBytes(&H1I + pCont))
        FilePut(ffile, ReverseBytes((pCont + 1) * 100))
        FileClose(ffile)

        ProFiles.CreateProFile(PROTO_ITEMS, pName)

        Array.Resize(Items_LST, pCont + 1)
        Items_LST(pCont).proFile = pName
        Items_LST(pCont).itemType = iType

        'save to lst file
        Dim lst(pCont) As String
        For n = 0 To pCont
            lst(n) = Items_LST(n).proFile
        Next
        File.WriteAllLines(SaveMOD_Path & itemsLstPath, lst)

        'Log 
        TextBox1.Text = "Update: " & SaveMOD_Path & itemsLstPath & vbCrLf & TextBox1.Text

        ListView2.Items.Add("<No Name>")
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(pName)
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(ItemTypesName(Items_LST(pCont).itemType))
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add("N/A")
        ListView2.Items(ListView2.Items.Count - 1).Selected = True
        ListView2.Items(ListView2.Items.Count - 1).Tag = pCont
        ListView2.EnsureVisible(ListView2.Items.Count - 1)

        Main.Create_ItemsForm(pCont)
    End Sub

    Private Sub DelToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If MsgBox("Delete Pro File?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            If TabControl1.SelectedIndex = 0 Then
                'del file
                Dim dProFile As String = SaveMOD_Path & PROTO_ITEMS & ListView2.FocusedItem.SubItems(1).Text
                If File.Exists(dProFile) Then
                    File.SetAttributes(dProFile, FileAttributes.Normal)
                    File.Delete(dProFile)

                    'Log 
                    TextBox1.Text = "Delete Pro: " & dProFile & vbCrLf & TextBox1.Text
                End If
                Dim pCont As Integer = UBound(Items_LST)
                If CInt(ListView2.FocusedItem.Tag) = pCont Then
                    Array.Resize(Items_LST, pCont)

                    'save to lst file
                    pCont -= 1
                    Dim lst(pCont) As String
                    For n As Integer = 0 To pCont
                        lst(n) = Items_LST(n).proFile
                    Next
                    File.WriteAllLines(SaveMOD_Path & itemsLstPath, lst)
                    ListView2.Items.RemoveAt(ListView2.FocusedItem.Index)

                    'Log 
                    TextBox1.Text = "Update: " & SaveMOD_Path & itemsLstPath & vbCrLf & TextBox1.Text
                Else
                    ListView2.Items.Item(ListView2.FocusedItem.Index).Text = "- " & Items_LST(ListView2.FocusedItem.Tag).itemName
                End If
            Else
                'del file
                Dim dProFile As String = SaveMOD_Path & PROTO_CRITTERS & ListView1.FocusedItem.SubItems(1).Text
                If File.Exists(dProFile) Then
                    File.SetAttributes(dProFile, FileAttributes.Normal)
                    File.Delete(dProFile)

                    'Log 
                    TextBox1.Text = "Delete Pro: " & dProFile & vbCrLf & TextBox1.Text
                End If
                Dim pCont As Integer = UBound(Critter_LST)
                If ListView1.FocusedItem.Index = pCont Then
                    ListView1.Items.RemoveAt(pCont)
                    Array.Resize(Critter_LST, pCont)
                    pCont -= 1

                    'save to lst file
                    Dim lst(pCont) As String
                    For n = 0 To pCont
                        lst(n) = Critter_LST(n).proFile
                    Next
                    File.WriteAllLines(SaveMOD_Path & crittersLstPath, lst)

                    'Log 
                    TextBox1.Text = "Update: " & SaveMOD_Path & crittersLstPath & vbCrLf & TextBox1.Text
                End If
            End If
        End If
    End Sub

    Private Sub CreateToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CreateToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Dim pIndx As Integer = ListView1.FocusedItem.Index

            FileSystem.CopyFile(DatFiles.CheckFile(PROTO_CRITTERS & Critter_LST(pIndx).proFile), "template", True)
            File.SetAttributes("template", FileAttributes.Normal Or FileAttributes.Archive)

            AddCritterPro()
        Else
            Dim pIndx As Integer = CInt(ListView2.FocusedItem.Tag)

            FileSystem.CopyFile(DatFiles.CheckFile(PROTO_ITEMS & Items_LST(pIndx).proFile), "template", True)
            File.SetAttributes("template", FileAttributes.Normal Or FileAttributes.Archive)

            AddItemPro(Items_LST(pIndx).itemType)
        End If
    End Sub

    'поиск ключевого слова
    Private Sub Find(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripButton1.Click, FindToolStripMenuItem.Click
        Dim n As Integer
        Dim LIST_VIEW As ListView = ListView1

        If ToolStripTextBox1.Text <> Nothing Then
            If TabControl1.SelectedIndex = 0 Then LIST_VIEW = ListView2
            LIST_VIEW.Focus()

            n = LIST_VIEW.FocusedItem.Index + 1
            If n >= LIST_VIEW.Items.Count Then n = 0
            If SearhLW(n, LIST_VIEW) >= LIST_VIEW.Items.Count Then
                If SearhLW(0, LIST_VIEW) >= LIST_VIEW.Items.Count Then Exit Sub
            End If

            LIST_VIEW.TopItem = LIST_VIEW.FocusedItem
            If LIST_VIEW.FocusedItem.Index > 10 Then
                LIST_VIEW.EnsureVisible(LIST_VIEW.FocusedItem.Index - 10)
            Else
                LIST_VIEW.EnsureVisible(LIST_VIEW.FocusedItem.Index)
            End If
        End If
    End Sub

    Private Function SearhLW(ByRef n As Integer, ByRef LIST_VIEW As ListView) As Integer
        For n = n To LIST_VIEW.Items.Count - 1
            If LIST_VIEW.Items(n).Text.IndexOf(ToolStripTextBox1.Text.Trim, StringComparison.OrdinalIgnoreCase) <> -1 Then
                LIST_VIEW.Items.Item(n).Selected = True
                LIST_VIEW.Items.Item(n).Focused = True
                Exit For
            End If
        Next

        Return n
    End Function

    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        If e.KeyData = Keys.Enter Then Find(Nothing, Nothing)
    End Sub

    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        Main.Create_CritterForm(ListView1.FocusedItem.Index)
    End Sub

    Private Sub ListView2_MouseDoubleClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView2.MouseDoubleClick
        Main.Create_ItemsForm(ListView2.FocusedItem.Tag)
    End Sub

    Private Sub HToolStripMenuItem_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles HToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Main.Create_CritterForm(ListView1.FocusedItem.Index)
        Else
            Main.Create_ItemsForm(ListView2.FocusedItem.Tag)
        End If
    End Sub

    Private Sub ToolStripSplitButton1_MouseHover(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripSplitButton1.MouseHover
        ToolStripSplitButton1.ShowDropDown()
    End Sub

    Private Sub ToolStripSplitButton2_ButtonClick(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripSplitButton2.ButtonClick
        Setting_Form.ShowDialog()
        Settings.SetEncoding()
        SetFormSettings()
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripButton9.Click
        If TabControl1.SelectedIndex = 1 Then
            Main.CreateCritterList()
        Else
            ClearFilter()
            fAllToolStripMenuItem1.Checked = True
            Main.CreateItemsList()
        End If

        Main.GetScriptLst()
    End Sub

    Private Sub AttrReadOnlyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AttrReadOnlyToolStripMenuItem.Click
        proRO = AttrReadOnlyToolStripMenuItem.Checked
    End Sub

    Private Sub ClearToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ClearToolStripMenuItem2.Click
        cArtCache = ClearToolStripMenuItem2.Checked
    End Sub

    Private Sub Cp886ToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Cp886ToolStripMenuItem.Click
        txtWin = Not (Cp886ToolStripMenuItem.Checked)
        Setting_Form.RadioButton1.Checked = txtWin
        Settings.SetEncoding()
        txtLvCp = False
    End Sub

    Private Sub ClearFilter()
        fAllToolStripMenuItem1.Checked = False
        fWeaponToolStripMenuItem3.Checked = False
        fAmmoToolStripMenuItem2.Checked = False
        fArmorToolStripMenuItem2.Checked = False
        fDrugToolStripMenuItem3.Checked = False
        fMiscToolStripMenuItem2.Checked = False
        fContainerToolStripMenuItem.Checked = False

        ListView2.ListViewItemSorter = Nothing
    End Sub

    Private Sub fAllToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fAllToolStripMenuItem1.Click
        ClearFilter()
        fAllToolStripMenuItem1.Checked = True
        Main.FilterCreateItemsList(ItemType.Unknown)
    End Sub

    Private Sub fWeaponToolStripMenuItem3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fWeaponToolStripMenuItem3.Click
        ClearFilter()
        fWeaponToolStripMenuItem3.Checked = True
        Main.FilterCreateItemsList(ItemType.Weapon)
    End Sub

    Private Sub fAmmoToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fAmmoToolStripMenuItem2.Click
        ClearFilter()
        fAmmoToolStripMenuItem2.Checked = True
        Main.FilterCreateItemsList(ItemType.Ammo)
    End Sub

    Private Sub fArmorToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fArmorToolStripMenuItem2.Click
        ClearFilter()
        fArmorToolStripMenuItem2.Checked = True
        Main.FilterCreateItemsList(ItemType.Armor)
    End Sub

    Private Sub fDrugToolStripMenuItem3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fDrugToolStripMenuItem3.Click
        ClearFilter()
        fDrugToolStripMenuItem3.Checked = True
        Main.FilterCreateItemsList(ItemType.Drugs)
    End Sub

    Private Sub fMiscToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fMiscToolStripMenuItem2.Click
        ClearFilter()
        fMiscToolStripMenuItem2.Checked = True
        Main.FilterCreateItemsList(ItemType.Misc)
    End Sub

    Private Sub fContainerToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fContainerToolStripMenuItem.Click
        ClearFilter()
        fContainerToolStripMenuItem.Checked = True
        Main.FilterCreateItemsList(ItemType.Container)
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As TabControlCancelEventArgs) Handles TabControl1.Selecting
        If TabControl1.SelectedIndex = 0 Then
            ToolStripSplitButton1.Enabled = True
        Else
            ToolStripSplitButton1.Enabled = False
        End If
        If ListView1.Items.Count = 0 Then CreateCritterList() ': TypeCrittersToolStripMenuItem.Enabled = True
    End Sub

    Private Sub ListView2_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs) Handles ListView2.ColumnClick
        ListView2.ListViewItemSorter = New Comparer.ListViewItemComparer(e.Column)
    End Sub

    Private Sub ListView1_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView1.MouseClick
        Dim checkFile As String = PROTO_CRITTERS & ListView1.FocusedItem.SubItems(1).Text

        If e.Button = Forms.MouseButtons.Right AndAlso ListView1.FocusedItem.Index = ListView1.Items.Count - 1 _
            Or (File.Exists(SaveMOD_Path & checkFile) And (File.Exists(GameDATA_Path & checkFile) Or File.Exists(Cache_Patch & checkFile))) Then
            DeleteToolStripMenuItem.Enabled = True
        Else
            DeleteToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub ListView2_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView2.MouseClick
        Dim checkFile As String = PROTO_ITEMS & ListView2.FocusedItem.SubItems(1).Text
        If (e.Button = Forms.MouseButtons.Right AndAlso CInt(ListView2.FocusedItem.Tag) = UBound(Items_LST)) _
            Or (File.Exists(SaveMOD_Path & checkFile) And (File.Exists(GameDATA_Path & checkFile) Or File.Exists(Cache_Patch & checkFile))) Then
            DeleteToolStripMenuItem.Enabled = True
        Else
            DeleteToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub MainGotFogus(ByVal sender As Object, ByVal e As EventArgs) Handles ListView2.MouseEnter, ListView1.MouseEnter
        If Not (DontHoverSelectToolStripMenuItem.Checked) AndAlso Not (Me.Focused) Then Me.Focus()
    End Sub

    Private Sub ListView1_ItemSelectionChanged(ByVal sender As Object, ByVal e As ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        On Error Resume Next
        ToolStripStatusLabel1.Text = DatFiles.CheckFile(PROTO_CRITTERS & Critter_LST(e.ItemIndex).proFile, , , False)
        ToolStripStatusLabel2.Text = "Critter PID: " & &H1000001 + e.ItemIndex
    End Sub

    Private Sub ListView2_ItemSelectionChanged(ByVal sender As Object, ByVal e As ListViewItemSelectionChangedEventArgs) Handles ListView2.ItemSelectionChanged
        On Error Resume Next
        ToolStripStatusLabel1.Text = DatFiles.CheckFile(PROTO_ITEMS & Items_LST(e.Item.Tag).proFile, , , False)
        ToolStripStatusLabel2.Text = "Item PID: " & StrDup(8 - Len(CStr(CInt(e.Item.Tag) + 1)), "0") & CInt(e.Item.Tag) + 1
    End Sub

    Private Sub TypeCrittersToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TypeCrittersToolStripMenuItem.Click
        TabControl1.SelectTab(1)
        Application.DoEvents()

        If ListView1.Items.Count = 0 Then Main.CreateCritterList()
        File.WriteAllBytes("template", My.Resources.critter)

        AddCritterPro()
    End Sub

    Private Sub pWeaponToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pWeaponToolStripMenuItem2.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.weapon)
        AddItemPro(ItemType.Weapon)
    End Sub

    Private Sub pAmmoToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pAmmoToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.ammo)
        AddItemPro(ItemType.Ammo)
    End Sub

    Private Sub pArmorToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pArmorToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.armor)
        AddItemPro(ItemType.Armor)
    End Sub

    Private Sub pDrugToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pDrugToolStripMenuItem2.Click
        File.WriteAllBytes("template", My.Resources.drugs)
        AddItemPro(ItemType.Drugs)
    End Sub

    Private Sub ContainerToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ContainerToolStripMenuItem.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.container)
        AddItemPro(ItemType.Container)
    End Sub

    Private Sub pMiscToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pMiscToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.misc)
        AddItemPro(ItemType.Misc)
    End Sub

    Private Sub KeyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles KeyToolStripMenuItem.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()

        File.WriteAllBytes("template", My.Resources.key)
        AddItemPro(ItemType.Key)
    End Sub

    Private Sub AboutToolStripButton7_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AboutToolStripButton7.Click
        AboutBox.ShowDialog()
    End Sub

    Private Sub ListView2_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView2.MouseDown
        If DontHoverSelectToolStripMenuItem.Checked AndAlso e.Button = Windows.Forms.MouseButtons.Middle Then
            ListView2.HoverSelection = True
        End If
    End Sub

    Private Sub ListView2_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView2.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Middle Then
            Create_ItemsForm(ListView2.FocusedItem.Tag)
            ListView2.HoverSelection = Not DontHoverSelectToolStripMenuItem.Checked
        End If
    End Sub

    Private Sub ListView1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView1.MouseDown
        If DontHoverSelectToolStripMenuItem.Checked AndAlso e.Button = Windows.Forms.MouseButtons.Middle Then
            ListView1.HoverSelection = True
        End If
    End Sub

    Private Sub ListView1_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles ListView1.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Middle Then
            Main.Create_CritterForm(ListView1.FocusedItem.Index)
            ListView1.HoverSelection = Not DontHoverSelectToolStripMenuItem.Checked
        End If
    End Sub

    Private Sub ImportCritterTableToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ImportCritterTableToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then Exit Sub
        TableLog_Form.ListBox1.Items.Clear()
        Table_Form.Critters_ImportTable(OpenFileDialog1.FileName)
    End Sub

    Private Sub ImportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ImportToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then Exit Sub
        TableLog_Form.ListBox1.Items.Clear()
        Table_Form.Items_ImportTable(OpenFileDialog1.FileName)
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExportToolStripMenuItem.Click
        SetParent(Table_Form.Handle.ToInt32, SplitContainer1.Handle.ToInt32)
        Table_Form.Show()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim cpath As String

        If TabControl1.SelectedIndex = 1 Then
            Dim pIndx As Integer = ListView1.FocusedItem.Index
            Dim pFile As String = Critter_LST(pIndx).proFile

            cpath = DatFiles.CheckFile(PROTO_CRITTERS & pFile, False)
            Hex_Form.LoadHex(cpath & PROTO_CRITTERS, pFile)
            Hex_Form.Tag = cpath & PROTO_CRITTERS & pFile
        Else
            Dim pFile As String = ListView2.FocusedItem.SubItems(1).Text

            cpath = DatFiles.CheckFile(PROTO_ITEMS & pFile, False)
            Hex_Form.LoadHex(cpath & PROTO_ITEMS, pFile)
            Hex_Form.Tag = cpath & PROTO_ITEMS & pFile
        End If
    End Sub

    Private Sub LinkLabel_LinkClicked(ByVal sender As Object, ByVal e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked, LinkLabel2.LinkClicked
        Dim link As LinkLabel = sender
        Process.Start("explorer", link.Text)
    End Sub

    Private Sub Main_Form_Resize(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Resize
        TextBox1.Height = Me.Height / 4
    End Sub

    Private Sub TextEditProFileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TextEditProFileToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Create_TxtEditForm(ListView1.FocusedItem.Index, 0)
        Else
            Create_TxtEditForm(ListView2.FocusedItem.Tag, 1)
        End If
    End Sub

    Private Sub ViewLogToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ViewLogToolStripMenuItem.Click
        TextBox1.Visible = ViewLogToolStripMenuItem.Checked
    End Sub

    Private Sub AIPacketToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AIPacketToolStripMenuItem.Click
        Main.Create_AIEditForm()
    End Sub

    Private Sub DontHoverSelectToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DontHoverSelectToolStripMenuItem.Click
        Settings.HoverSelect = Not (DontHoverSelectToolStripMenuItem.Checked)
        ListView1.HoverSelection = Settings.HoverSelect
        ListView2.HoverSelection = Settings.HoverSelect
        If DontHoverSelectToolStripMenuItem.Checked Then
            ListView1.Activation = ItemActivation.Standard
            ListView2.Activation = ItemActivation.Standard
        Else
            ListView1.Activation = ItemActivation.OneClick
            ListView2.Activation = ItemActivation.OneClick
        End If
    End Sub

    Private Sub MassCreateProfiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MassCreateProfilesToolStripMenuItem.Click
        Dim MassCreateFrm As New MassCreate()
    End Sub

End Class
