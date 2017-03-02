
Public Class Main_Form

    'Private ClickXClose As Boolean = False
    'Private Const SC_CLOSE As Int32 = &HF060
    '
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    ' Нажата кнопка Close.
    '    If m.WParam.ToInt32() = SC_CLOSE Then ClickXClose = True
    '    MyBase.WndProc(m)
    'End Sub

    ' Implements the manual sorting of items by columns.
    Class ListViewItemComparer
        Implements IComparer
        Private col As Integer
        Public Sub New()
            col = 0
        End Sub
        Public Sub New(ByVal column As Integer)
            col = column
        End Sub
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            On Error Resume Next
            Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
        End Function
    End Class

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        SplashScreen.Hide()
        Timer1.Stop()
        Me.Focus()
    End Sub

    Private Sub Main_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Main_Module.Main()
        Timer1.Start()
        TabControl1.Visible = True
        SplitContainer1.SplitterDistance = SplitSize
    End Sub

    Private Sub Main_Form_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Cp886ToolStripMenuItem.Checked = Not (txtWin) And Not (txtLvCp)
        AttrReadOnlyToolStripMenuItem.Checked = proRO
        ClearToolStripMenuItem2.Checked = cArtCache
        Me.Text &= My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build & "  by " & My.Application.Info.Copyright
        ')String.Format("{#,##0.00;}", 
    End Sub

    Private Sub Main_Form_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If cCache Then Clear_Cache()
            If cArtCache AndAlso cCache = False Then Clear_Art_Cache()
            Setting_Form.Setting_Form_Load(sender, e)
            Setting_Form.Store_Ini()
            Application.Exit()
        End If
    End Sub

    Private Sub ListView1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDoubleClick
        Create_CritterForm(ListView1.FocusedItem.Index)
    End Sub

    Private Sub ListView2_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView2.MouseDoubleClick
        Create_ItemsForm(ListView2.FocusedItem.Tag)
    End Sub

    Private Sub HToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Create_CritterForm(ListView1.FocusedItem.Index)
        Else
            Create_ItemsForm(ListView2.FocusedItem.Tag)
        End If
    End Sub

    Private Sub ToolStripSplitButton1_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripSplitButton1.MouseHover
        ToolStripSplitButton1.ShowDropDown()
    End Sub

    Private Sub ToolStripSplitButton2_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripSplitButton2.ButtonClick
        Setting_Form.ShowDialog()
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        If TabControl1.SelectedIndex = 1 Then
            ListView1.Items.Clear()
            ListView1.Visible = False
            ReDim Critter_LST(0)
            CreateCritterList()
        Else
            ClearFilter()
            fAllToolStripMenuItem1.Checked = True
            ReDim Items_LST(0, 0)
            CreateItemsList()
        End If
        GetScriptLst()
    End Sub

    Private Sub AttrReadOnlyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AttrReadOnlyToolStripMenuItem.Click
        proRO = AttrReadOnlyToolStripMenuItem.Checked
    End Sub

    Private Sub ClearToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem2.Click
        cArtCache = ClearToolStripMenuItem2.Checked
    End Sub

    Private Sub Cp886ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cp886ToolStripMenuItem.Click
        txtWin = Not (Cp886ToolStripMenuItem.Checked)
        Setting_Form.RadioButton1.Checked = txtWin
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
        ListView2.Items.Clear()
        ListView2.Visible = False
    End Sub

    Private Sub fAllToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fAllToolStripMenuItem1.Click
        ClearFilter()
        fAllToolStripMenuItem1.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fWeaponToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fWeaponToolStripMenuItem3.Click
        ClearFilter()
        fWeaponToolStripMenuItem3.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fAmmoToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fAmmoToolStripMenuItem2.Click
        ClearFilter()
        fAmmoToolStripMenuItem2.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fArmorToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fArmorToolStripMenuItem2.Click
        ClearFilter()
        fArmorToolStripMenuItem2.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fDrugToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fDrugToolStripMenuItem3.Click
        ClearFilter()
        fDrugToolStripMenuItem3.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fMiscToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fMiscToolStripMenuItem2.Click
        ClearFilter()
        fMiscToolStripMenuItem2.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub fContainerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fContainerToolStripMenuItem.Click
        ClearFilter()
        fContainerToolStripMenuItem.Checked = True
        FilterCreateItemsList()
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        If TabControl1.SelectedIndex = 0 Then ToolStripSplitButton1.Enabled = True Else ToolStripSplitButton1.Enabled = False
        If ListView1.Items.Count = 0 Then CreateCritterList() ': TypeCrittersToolStripMenuItem.Enabled = True
    End Sub

    Private Sub ListView2_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView2.ColumnClick
        ListView2.ListViewItemSorter = New ListViewItemComparer(e.Column)
    End Sub

    Private Sub ListView1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right AndAlso ListView1.FocusedItem.Index = ListView1.Items.Count - 1 _
            Or (IO.File.Exists(SaveMOD_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text) _
            And (IO.File.Exists(GameDATA_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text) _
            Or IO.File.Exists(Cache_Patch & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text))) Then
            DeleteToolStripMenuItem.Enabled = True
        Else
            DeleteToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub ListView2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView2.MouseClick
        If (e.Button = Windows.Forms.MouseButtons.Right AndAlso ListView2.FocusedItem.Tag = UBound(Items_LST)) _
            Or (IO.File.Exists(SaveMOD_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text) _
            And (IO.File.Exists(GameDATA_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text) _
            Or IO.File.Exists(Cache_Patch & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text))) Then
            DeleteToolStripMenuItem.Enabled = True
        Else
            DeleteToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub ListView2_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.MouseEnter
        If Not (Me.Focused) Then Me.Focus()
    End Sub

    Private Sub ListView1_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.MouseEnter
        If Not (Me.Focused) Then Me.Focus()
    End Sub

    Private Sub ListView1_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        On Error Resume Next
        ToolStripStatusLabel1.Text = Check_File("proto\critters\" & Critter_LST(e.ItemIndex), , False) & "\proto\critters\" & Critter_LST(e.ItemIndex)
        ToolStripStatusLabel2.Text = "Critter PID: " & &H1000001 + e.ItemIndex
    End Sub

    Private Sub ListView2_ItemSelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles ListView2.ItemSelectionChanged
        On Error Resume Next
        ToolStripStatusLabel1.Text = Check_File("proto\items\" & Items_LST(e.Item.Tag, 0), , False) & "\proto\items\" & Items_LST(e.Item.Tag, 0)
        ToolStripStatusLabel2.Text = "Item PID: " & StrDup(8 - Len(CStr(e.Item.Tag + 1)), "0") & e.Item.Tag + 1
    End Sub

    Private Sub TypeCrittersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TypeCrittersToolStripMenuItem.Click
        TabControl1.SelectTab(1)
        Application.DoEvents()
        If ListView1.Items.Count = 0 Then CreateCritterList()
        IO.File.WriteAllBytes("template", My.Resources.critter)
        AddCritterPro()
    End Sub

    Private Sub AddCritterPro()
        Dim pCont As UShort = ListView1.Items.Count
        Dim pName As String = StrDup(8 - (pCont + 1).ToString.Length, "0") & (pCont + 1).ToString & ".pro"
        Dim ffile As Byte = FreeFile()
        FileOpen(ffile, "template", OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
        FilePut(ffile, ReverseBytes(&H1000001I + pCont))
        FilePut(ffile, ReverseBytes((pCont + 1) * 100))
        FileClose(ffile)
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & pName) Then
            IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & pName, IO.FileAttributes.Normal)
            IO.File.Delete(SaveMOD_Path & "\proto\critters\" & pName)
        End If
        IO.File.Move("template", SaveMOD_Path & "\proto\critters\" & pName)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & pName, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        'Log 
        TextBox1.Text = "Create: " & SaveMOD_Path & "\proto\critters\" & pName & vbCrLf & TextBox1.Text
        '
        ReDim Preserve Critter_LST(pCont)
        ReDim Preserve Critter_NAME(pCont)
        Critter_LST(pCont) = pName
        IO.File.WriteAllLines(SaveMOD_Path & "\proto\critters\critters.lst", Critter_LST)
        'Log 
        TextBox1.Text = "Update: " & SaveMOD_Path & "\proto\critters\critters.lst" & vbCrLf & TextBox1.Text
        '
        ListView1.Items.Add("<No Name>")
        ListView1.Items(pCont).SubItems.Add(pName)
        ListView1.Items(pCont).SubItems.Add("N/A")
        ListView1.Items(pCont).SubItems.Add(&H1000001 + pCont)
        ListView1.Items(pCont).Selected = True
        ListView1.EnsureVisible(pCont)
        '
        Create_CritterForm(pCont)
    End Sub

    Private Sub pWeaponToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pWeaponToolStripMenuItem2.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.weapon)
        AddItemPro("Weapon")
    End Sub

    Private Sub pAmmoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pAmmoToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.ammo)
        AddItemPro("Ammo")
    End Sub

    Private Sub pArmorToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pArmorToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.armor)
        AddItemPro("Armor")
    End Sub

    Private Sub pDrugToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pDrugToolStripMenuItem2.Click
        IO.File.WriteAllBytes("template", My.Resources.drugs)
        AddItemPro("Drugs")
    End Sub

    Private Sub ContainerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContainerToolStripMenuItem.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.container)
        AddItemPro("Container")
    End Sub

    Private Sub pMiscToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pMiscToolStripMenuItem1.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.misc)
        AddItemPro("Misc")
    End Sub

    Private Sub KeyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeyToolStripMenuItem.Click
        TabControl1.SelectTab(0)
        Application.DoEvents()
        IO.File.WriteAllBytes("template", My.Resources.key)
        AddItemPro("Key")
    End Sub

    Private Sub AddItemPro(ByVal iType As String)
        Dim pCont As UShort = UBound(Items_LST) + 1 'ListView2.Items.Count
        Dim pName As String = StrDup(8 - (pCont + 1).ToString.Length, "0") & (pCont + 1).ToString & ".pro"
        Dim ffile As Byte = FreeFile()
        FileOpen(ffile, "template", OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
        FilePut(ffile, ReverseBytes(&H1I + pCont))
        FilePut(ffile, ReverseBytes((pCont + 1) * 100))
        FileClose(ffile)
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & pName) Then
            IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & pName, IO.FileAttributes.Normal)
            IO.File.Delete(SaveMOD_Path & "\proto\items\" & pName)
        End If
        IO.File.Move("template", SaveMOD_Path & "\proto\items\" & pName)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & pName, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        'Log 
        TextBox1.Text = "Create: " & SaveMOD_Path & "\proto\items\" & pName & vbCrLf & TextBox1.Text
        '
        Dim a(,) As String = Items_LST
        ReDim Items_LST(pCont, 1)
        ReDim Preserve Items_NAME(pCont)
        ffile = FreeFile()
        FileOpen(ffile, SaveMOD_Path & "\proto\items\items.lst", OpenMode.Output)
        Dim n As UShort
        For n = 0 To pCont - 1
            Items_LST(n, 0) = a(n, 0) : Items_LST(n, 1) = a(n, 1)
            PrintLine(ffile, a(n, 0))
        Next
        PrintLine(ffile, pName)
        FileClose(ffile)
        'Log 
        TextBox1.Text = "Update: " & SaveMOD_Path & "\proto\items\items.lst" & vbCrLf & TextBox1.Text
        '
        Items_LST(pCont, 0) = pName : Items_LST(pCont, 1) = iType
        '
        ListView2.Items.Add("<No Name>")
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(pName)
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(Items_LST(pCont, 1))
        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add("N/A")
        ListView2.Items(ListView2.Items.Count - 1).Selected = True
        ListView2.Items(ListView2.Items.Count - 1).Tag = pCont
        ListView2.EnsureVisible(ListView2.Items.Count - 1)
        '
        Create_ItemsForm(pCont)
    End Sub

    Private Sub DelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If MsgBox("Delete Pro File?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            If TabControl1.SelectedIndex = 0 Then
                'del file
                If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text) Then
                    IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text, IO.FileAttributes.Normal)
                    IO.File.Delete(SaveMOD_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text)
                    'Log 
                    TextBox1.Text = "Delete: " & SaveMOD_Path & "\proto\items\" & ListView2.FocusedItem.SubItems(1).Text & vbCrLf & TextBox1.Text
                End If
                Dim pCont As UShort = UBound(Items_LST)
                If ListView2.FocusedItem.Tag = pCont Then
                    Dim a(,) As String = Items_LST
                    ReDim Items_LST(pCont - 1, 1)
                    ReDim Preserve Items_NAME(pCont - 1)
                    Dim ffile As Byte = FreeFile()
                    FileOpen(ffile, SaveMOD_Path & "\proto\items\items.lst", OpenMode.Output)
                    For n As UShort = 0 To pCont - 1
                        Items_LST(n, 0) = a(n, 0) : Items_LST(n, 1) = a(n, 1)
                        PrintLine(ffile, a(n, 0))
                    Next
                    FileClose(ffile)
                    ListView2.Items.RemoveAt(ListView2.FocusedItem.Index)
                    'Log 
                    TextBox1.Text = "Update: " & SaveMOD_Path & "\proto\items\items.lst" & vbCrLf & TextBox1.Text
                Else
                    ListView2.Items.Item(ListView2.FocusedItem.Index).Text = "- " & Items_NAME(ListView2.FocusedItem.Tag)
                End If
            Else
                'del file
                If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text) Then
                    IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text, IO.FileAttributes.Normal)
                    IO.File.Delete(SaveMOD_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text)
                    'Log 
                    TextBox1.Text = "Delete: " & SaveMOD_Path & "\proto\critters\" & ListView1.FocusedItem.SubItems(1).Text & vbCrLf & TextBox1.Text
                End If
                Dim pCont As UShort = UBound(Critter_LST)
                If ListView1.FocusedItem.Index = pCont Then
                    ReDim Preserve Critter_LST(pCont - 1)
                    ReDim Preserve Critter_NAME(pCont - 1)
                    IO.File.WriteAllLines(SaveMOD_Path & "\proto\critters\critters.lst", Critter_LST)
                    ListView1.Items.RemoveAt(pCont)
                    'Log 
                    TextBox1.Text = "Update: " & SaveMOD_Path & "\proto\critters\critters.lst" & vbCrLf & TextBox1.Text
                End If
            End If
        End If
    End Sub

    Private Sub CreateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Dim pIndx As UShort = ListView1.FocusedItem.Index
            Current_Path = Check_File("proto\critters\" & Critter_LST(pIndx))
            My.Computer.FileSystem.CopyFile(Current_Path & "\proto\critters\" & Critter_LST(pIndx), "template", True)
            IO.File.SetAttributes("template", IO.FileAttributes.Normal Or IO.FileAttributes.Archive)
            AddCritterPro()
        Else
            Dim pIndx As UShort = ListView2.FocusedItem.Tag
            Current_Path = Check_File("proto\items\" & Items_LST(pIndx, 0))
            My.Computer.FileSystem.CopyFile(Current_Path & "\proto\items\" & Items_LST(pIndx, 0), "template", True)
            IO.File.SetAttributes("template", IO.FileAttributes.Normal Or IO.FileAttributes.Archive)
            AddItemPro(Items_LST(pIndx, 1))
        End If
    End Sub

    Private Sub AboutToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripButton7.Click
        AboutBox.ShowDialog()
    End Sub

    Private Sub ListView2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView2.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Middle Then Create_ItemsForm(ListView2.FocusedItem.Tag)
    End Sub

    Private Sub ListView1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Middle Then Create_CritterForm(ListView1.FocusedItem.Index)
    End Sub

    Private Sub ImportCritterTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportCritterTableToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then Exit Sub
        Table_Form.Import_CR_Table(OpenFileDialog1.FileName)
    End Sub

    Private Sub ImportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportToolStripMenuItem.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then Exit Sub
        Table_Form.Import_Table(OpenFileDialog1.FileName)
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        SetParent(Table_Form.Handle.ToInt32, SplitContainer1.Panel1.Handle.ToInt32)
        Table_Form.Show()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If TabControl1.SelectedIndex = 1 Then
            Dim pIndx As UShort = ListView1.FocusedItem.Index
            Current_Path = Check_File("proto\critters\" & Critter_LST(pIndx))
            'SetParent(Hex_Form.Handle.ToInt32, SplitContainer1.Panel1.Handle.ToInt32)
            Hex_Form.LoadHex(Current_Path & "\proto\critters\", Critter_LST(pIndx))
            Hex_Form.Tag = Current_Path & "\proto\critters\" & Critter_LST(pIndx)
        Else
            Dim pFile As String = ListView2.FocusedItem.SubItems(1).Text
            Current_Path = Check_File("proto\items\" & pFile)
            Hex_Form.LoadHex(Current_Path & "\proto\items\", pFile)
            Hex_Form.Tag = Current_Path & "\proto\items\" & pFile
        End If
    End Sub

    'поиск ключевого слова
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim n As UShort
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

    Private Function SearhLW(ByRef n As UShort, ByRef LIST_VIEW As ListView)
        For n = n To LIST_VIEW.Items.Count - 1
            If LIST_VIEW.Items(n).Text.IndexOf(ToolStripTextBox1.Text.Trim, StringComparison.OrdinalIgnoreCase) <> -1 Then
                LIST_VIEW.Items.Item(n).Selected = True
                LIST_VIEW.Items.Item(n).Focused = True
                Exit For
            End If
        Next
        Return n
    End Function

    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        If e.KeyData = Keys.Enter Then ToolStripButton1_Click(sender, e)
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("explorer", LinkLabel1.Text)
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("explorer", LinkLabel2.Text)
    End Sub

    Private Sub Main_Form_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        TextBox1.Height = Me.Height / 4
    End Sub

    Private Sub TextEditProFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextEditProFileToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 1 Then
            Create_TxtEditForm(ListView1.FocusedItem.Index, 0)
        Else
            Create_TxtEditForm(ListView2.FocusedItem.Tag, 1)
        End If
    End Sub

    Private Sub ViewLogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewLogToolStripMenuItem.Click
        TextBox1.Visible = ViewLogToolStripMenuItem.Checked
    End Sub

End Class
