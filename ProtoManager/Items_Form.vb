Imports System.Drawing
Imports System.IO
Imports System.Text
Imports Prototypes

Friend Class Items_Form

    Private CommonItem As CmItemPro
    Private WeaponItem As WpItemPro
    Private ArmorItem As ArItemPro
    Private AmmoItem As AmItemPro
    Private KeyItem As kItemPro
    Private DrugItem As DgItemPro
    Private MiscItem As McItemPro
    Private ContanerItem As CnItemPro
    '
    Private ReadOnly iLST_Index As Integer  'индекс предмета в lst файле
    Private frmReady As Boolean
    Private ReloadPro As Boolean
    Private cPath As String = Nothing

    Sub New(ByVal iLST_Index As Integer)
        InitializeComponent()
        Me.iLST_Index = iLST_Index
    End Sub

    Friend Sub IniItemsForm()
        ComboBox10.Items.AddRange(Misc_NAME)
        ComboBox11.Items.AddRange(Perk_NAME)
        ComboBox12.Items.AddRange(CaliberNAME)

        For n = 0 To UBound(AmmoPID)
            ComboBox13.Items.Add(AmmoNAME(n))
        Next

        If Critters_FRM IsNot Nothing Then
            ComboBox16.Items.AddRange(Critters_FRM)
            ComboBox17.Items.AddRange(Critters_FRM)
        End If
        ComboBox18.Items.AddRange(Perk_NAME)
        ComboBox22.Items.AddRange(Perk_NAME)
        ComboBox23.Items.AddRange(CaliberNAME)

        For n = 0 To UBound(AmmoPID)
            ComboBox24.Items.Add(AmmoNAME(n))
        Next

        ComboBox25.Items.AddRange(CaliberNAME)

        'тип предмета
        Select Case Items_LST(iLST_Index).itemType
            Case ItemType.Weapon
                TabControl1.TabPages.RemoveAt(3)
                TabControl1.TabPages.RemoveAt(2)
                TabControl1.TabPages.RemoveAt(1)
            Case ItemType.Armor
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(3))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(2))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(0))
            Case ItemType.Drugs
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(3))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(1))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(0))
            Case ItemType.Key, ItemType.Container
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(3))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(2))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(1))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(0))
            Case Else
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(2))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(1))
                TabControl1.TabPages.Remove(TabControl1.TabPages.Item(0))
        End Select

        ComboBox1.Items.AddRange(Items_FRM)
        ComboBox2.Items.AddRange(Iven_FRM)
        ComboBox9.Items.AddRange(Scripts_Lst)

        LoadProData()
        Me.Show()
    End Sub

    Private Sub LoadProData()
        Dim ProFile As String = PROTO_ITEMS & Items_LST(iLST_Index).proFile

        If cPath = Nothing Then cPath = DatFiles.CheckFile(ProFile)
        ProFile = cPath & ProFile
        ProFiles.LoadItemProData(ProFile, Items_LST(iLST_Index).itemType, CommonItem, WeaponItem, ArmorItem, AmmoItem, DrugItem, MiscItem, ContanerItem, KeyItem)
    End Sub

    Private Sub Items_Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        SetCommonValue_Form()
        Me.Text = TextBox29.Text & " [" & Items_LST(iLST_Index).proFile & "]"

        Select Case Items_LST(iLST_Index).itemType
            Case ItemType.Weapon
                SetWeaponValue_Form()
            Case ItemType.Armor
                SetArmorValue_Form()
            Case ItemType.Ammo
                SetAmmoValue_Form()
            Case ItemType.Container
                SetContanerValue_Form()
            Case ItemType.Drugs
                SetDrugsValue_Form()
            Case ItemType.Misc
                SetMiscValue_Form()
        End Select

        frmReady = True
    End Sub

    'Возвращает Имя или Описание предмета из msg файла
    Private Function GetNameItemMsg(ByVal NameID As Integer, Optional ByVal Desc As Boolean = False) As String
        If Desc Then NameID += 1
        Messages.GetMsgData("pro_item.msg")

        Return Messages.GetNameObject(NameID)
    End Function

    Private Sub SetCommonValue_Form()
        ComboBox7.SelectedIndex = CommonItem.ObjType
        TextBox29.Text = GetNameItemMsg(CommonItem.DescID)
        TextBox30.Text = GetNameItemMsg(CommonItem.DescID, True)
        TextBox33.Text = CommonItem.ProtoID
        TextBox33.Text = StrDup(8 - Len(TextBox33.Text), "0") & TextBox33.Text
        ComboBox1.SelectedIndex = CommonItem.FrmID
        If CommonItem.InvFID <> -1 Then ComboBox2.SelectedIndex = 1 + (CommonItem.InvFID - &H7000000) Else ComboBox2.SelectedIndex = 0
        ComboBox8.SelectedIndex = CommonItem.MaterialID
        If CommonItem.ScriptID <> -1 Then ComboBox9.SelectedIndex = 1 + (CommonItem.ScriptID - &H3000000) Else ComboBox9.SelectedIndex = 0
        For n = 0 To ComboBox3.Items.Count - 1
            If CChar(ComboBox3.Items(n)) = Chr(CommonItem.SoundID) Then ComboBox3.SelectedIndex = n
        Next
        '
        NumericUpDown1.Value = CommonItem.Cost
        NumericUpDown36.Value = CommonItem.LightDis
        NumericUpDown37.Value = Math.Round(CommonItem.LightInt * 100 / &HFFFF)
        NumericUpDown38.Value = CommonItem.Weight
        NumericUpDown39.Value = CommonItem.Size
        NumericUpDown64.Value = CommonItem.DescID
        'Flags
        CheckBox1.Checked = CommonItem.Falgs And &H8
        CheckBox2.Checked = CommonItem.Falgs And &H10
        CheckBox3.Checked = CommonItem.Falgs And &H800
        CheckBox4.Checked = CommonItem.Falgs And &H80000000
        CheckBox5.Checked = CommonItem.Falgs And &H20000000
        'CheckBox12.Checked = CommonItem.Falgs And &H20
        CheckBox24.Checked = CommonItem.Falgs And &H1000

        CheckBox6.Checked = CommonItem.Falgs And &H8000
        If Not CheckBox6.Checked Then
            RadioButton1.Checked = CommonItem.Falgs And &H10000
            RadioButton4.Checked = CommonItem.Falgs And &H20000
            RadioButton2.Checked = CommonItem.Falgs And &H40000
            RadioButton5.Checked = CommonItem.Falgs And &H80000
            RadioButton3.Checked = CommonItem.Falgs And &H4000
        End If

        CheckBox7.Checked = CommonItem.FalgsExt And &H800
        CheckBox8.Checked = CommonItem.FalgsExt And &H1000
        CheckBox9.Checked = CommonItem.FalgsExt And &H2000
        'CheckBox10.Checked = CommonItem.FalgsExt And &H4000
        CheckBox11.Checked = CommonItem.FalgsExt And &H8000
        CheckBox13.Checked = CommonItem.FalgsExt And &H800000
    End Sub

    Private Sub SetWeaponValue_Form()
        NumericUpDown2.Value = WeaponItem.MinDmg
        NumericUpDown3.Value = WeaponItem.MaxDmg
        NumericUpDown4.Value = WeaponItem.MaxRangeP
        NumericUpDown5.Value = WeaponItem.MaxRangeS
        NumericUpDown6.Value = WeaponItem.MinST
        NumericUpDown7.Value = WeaponItem.MPCostP
        NumericUpDown8.Value = WeaponItem.MPCostS
        NumericUpDown9.Value = WeaponItem.CritFail
        NumericUpDown10.Value = WeaponItem.Rounds
        NumericUpDown11.Value = WeaponItem.MaxAmmo

        ComboBox4.SelectedIndex = WeaponItem.AnimCode
        ComboBox5.SelectedIndex = WeaponItem.DmgType

        For n As SByte = 0 To ComboBox6.Items.Count - 1
            If CChar(ComboBox6.Items(n)) = Chr(WeaponItem.wSoundID) Then ComboBox6.SelectedIndex = n
        Next

        If WeaponItem.ProjPID <> -1 Then ComboBox10.SelectedIndex = WeaponItem.ProjPID - &H5000000 Else ComboBox10.SelectedIndex = 0
        If WeaponItem.Perk <> -1 Then ComboBox11.SelectedIndex = WeaponItem.Perk + 1 Else ComboBox11.SelectedIndex = 0

        ComboBox12.SelectedIndex = WeaponItem.Caliber

        If WeaponItem.AmmoPID <> -1 Then
            Dim aPid As Integer = WeaponItem.AmmoPID
            For n As SByte = 0 To UBound(AmmoPID)
                If aPid = AmmoPID(n) Then
                    ComboBox13.SelectedIndex = n + 1
                    Exit For
                End If
            Next
        Else
            ComboBox13.SelectedIndex = 0
        End If

        ComboBox14.SelectedIndex = CommonItem.FalgsExt And &HF
        ComboBox15.SelectedIndex = (CommonItem.FalgsExt >> 4) And &HF
        CheckBox21.Checked = CommonItem.FalgsExt And &H100
        CheckBox22.Checked = CommonItem.FalgsExt And &H200
    End Sub

    Private Sub SetArmorValue_Form()
        NumericUpDown12.Value = ArmorItem.AC

        NumericUpDown56.Value = ArmorItem.DTNormal
        NumericUpDown57.Value = ArmorItem.DTLaser
        NumericUpDown58.Value = ArmorItem.DTFire
        NumericUpDown59.Value = ArmorItem.DTPlasma
        NumericUpDown60.Value = ArmorItem.DTElectrical
        NumericUpDown61.Value = ArmorItem.DTEMP
        NumericUpDown62.Value = ArmorItem.DTExplode

        NumericUpDown71.Value = ArmorItem.DRNormal
        NumericUpDown70.Value = ArmorItem.DRLaser
        NumericUpDown69.Value = ArmorItem.DRFire
        NumericUpDown68.Value = ArmorItem.DRPlasma
        NumericUpDown67.Value = ArmorItem.DRElectrical
        NumericUpDown65.Value = ArmorItem.DREMP
        NumericUpDown63.Value = ArmorItem.DRExplode

        ComboBox16.SelectedIndex = ArmorItem.MaleFID - &H1000000
        ComboBox17.SelectedIndex = ArmorItem.FemaleFID - &H1000000

        If ArmorItem.Perk <> -1 Then ComboBox18.SelectedIndex = ArmorItem.Perk + 1 Else ComboBox18.SelectedIndex = 0
    End Sub

    Private Sub SetAmmoValue_Form()
        ComboBox23.SelectedIndex = AmmoItem.Caliber
        NumericUpDown26.Value = AmmoItem.Quantity
        NumericUpDown27.Value = AmmoItem.ACAdjust
        NumericUpDown28.Value = AmmoItem.DRAdjust
        NumericUpDown29.Value = AmmoItem.DamMult
        NumericUpDown30.Value = AmmoItem.DamDiv

        GroupBox24.Enabled = False
    End Sub

    Private Sub SetMiscValue_Form()
        If MiscItem.PowerPID <> -1 Then
            Dim Pid As Integer = MiscItem.PowerPID
            For n = 0 To UBound(AmmoPID)
                If Pid = AmmoPID(n) Then
                    ComboBox24.SelectedIndex = n + 1
                    Exit For
                End If
            Next
        Else
            ComboBox24.SelectedIndex = 0
        End If

        If (MiscItem.Charges >= 0) And (MiscItem.Charges <= 32000) Then
            NumericUpDown31.Value = MiscItem.Charges
        Else
            NumericUpDown31.Value = -1
        End If

        On Error Resume Next
        ComboBox25.SelectedIndex = MiscItem.PowerType
        GroupBox23.Enabled = False
    End Sub

    Private Sub SetContanerValue_Form()
        NumericUpDown32.Value = ContanerItem.MaxSize
        CheckBox15.Checked = ContanerItem.OpenFlags And &H1
        GroupBox25.Enabled = True
    End Sub

    Private Sub SetDrugsValue_Form()
        If DrugItem.Stat0 <> -1 Then
            ComboBox19.SelectedIndex = 2 + DrugItem.Stat0
        Else
            ComboBox19.SelectedIndex = 1
        End If

        If DrugItem.Stat1 <> -1 Then
            ComboBox20.SelectedIndex = 2 + DrugItem.Stat1
        Else
            ComboBox20.SelectedIndex = 1
        End If

        If DrugItem.Stat2 <> -1 Then
            ComboBox21.SelectedIndex = 2 + DrugItem.Stat2
        Else
            ComboBox21.SelectedIndex = 1
        End If

        NumericUpDown13.Value = DrugItem.iAmount0
        NumericUpDown14.Value = DrugItem.iAmount1
        NumericUpDown15.Value = DrugItem.iAmount2
        NumericUpDown16.Value = DrugItem.fAmount0
        NumericUpDown17.Value = DrugItem.fAmount1
        NumericUpDown18.Value = DrugItem.fAmount2
        NumericUpDown19.Value = DrugItem.sAmount0
        NumericUpDown20.Value = DrugItem.sAmount1
        NumericUpDown21.Value = DrugItem.sAmount2
        NumericUpDown22.Value = DrugItem.Duration1
        NumericUpDown23.Value = DrugItem.Duration2

        If DrugItem.W_Effect <> -1 Then ComboBox22.SelectedIndex = 1 + DrugItem.W_Effect Else ComboBox22.SelectedIndex = 0

        NumericUpDown24.Value = DrugItem.AddictionRate
        NumericUpDown25.Value = DrugItem.W_Onset
    End Sub

    Private Sub ComboBox1_Changed(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim frm As String = GetImageName(ComboBox1.SelectedItem, ".")

        If frm IsNot Nothing Then
            Dim pfile As String = Cache_Patch & ART_ITEMS & frm & ".gif"
            If Not File.Exists(pfile) Then ItemFrmGif("items\", frm)
            Dim img As Image = Image.FromFile(pfile)
            If img.Width > PictureBox1.Size.Width OrElse img.Size.Height > PictureBox1.Size.Height Then
                PictureBox1.BackgroundImageLayout = ImageLayout.Zoom
            Else
                PictureBox1.BackgroundImageLayout = ImageLayout.Center
            End If
            PictureBox1.BackgroundImage = img
            Exit Sub
        End If

        'BadFrm
        PictureBox1.BackgroundImage = My.Resources.RESERVAA
    End Sub

    Private Sub ComboBox2_Changed(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim frm As String = ComboBox2.SelectedItem

        If frm = "None" Then
            PictureBox4.BackgroundImage = Nothing
            Exit Sub
        End If

        frm = GetImageName(frm, ".")
        If frm IsNot Nothing Then
            Dim pfile As String = Cache_Patch & ART_INVEN & frm & ".gif"
            If Not File.Exists(pfile) Then ItemFrmGif("inven\", frm)
            Dim img As Image = Image.FromFile(pfile)
            If img.Width > PictureBox4.Size.Width Then
                PictureBox4.BackgroundImageLayout = ImageLayout.Zoom
            Else
                PictureBox4.BackgroundImageLayout = ImageLayout.Center
            End If
            PictureBox4.BackgroundImage = img
            Exit Sub
        End If

        'BadFrm
        PictureBox4.BackgroundImage = My.Resources.RESERVAA
    End Sub

    Private Sub GenderFID(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox17.SelectedIndexChanged, ComboBox16.SelectedIndexChanged
        Dim frm As String = GetImageName(CType(sender, ComboBox).SelectedItem, ",")

        If frm IsNot Nothing Then
            Dim pfile As String = Cache_Patch & ART_CRITTERS & frm & "aa.gif"
            If Not File.Exists(pfile) Then DatFiles.CritterFrmGif(frm)
            Dim img As Image = Image.FromFile(pfile)
            If CType(sender, ComboBox).Name.ToString = "ComboBox16" Then
                PictureBox2.Image = img
            Else
                PictureBox3.Image = img
            End If
            Exit Sub
        End If

        'BadFrm
        Dim resFile As Image = My.Resources.RESERVAA
        If CType(sender, ComboBox).Name.ToString = "ComboBox16" Then
            PictureBox2.Image = resFile
        Else
            PictureBox3.Image = resFile
        End If
    End Sub

    Private Function GetImageName(ByVal frm As String, ByVal symbol As String) As String
        Dim n As Integer = frm.IndexOf(symbol)

        If n <= 0 Then Return Nothing 'BadFrm
        frm = frm.Remove(n)
        If frmReady Then Button6.Enabled = True

        Return frm
    End Function

    'Save to Pro
    Private Sub Save_Pro(ByVal sender As Object, ByVal e As EventArgs) Handles Button6.Click
        'Common
        CommonItem.FrmID = ComboBox1.SelectedIndex
        If ComboBox2.SelectedIndex > 0 Then CommonItem.InvFID = (ComboBox2.SelectedIndex - 1) + &H7000000 Else CommonItem.InvFID = &HFFFFFFFF
        CommonItem.MaterialID = ComboBox8.SelectedIndex
        'CommonItem.ObjType = ComboBox7.SelectedIndex
        If ComboBox9.SelectedIndex > 0 Then CommonItem.ScriptID = (ComboBox9.SelectedIndex - 1) + &H3000000 Else CommonItem.ScriptID = &HFFFFFFFF
        CommonItem.SoundID = Asc(ComboBox3.Text)

        CommonItem.Cost = NumericUpDown1.Value
        CommonItem.LightDis = NumericUpDown36.Value
        CommonItem.LightInt = Math.Round((NumericUpDown37.Value * &HFFFF) / 100)
        CommonItem.Weight = NumericUpDown38.Value
        CommonItem.Size = NumericUpDown39.Value
        CommonItem.DescID = NumericUpDown64.Value
        'Flags
        Dim flags As Integer = CommonItem.Falgs
        If CheckBox1.Checked Then flags = flags Or &H8 Else flags = flags And (Not &H8)
        If CheckBox2.Checked Then flags = flags Or &H10 Else flags = flags And (Not &H10)
        If CheckBox3.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox4.Checked Then flags = flags Or &H80000000 Else flags = flags And (Not &H80000000)
        If CheckBox5.Checked Then flags = flags Or &H20000000 Else flags = flags And (Not &H20000000)
        If CheckBox24.Checked Then flags = flags Or &H1000 Else flags = flags And (Not &H1000)

        flags = flags And &HFFF03FFF
        If CheckBox6.Checked Then
            flags = flags Or &H8000
        Else
            If RadioButton1.Checked Then
                flags = flags Or &H10000
            ElseIf RadioButton4.Checked Then
                flags = flags Or &H20000
            ElseIf RadioButton2.Checked Then
                flags = flags Or &H40000
            ElseIf RadioButton5.Checked Then
                flags = flags Or &H80000
            ElseIf RadioButton3.Checked Then
                flags = flags Or &H4000
            End If
        End If
        CommonItem.Falgs = flags

        flags = CommonItem.FalgsExt
        If CheckBox7.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox8.Checked Then flags = flags Or &H1000 Else flags = flags And (Not &H1000)
        If CheckBox9.Checked Then flags = flags Or &H2000 Else flags = flags And (Not &H2000)
        'If CheckBox10.Checked Then flags = flags Or &H4000 Else flags = flags And (Not &H4000)
        If CheckBox11.Checked Then flags = flags Or &H8000 Else flags = flags And (Not &H8000)
        If CheckBox13.Checked Then flags = flags Or &H800000 Else flags = flags And Not (&H800000)
        CommonItem.FalgsExt = flags

        Select Case ComboBox7.SelectedIndex
            Case ItemType.Weapon
                Items_LST(iLST_Index).itemType = ItemType.Weapon
                WeaponItem.MinDmg = NumericUpDown2.Value
                WeaponItem.MaxDmg = NumericUpDown3.Value
                WeaponItem.MaxRangeP = NumericUpDown4.Value
                WeaponItem.MaxRangeS = NumericUpDown5.Value
                WeaponItem.MinST = NumericUpDown6.Value
                WeaponItem.MPCostP = NumericUpDown7.Value
                WeaponItem.MPCostS = NumericUpDown8.Value
                WeaponItem.CritFail = NumericUpDown9.Value
                WeaponItem.Rounds = NumericUpDown10.Value
                WeaponItem.MaxAmmo = NumericUpDown11.Value
                WeaponItem.AnimCode = ComboBox4.SelectedIndex
                WeaponItem.DmgType = ComboBox5.SelectedIndex
                WeaponItem.wSoundID = Asc(ComboBox6.Text)
                If ComboBox10.SelectedIndex > 0 Then WeaponItem.ProjPID = ComboBox10.SelectedIndex + &H5000000 Else WeaponItem.ProjPID = &HFFFFFFFF
                If ComboBox11.SelectedIndex > 0 Then WeaponItem.Perk = ComboBox11.SelectedIndex - 1 Else WeaponItem.Perk = &HFFFFFFFF
                WeaponItem.Caliber = ComboBox12.SelectedIndex
                If ComboBox13.SelectedIndex > 0 Then WeaponItem.AmmoPID = AmmoPID(ComboBox13.SelectedIndex - 1) Else WeaponItem.AmmoPID = &HFFFFFFFF
                flags = CommonItem.FalgsExt And &HFFFFFF00
                flags = flags Or ComboBox15.SelectedIndex << 4 Or ComboBox14.SelectedIndex
                If CheckBox21.Checked Then flags = flags Or &H100 Else flags = flags And (Not &H100)
                If CheckBox22.Checked Then flags = flags Or &H200 Else flags = flags And (Not &H200)
                CommonItem.FalgsExt = flags
                'Save
                SubSave_Pro(ItemType.Weapon)
            Case ItemType.Armor
                Items_LST(iLST_Index).itemType = ItemType.Armor
                ArmorItem.DTNormal = NumericUpDown56.Value
                ArmorItem.DTLaser = NumericUpDown57.Value
                ArmorItem.DTFire = NumericUpDown58.Value
                ArmorItem.DTPlasma = NumericUpDown59.Value
                ArmorItem.DTElectrical = NumericUpDown60.Value
                ArmorItem.DTEMP = NumericUpDown61.Value
                ArmorItem.DTExplode = NumericUpDown62.Value
                ArmorItem.DRNormal = NumericUpDown71.Value
                ArmorItem.DRLaser = NumericUpDown70.Value
                ArmorItem.DRFire = NumericUpDown69.Value
                ArmorItem.DRPlasma = NumericUpDown68.Value
                ArmorItem.DRElectrical = NumericUpDown67.Value
                ArmorItem.DREMP = NumericUpDown65.Value
                ArmorItem.DRExplode = NumericUpDown63.Value
                ArmorItem.AC = NumericUpDown12.Value
                ArmorItem.MaleFID = ComboBox16.SelectedIndex + &H1000000
                ArmorItem.FemaleFID = ComboBox17.SelectedIndex + &H1000000
                If ComboBox18.SelectedIndex > 0 Then ArmorItem.Perk = ComboBox18.SelectedIndex - 1 Else ArmorItem.Perk = &HFFFFFFFF
                'Save
                SubSave_Pro(ItemType.Armor)
            Case ItemType.Drugs
                Items_LST(iLST_Index).itemType = ItemType.Drugs
                If ComboBox19.SelectedIndex > 1 Then DrugItem.Stat0 = ComboBox19.SelectedIndex - 2 Else DrugItem.Stat0 = (&HFFFFFFFE + ComboBox19.SelectedIndex)
                If ComboBox20.SelectedIndex > 1 Then DrugItem.Stat1 = ComboBox20.SelectedIndex - 2 Else DrugItem.Stat1 = (&HFFFFFFFE + ComboBox20.SelectedIndex)
                If ComboBox21.SelectedIndex > 1 Then DrugItem.Stat2 = ComboBox21.SelectedIndex - 2 Else DrugItem.Stat2 = (&HFFFFFFFE + ComboBox21.SelectedIndex)
                DrugItem.iAmount0 = NumericUpDown13.Value
                DrugItem.iAmount1 = NumericUpDown14.Value
                DrugItem.iAmount2 = NumericUpDown15.Value
                DrugItem.fAmount0 = NumericUpDown16.Value
                DrugItem.fAmount1 = NumericUpDown17.Value
                DrugItem.fAmount2 = NumericUpDown18.Value
                DrugItem.sAmount0 = NumericUpDown19.Value
                DrugItem.sAmount1 = NumericUpDown20.Value
                DrugItem.sAmount2 = NumericUpDown21.Value
                DrugItem.Duration1 = NumericUpDown22.Value
                DrugItem.Duration2 = NumericUpDown23.Value
                If ComboBox22.SelectedIndex > 0 Then DrugItem.W_Effect = ComboBox22.SelectedIndex - 1 Else DrugItem.W_Effect = &HFFFFFFFF
                DrugItem.AddictionRate = NumericUpDown24.Value
                DrugItem.W_Onset = NumericUpDown25.Value
                'Save
                SubSave_Pro(ItemType.Drugs)
            Case ItemType.Ammo
                Items_LST(iLST_Index).itemType = ItemType.Ammo
                AmmoItem.Caliber = ComboBox23.SelectedIndex
                AmmoItem.Quantity = NumericUpDown26.Value
                AmmoItem.ACAdjust = NumericUpDown27.Value
                AmmoItem.DRAdjust = NumericUpDown28.Value
                AmmoItem.DamMult = NumericUpDown29.Value
                AmmoItem.DamDiv = NumericUpDown30.Value
                'Save
                SubSave_Pro(ItemType.Ammo)
            Case ItemType.Misc
                Items_LST(iLST_Index).itemType = ItemType.Misc
                If NumericUpDown31.Value <> -1 Then MiscItem.Charges = NumericUpDown31.Value
                If ComboBox24.SelectedIndex > 0 Then MiscItem.PowerPID = AmmoPID(ComboBox24.SelectedIndex - 1) Else MiscItem.PowerPID = &HFFFFFFFF
                If ComboBox25.SelectedIndex <> -1 Then MiscItem.PowerType = ComboBox25.SelectedIndex
                'Save
                SubSave_Pro(ItemType.Misc)
            Case ItemType.Container
                Items_LST(iLST_Index).itemType = ItemType.Container
                ContanerItem.MaxSize = NumericUpDown32.Value
                If CheckBox15.Checked Then ContanerItem.OpenFlags = &H1 Else ContanerItem.OpenFlags = 0
                'Save
                SubSave_Pro(ItemType.Container)
            Case Else 'Key
                Items_LST(iLST_Index).itemType = ItemType.Key
                'Save
                SubSave_Pro(ItemType.Key)
        End Select

        Dim indx As UShort = LW_SearhItemIndex(iLST_Index, Main_Form.ListView2)
        If indx <> Nothing Then
            If Main_Form.ListView2.Items(indx).SubItems(2).Text <> ComboBox7.Text Then
                Main_Form.ListView2.Items(indx).SubItems(2).Text = ComboBox7.Text
            End If
            Main_Form.ListView2.Items(indx).Text = "* " & Items_LST(Main_Form.ListView2.Items(indx).Tag).itemName
        End If

        cPath = DatFiles.CheckFile(PROTO_ITEMS & Items_LST(iLST_Index).proFile)
        Button6.Enabled = False
    End Sub

    Private Sub SubSave_Pro(ByVal iType As Integer)
        Dim proFile As String = SaveMOD_Path & PROTO_ITEMS & Items_LST(iLST_Index).proFile

        If Not Directory.Exists(SaveMOD_Path & PROTO_ITEMS) Then Directory.CreateDirectory(SaveMOD_Path & PROTO_ITEMS)
        ProFiles.SaveItemProData(proFile, iType, CommonItem, WeaponItem, ArmorItem, AmmoItem, DrugItem, MiscItem, ContanerItem, KeyItem)

        'Log
        Main_Form.TextBox1.Text = "Save Pro: " & proFile & vbCrLf & Main_Form.TextBox1.Text
    End Sub

    Private Sub SaveItemMsg(ByVal str As String, Optional ByRef Desc As Boolean = False)
        Dim ID As Integer = NumericUpDown64.Value 'DescID

        Messages.GetMsgData("pro_item.msg", False)
        If Messages.AddTextMSG(str, ID, Desc) Then
            MsgBox("You can not add value to the Msg file." & vbLf & "Not found msg line #:" & ID, MsgBoxStyle.Critical, "Error: pro_item.msg")
            Exit Sub
        End If
        'Save
        Messages.SaveMSGFile("pro_item.msg")

        'Update Name Item List
        If Not (Desc) Then
            Button4.Enabled = False
            Items_LST(iLST_Index).itemName = str
            Dim indx As Integer = LW_SearhItemIndex(iLST_Index, Main_Form.ListView2)
            If indx <> Nothing Then Main_Form.ListView2.Items(indx).Text = "? " & str
        Else
            Button5.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles CheckBox6.CheckedChanged
        If frmReady And CheckBox6.Focused Then
            frmReady = False
            RadioButton1.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            Button6.Enabled = True
            frmReady = True
        End If
    End Sub

    Private Sub SaveEnable(ByVal sender As Object, ByVal e As EventArgs) Handles NumericUpDown2.ValueChanged, NumericUpDown9.ValueChanged, _
        NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, _
        NumericUpDown3.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown71.ValueChanged, NumericUpDown70.ValueChanged, _
        NumericUpDown69.ValueChanged, NumericUpDown68.ValueChanged, NumericUpDown67.ValueChanged, NumericUpDown65.ValueChanged, NumericUpDown63.ValueChanged, _
        NumericUpDown62.ValueChanged, NumericUpDown61.ValueChanged, NumericUpDown60.ValueChanged, NumericUpDown59.ValueChanged, NumericUpDown58.ValueChanged, _
        NumericUpDown57.ValueChanged, NumericUpDown56.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown64.ValueChanged, NumericUpDown39.ValueChanged, _
        NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, _
        NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown25.ValueChanged, _
        NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, _
        NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, _
        NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown1.ValueChanged
        '
        If frmReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable2(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, _
        ComboBox4.SelectedIndexChanged, ComboBox15.SelectedIndexChanged, ComboBox14.SelectedIndexChanged, ComboBox13.SelectedIndexChanged, ComboBox12.SelectedIndexChanged, _
        ComboBox11.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, ComboBox18.SelectedIndexChanged, ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, _
        ComboBox3.SelectedIndexChanged, ComboBox25.SelectedIndexChanged, ComboBox24.SelectedIndexChanged, ComboBox23.SelectedIndexChanged, ComboBox22.SelectedIndexChanged, _
        ComboBox21.SelectedIndexChanged, ComboBox20.SelectedIndexChanged, ComboBox19.SelectedIndexChanged
        '
        If frmReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable3(ByVal sender As Object, ByVal e As EventArgs) Handles CheckBox21.CheckedChanged, CheckBox22.CheckedChanged, CheckBox9.CheckedChanged, _
        CheckBox8.CheckedChanged, CheckBox7.CheckedChanged, CheckBox5.CheckedChanged, CheckBox4.CheckedChanged, CheckBox3.CheckedChanged, CheckBox24.CheckedChanged, CheckBox2.CheckedChanged, _
        CheckBox15.CheckedChanged, CheckBox13.CheckedChanged, CheckBox11.CheckedChanged, CheckBox1.CheckedChanged
        '
        If frmReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable4(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton2.CheckedChanged, RadioButton5.CheckedChanged, RadioButton4.CheckedChanged, RadioButton3.CheckedChanged, RadioButton1.CheckedChanged
        If frmReady Then
            Button6.Enabled = True
            CheckBox6.Checked = False
        End If
    End Sub

    Private Sub Restore(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim tPatch As String = cPath

        If DatFiles.UnDatFile(PROTO_ITEMS & Items_LST(iLST_Index).proFile, ItemTypesProLen(Items_LST(iLST_Index).itemType)) Then
            Button6.Enabled = True
            frmReady = False
            ReloadPro = True
            cPath = Cache_Patch
            LoadProData()
            Items_Form_Load(Nothing, Nothing)
            ReloadPro = False
        Else
            MsgBox("This prototype file does not exist in Master.dat.", MsgBoxStyle.Exclamation)
        End If

        cPath = tPatch
    End Sub

    Private Sub Reload(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        LoadProData()
        frmReady = False
        ReloadPro = True
        Items_Form_Load(Nothing, Nothing)
        ReloadPro = False
        Button2.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ReloadPro OrElse (frmReady AndAlso MsgBox("Do you want to change the subtype of this object?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
            TabControl1.Visible = False
            If TabControl1.TabCount > 1 Then
                TabControl1.SelectTab(1)
                TabControl1.TabPages.RemoveAt(0)
            End If

            Select Case ComboBox7.SelectedIndex
                Case ItemType.Armor
                    CommonItem.ObjType = ItemType.Armor
                    If frmReady Then
                        ArmorItem.FemaleFID = &H1000000
                        ArmorItem.MaleFID = &H1000000
                    End If
                    TabControl1.TabPages.Insert(0, TabPage2)
                Case ItemType.Drugs
                    CommonItem.ObjType = ItemType.Drugs
                    TabControl1.TabPages.Insert(0, TabPage4)
                Case ItemType.Weapon
                    CommonItem.ObjType = ItemType.Weapon
                    If frmReady Then
                        WeaponItem.ProjPID = &H5000000
                        WeaponItem.AmmoPID = -1
                    End If
                    TabControl1.TabPages.Insert(0, TabPage1)
                    ComboBox6.SelectedIndex = 0
                Case ItemType.Ammo
                    CommonItem.ObjType = ItemType.Ammo
                    If frmReady Then
                        AmmoItem.DamMult = 1
                        AmmoItem.DamDiv = 1
                    End If
                    TabControl1.TabPages.Insert(0, TabPage5)
                    GroupBox23.Enabled = True
                    GroupBox24.Enabled = False
                Case ItemType.Misc
                    If frmReady Then
                        CommonItem.ObjType = ItemType.Misc
                        MiscItem.PowerPID = -1
                    End If
                    TabControl1.TabPages.Insert(0, TabPage5)
                    GroupBox23.Enabled = False
                    GroupBox24.Enabled = True
            End Select

            If frmReady Then
                CommonItem.FalgsExt = CommonItem.FalgsExt And &HCFFFFF
                frmReady = False
                Select Case ComboBox7.SelectedIndex
                    Case ItemType.Weapon
                        SetWeaponValue_Form()
                    Case ItemType.Armor
                        SetArmorValue_Form()
                    Case ItemType.Ammo
                        SetAmmoValue_Form()
                    Case ItemType.Container
                        SetContanerValue_Form()
                    Case ItemType.Drugs
                        SetDrugsValue_Form()
                    Case ItemType.Misc
                        SetMiscValue_Form()
                End Select
                frmReady = True
            End If

            TabControl1.SelectTab(0)
            Button6.Enabled = True
            TabControl1.Visible = True
        ElseIf frmReady Then
            frmReady = False
            ComboBox7.SelectedIndex = CommonItem.ObjType
            frmReady = True
        End If
    End Sub

    Private Sub TextBox29_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox29.TextChanged
        If frmReady Then Button4.Enabled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        SaveItemMsg(TextBox29.Text)
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox30.TextChanged
        If frmReady Then Button5.Enabled = True
    End Sub

    Private Sub Button5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button5.Click
        SaveItemMsg(TextBox30.Text, True)
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Items_Form_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Button6.Enabled Then
            Dim btn As MsgBoxResult = MsgBox("Save changes to Pro file?", MsgBoxStyle.YesNoCancel, "Attention!")
            If btn = MsgBoxResult.Yes Then
                Save_Pro(sender, e)
            ElseIf btn = MsgBoxResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Items_Form_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        Main_Form.Focus()
        Me.Dispose()
    End Sub

    Private Sub Items_Form_Activated(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Activated
        If frmReady Then Main_Form.ToolStripStatusLabel1.Text = cPath & PROTO_ITEMS & Items_LST(iLST_Index).proFile
    End Sub

    Private Sub Button6_EnabledChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Button6.EnabledChanged
        If Button6.Enabled Then Button2.Enabled = True
    End Sub

    'Private Sub Items_Form_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
    'For Each upDown As NumericUpDown In Me.Controls.OfType(Of NumericUpDown)()
    'AddHandler upDown.ValueChanged, AddressOf SaveEnable
    'Next
    'End Sub
End Class