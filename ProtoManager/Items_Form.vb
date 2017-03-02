Imports System.Drawing

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
    Private f_LST_Index As UShort         'индекс предмета в lst файле
    Private fReady As Boolean
    Private ReloadPro As Boolean
    Private cPath As String = Nothing

    Friend Sub Ini_ItemsForm(ByRef LST_Index As UShort)
        If cPath = Nothing Then cPath = Check_File("proto\items\" & Items_LST(LST_Index, 0))
        Dim fFile As Byte = FreeFile()
        FileOpen(fFile, cPath & "\proto\items\" & Items_LST(LST_Index, 0), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        FileGet(fFile, CommonItem)
        Select Case Items_LST(LST_Index, 1)
            Case "Weapon"
                FileGet(fFile, WeaponItem)
            Case "Armor"
                FileGet(fFile, ArmorItem)
            Case "Ammo"
                FileGet(fFile, AmmoItem)
            Case "Container"
                FileGet(fFile, ContanerItem)
            Case "Drugs"
                FileGet(fFile, DrugItem)
            Case "Misc"
                FileGet(fFile, MiscItem)
            Case "Key"
                FileGet(fFile, KeyItem)
        End Select
        FileClose(fFile)
        If Not Me.Visible Then f_LST_Index = LST_Index
    End Sub

    Private Sub Items_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetCommonValue_Form()
        Me.Text = TextBox29.Text & " [" & Items_LST(f_LST_Index, 0) & "]"
        Select Case Items_LST(f_LST_Index, 1)
            Case "Weapon"
                SetWeaponValue_Form()
            Case "Armor"
                SetArmorValue_Form()
            Case "Ammo"
                SetAmmoValue_Form()
            Case "Container"
                SetContanerValue_Form()
            Case "Drugs"
                SetDrugsValue_Form()
            Case "Misc"
                SetMiscValue_Form()
        End Select
        fReady = True
    End Sub

    'Возвращает Имя или Описание предмета из msg файла
    Private Function GetNameItemMsg(ByRef NameID As Integer, Optional ByRef Desc As Boolean = False) As String
        If Desc Then NameID += 1
        GetMsgData("pro_item.msg")
        'Dim msg_Path As String = Check_File("Text\English\Game\pro_item.msg")
        'If txtWin Then MSG_DATATEXT = IO.File.ReadAllLines(msg_Path & "\Text\English\Game\pro_item.msg", System.Text.Encoding.Default) _
        'Else MSG_DATATEXT = IO.File.ReadAllLines(msg_Path & "\Text\English\Game\pro_item.msg", System.Text.Encoding.GetEncoding("cp866"))
        'If txtLvCp Then EncodingLevCorp()
        Return GetNameCritter(NameID)
    End Function

    Private Sub SetCommonValue_Form()
        Dim n As SByte
        TextBox29.Text = GetNameItemMsg(ReverseBytes(CommonItem.DescID))
        TextBox30.Text = GetNameItemMsg(ReverseBytes(CommonItem.DescID), True)
        TextBox33.Text = ReverseBytes(CommonItem.ProtoID)
        TextBox33.Text = StrDup(8 - Len(TextBox33.Text), "0") & TextBox33.Text
        ComboBox1.SelectedIndex = ReverseBytes(CommonItem.FrmID)
        If CommonItem.InvFID <> -1 Then ComboBox2.SelectedIndex = 1 + (ReverseBytes(CommonItem.InvFID) - &H7000000) Else ComboBox2.SelectedIndex = 0
        ComboBox8.SelectedIndex = ReverseBytes(CommonItem.MaterialID)
        ComboBox7.SelectedIndex = ReverseBytes(CommonItem.ObjType)
        If CommonItem.ScriptID <> -1 Then ComboBox9.SelectedIndex = 1 + (ReverseBytes(CommonItem.ScriptID) - &H3000000) Else ComboBox9.SelectedIndex = 0
        For n = 0 To ComboBox3.Items.Count - 1
            If ComboBox3.Items(n) = Chr(CommonItem.SoundID) Then ComboBox3.SelectedIndex = n
        Next
        '
        NumericUpDown1.Value = ReverseBytes(CommonItem.Cost)
        NumericUpDown36.Value = ReverseBytes(CommonItem.LightDis)
        NumericUpDown37.Value = Math.Round(ReverseBytes(CommonItem.LightInt) * 100 / &HFFFF)
        NumericUpDown38.Value = ReverseBytes(CommonItem.Weight)
        NumericUpDown39.Value = ReverseBytes(CommonItem.Size)
        NumericUpDown64.Value = ReverseBytes(CommonItem.DescID)
        'Flags
        CheckBox1.Checked = ReverseBytes(CommonItem.Falgs) And &H8
        CheckBox2.Checked = ReverseBytes(CommonItem.Falgs) And &H10
        CheckBox3.Checked = ReverseBytes(CommonItem.Falgs) And &H800
        CheckBox4.Checked = ReverseBytes(CommonItem.Falgs) And &H80000000
        CheckBox5.Checked = ReverseBytes(CommonItem.Falgs) And &H20000000
        'CheckBox12.Checked = ReverseBytes(CommonItem.Falgs) And &H20
        CheckBox24.Checked = ReverseBytes(CommonItem.Falgs) And &H1000
        '
        CheckBox6.Checked = ReverseBytes(CommonItem.Falgs) And &H8000
        If Not CheckBox6.Checked Then
            RadioButton1.Checked = ReverseBytes(CommonItem.Falgs) And &H10000
            RadioButton4.Checked = ReverseBytes(CommonItem.Falgs) And &H20000
            RadioButton2.Checked = ReverseBytes(CommonItem.Falgs) And &H40000
            RadioButton5.Checked = ReverseBytes(CommonItem.Falgs) And &H80000
            RadioButton3.Checked = ReverseBytes(CommonItem.Falgs) And &H4000
        End If
        '
        CheckBox7.Checked = ReverseBytes(CommonItem.FalgsExt) And &H800
        CheckBox8.Checked = ReverseBytes(CommonItem.FalgsExt) And &H1000
        CheckBox9.Checked = ReverseBytes(CommonItem.FalgsExt) And &H2000
        'CheckBox10.Checked = ReverseBytes(CommonItem.FalgsExt) And &H4000
        CheckBox11.Checked = ReverseBytes(CommonItem.FalgsExt) And &H8000
        CheckBox13.Checked = ReverseBytes(CommonItem.FalgsExt) And &H800000
    End Sub

    Private Sub SetWeaponValue_Form()
        NumericUpDown2.Value = ReverseBytes(WeaponItem.MinDmg)
        NumericUpDown3.Value = ReverseBytes(WeaponItem.MaxDmg)
        NumericUpDown4.Value = ReverseBytes(WeaponItem.MaxRangeP)
        NumericUpDown5.Value = ReverseBytes(WeaponItem.MaxRangeS)
        NumericUpDown6.Value = ReverseBytes(WeaponItem.MinST)
        NumericUpDown7.Value = ReverseBytes(WeaponItem.MPCostP)
        NumericUpDown8.Value = ReverseBytes(WeaponItem.MPCostS)
        NumericUpDown9.Value = ReverseBytes(WeaponItem.CritFail)
        NumericUpDown10.Value = ReverseBytes(WeaponItem.Rounds)
        NumericUpDown11.Value = ReverseBytes(WeaponItem.MaxAmmo)
        '
        ComboBox4.SelectedIndex = ReverseBytes(WeaponItem.AnimCode)
        ComboBox5.SelectedIndex = ReverseBytes(WeaponItem.DmgType)
        For n As SByte = 0 To ComboBox6.Items.Count - 1
            If ComboBox6.Items(n) = Chr(WeaponItem.wSoundID) Then ComboBox6.SelectedIndex = n
        Next
        If WeaponItem.ProjPID <> -1 Then ComboBox10.SelectedIndex = ReverseBytes(WeaponItem.ProjPID) - &H5000000 Else ComboBox10.SelectedIndex = 0
        If WeaponItem.Perk <> -1 Then ComboBox11.SelectedIndex = ReverseBytes(WeaponItem.Perk) + 1 Else ComboBox11.SelectedIndex = 0
        ComboBox12.SelectedIndex = ReverseBytes(WeaponItem.Caliber)
        If WeaponItem.AmmoPID <> -1 Then
            Dim aPid As Integer = ReverseBytes(WeaponItem.AmmoPID)
            For n As SByte = 0 To UBound(AmmoPID)
                If aPid = AmmoPID(n) Then ComboBox13.SelectedIndex = n + 1 : Exit For
            Next
        Else : ComboBox13.SelectedIndex = 0 : End If
        '
        ComboBox14.SelectedIndex = ReverseBytes(CommonItem.FalgsExt) And &HF
        ComboBox15.SelectedIndex = (ReverseBytes(CommonItem.FalgsExt) >> 4) And &HF
        CheckBox21.Checked = ReverseBytes(CommonItem.FalgsExt) And &H100
        CheckBox22.Checked = ReverseBytes(CommonItem.FalgsExt) And &H200
    End Sub

    Private Sub SetArmorValue_Form()
        NumericUpDown12.Value = ReverseBytes(ArmorItem.AC)
        '
        NumericUpDown56.Value = ReverseBytes(ArmorItem.DTNormal)
        NumericUpDown57.Value = ReverseBytes(ArmorItem.DTLaser)
        NumericUpDown58.Value = ReverseBytes(ArmorItem.DTFire)
        NumericUpDown59.Value = ReverseBytes(ArmorItem.DTPlasma)
        NumericUpDown60.Value = ReverseBytes(ArmorItem.DTElectrical)
        NumericUpDown61.Value = ReverseBytes(ArmorItem.DTEMP)
        NumericUpDown62.Value = ReverseBytes(ArmorItem.DTExplode)
        '
        NumericUpDown71.Value = ReverseBytes(ArmorItem.DRNormal)
        NumericUpDown70.Value = ReverseBytes(ArmorItem.DRLaser)
        NumericUpDown69.Value = ReverseBytes(ArmorItem.DRFire)
        NumericUpDown68.Value = ReverseBytes(ArmorItem.DRPlasma)
        NumericUpDown67.Value = ReverseBytes(ArmorItem.DRElectrical)
        NumericUpDown65.Value = ReverseBytes(ArmorItem.DREMP)
        NumericUpDown63.Value = ReverseBytes(ArmorItem.DRExplode)
        '
        ComboBox16.SelectedIndex = ReverseBytes(ArmorItem.MaleFID) - &H1000000
        ComboBox17.SelectedIndex = ReverseBytes(ArmorItem.FemaleFID) - &H1000000
        If ArmorItem.Perk <> -1 Then ComboBox18.SelectedIndex = ReverseBytes(ArmorItem.Perk) + 1 Else ComboBox18.SelectedIndex = 0
    End Sub

    Private Sub SetAmmoValue_Form()
        ComboBox23.SelectedIndex = ReverseBytes(AmmoItem.Caliber)
        NumericUpDown26.Value = ReverseBytes(AmmoItem.Quantity)
        NumericUpDown27.Value = ReverseBytes(AmmoItem.ACAdjust)
        NumericUpDown28.Value = ReverseBytes(AmmoItem.DRAdjust)
        NumericUpDown29.Value = ReverseBytes(AmmoItem.DamMult)
        NumericUpDown30.Value = ReverseBytes(AmmoItem.DamDiv)
        GroupBox24.Enabled = False
    End Sub

    Private Sub SetMiscValue_Form()
        If MiscItem.PowerPID <> -1 Then
            Dim Pid As Integer = ReverseBytes(MiscItem.PowerPID)
            For n As SByte = 0 To UBound(AmmoPID)
                If Pid = AmmoPID(n) Then ComboBox24.SelectedIndex = n + 1 : Exit For
            Next
        Else : ComboBox24.SelectedIndex = 0 : End If
        If ReverseBytes(MiscItem.Charges) >= 0 And ReverseBytes(MiscItem.Charges) <= 32000 Then NumericUpDown31.Value = ReverseBytes(MiscItem.Charges) Else NumericUpDown31.Value = -1
        On Error Resume Next
        ComboBox25.SelectedIndex = ReverseBytes(MiscItem.PowerType)
        GroupBox23.Enabled = False
    End Sub

    Private Sub SetContanerValue_Form()
        NumericUpDown32.Value = ReverseBytes(ContanerItem.MaxSize)
        CheckBox15.Checked = ReverseBytes(ContanerItem.OpenFlags) And &H1
        GroupBox25.Enabled = True
    End Sub

    Private Sub SetDrugsValue_Form()
        If DrugItem.Stat0 <> -1 Then ComboBox19.SelectedIndex = 2 + ReverseBytes(DrugItem.Stat0) Else ComboBox19.SelectedIndex = 1
        If DrugItem.Stat1 <> -1 Then ComboBox20.SelectedIndex = 2 + ReverseBytes(DrugItem.Stat1) Else ComboBox20.SelectedIndex = 1
        If DrugItem.Stat2 <> -1 Then ComboBox21.SelectedIndex = 2 + ReverseBytes(DrugItem.Stat2) Else ComboBox21.SelectedIndex = 1
        NumericUpDown13.Value = ReverseBytes(DrugItem.iAmount0)
        NumericUpDown14.Value = ReverseBytes(DrugItem.iAmount1)
        NumericUpDown15.Value = ReverseBytes(DrugItem.iAmount2)
        NumericUpDown16.Value = ReverseBytes(DrugItem.fAmount0)
        NumericUpDown17.Value = ReverseBytes(DrugItem.fAmount1)
        NumericUpDown18.Value = ReverseBytes(DrugItem.fAmount2)
        NumericUpDown19.Value = ReverseBytes(DrugItem.sAmount0)
        NumericUpDown20.Value = ReverseBytes(DrugItem.sAmount1)
        NumericUpDown21.Value = ReverseBytes(DrugItem.sAmount2)
        NumericUpDown22.Value = ReverseBytes(DrugItem.Duration1)
        NumericUpDown23.Value = ReverseBytes(DrugItem.Duration2)
        '
        If DrugItem.W_Effect <> -1 Then ComboBox22.SelectedIndex = 1 + ReverseBytes(DrugItem.W_Effect) Else ComboBox22.SelectedIndex = 0
        NumericUpDown24.Value = ReverseBytes(DrugItem.AddictionRate)
        NumericUpDown25.Value = ReverseBytes(DrugItem.W_Onset)
    End Sub

    Private Sub ComboBox1_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim s As String = ComboBox1.SelectedItem
        Dim n As Byte = InStr(s, ".", CompareMethod.Text)
        If n = 0 Then GoTo BadFrm
        s = Strings.Left(s, n - 1)
        If fReady Then Button6.Enabled = True
        On Error GoTo BadFrm
        If My.Computer.FileSystem.FileExists(Cache_Patch & "\art\items\" & s & ".gif") = False Then
            ItemFrmGif("items\", s)
        End If
        Dim img As Image = Image.FromFile(Cache_Patch & "\art\items\" & s & ".gif")
        If img.Width > PictureBox1.Size.Width Or img.Size.Height > PictureBox1.Size.Height Then PictureBox1.BackgroundImageLayout = ImageLayout.Zoom Else PictureBox1.BackgroundImageLayout = ImageLayout.Center
        PictureBox1.BackgroundImage = img
        Exit Sub
BadFrm:
        PictureBox1.BackgroundImage = My.Resources.RESERVAA
    End Sub

    Private Sub ComboBox2_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim s As String = ComboBox2.SelectedItem
        If s = "None" Then PictureBox4.BackgroundImage = Nothing : Exit Sub
        Dim n As Byte = InStr(s, ".", CompareMethod.Text)
        If n = 0 Then GoTo BadFrm
        s = Strings.Left(s, n - 1)
        If fReady Then Button6.Enabled = True
        On Error GoTo BadFrm
        If My.Computer.FileSystem.FileExists(Cache_Patch & "\art\inven\" & s & ".gif") = False Then
            ItemFrmGif("inven\", s)
        End If
        Dim img As Image = Image.FromFile(Cache_Patch & "\art\inven\" & s & ".gif")
        If img.Width > PictureBox4.Size.Width Then PictureBox4.BackgroundImageLayout = ImageLayout.Zoom Else PictureBox4.BackgroundImageLayout = ImageLayout.Center
        PictureBox4.BackgroundImage = img
        Exit Sub
BadFrm:
        PictureBox4.BackgroundImage = My.Resources.RESERVAA
    End Sub

    Private Sub GenderFID(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox17.SelectedIndexChanged, ComboBox16.SelectedIndexChanged
        Dim s As String = sender.SelectedItem
        Dim n As Byte = InStr(s, ",", CompareMethod.Text)
        If n = 0 Then GoTo BadFrm
        s = Strings.Left(s, n - 1)
        If fReady Then Button6.Enabled = True
        On Error GoTo BadFrm
        If My.Computer.FileSystem.FileExists(Cache_Patch & "\art\critters\" & s & "aa.gif") = False Then
            Frm_to_Gif(s)
        End If
        If sender.Name = "ComboBox16" Then
            PictureBox2.Image = Image.FromFile(Cache_Patch & "\art\critters\" & s & "aa.gif")
        Else
            PictureBox3.Image = Image.FromFile(Cache_Patch & "\art\critters\" & s & "aa.gif")
        End If
        Exit Sub
BadFrm:
        If sender.Name = "ComboBox16" Then
            PictureBox2.Image = My.Resources.ResourceManager.GetObject(UCase(s) & "AA")
        Else
            PictureBox3.Image = My.Resources.ResourceManager.GetObject(UCase(s) & "AA")
        End If
    End Sub

    'Save to Pro
    Private Sub Save_Pro(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'Common
        CommonItem.FrmID = ReverseBytes(ComboBox1.SelectedIndex)
        If ComboBox2.SelectedIndex > 0 Then CommonItem.InvFID = ReverseBytes((ComboBox2.SelectedIndex - 1) + &H7000000) Else CommonItem.InvFID = &HFFFFFFFF
        CommonItem.MaterialID = ReverseBytes(ComboBox8.SelectedIndex)
        'CommonItem.ObjType = ReverseBytes(ComboBox7.SelectedIndex)
        If ComboBox9.SelectedIndex > 0 Then CommonItem.ScriptID = ReverseBytes((ComboBox9.SelectedIndex - 1) + &H3000000) Else CommonItem.ScriptID = &HFFFFFFFF
        CommonItem.SoundID = Asc(ComboBox3.Text)
        '
        CommonItem.Cost = ReverseBytes(NumericUpDown1.Value)
        CommonItem.LightDis = ReverseBytes(NumericUpDown36.Value)
        CommonItem.LightInt = ReverseBytes(Math.Round((NumericUpDown37.Value * &HFFFF) / 100))
        CommonItem.Weight = ReverseBytes(NumericUpDown38.Value)
        CommonItem.Size = ReverseBytes(NumericUpDown39.Value)
        CommonItem.DescID = ReverseBytes(NumericUpDown64.Value)
        'Flags
        Dim flags As Integer = ReverseBytes(CommonItem.Falgs)
        If CheckBox1.Checked Then flags = flags Or &H8 Else flags = flags And (Not &H8)
        If CheckBox2.Checked Then flags = flags Or &H10 Else flags = flags And (Not &H10)
        If CheckBox3.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox4.Checked Then flags = flags Or &H80000000 Else flags = flags And (Not &H80000000)
        If CheckBox5.Checked Then flags = flags Or &H20000000 Else flags = flags And (Not &H20000000)
        If CheckBox24.Checked Then flags = flags Or &H1000 Else flags = flags And (Not &H1000)
        '
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
        CommonItem.Falgs = ReverseBytes(flags)
        '
        flags = ReverseBytes(CommonItem.FalgsExt)
        If CheckBox7.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox8.Checked Then flags = flags Or &H1000 Else flags = flags And (Not &H1000)
        If CheckBox9.Checked Then flags = flags Or &H2000 Else flags = flags And (Not &H2000)
        'If CheckBox10.Checked Then flags = flags Or &H4000 Else flags = flags And (Not &H4000)
        If CheckBox11.Checked Then flags = flags Or &H8000 Else flags = flags And (Not &H8000)
        '
        If CheckBox13.Checked Then flags = flags Or &H800000 Else flags = flags And Not (&H800000)
        CommonItem.FalgsExt = ReverseBytes(flags)
        '
        Select Case ComboBox7.SelectedIndex
            Case 3 '"Weapon"
                Items_LST(f_LST_Index, 1) = "Weapon"
                WeaponItem.MinDmg = ReverseBytes(NumericUpDown2.Value)
                WeaponItem.MaxDmg = ReverseBytes(NumericUpDown3.Value)
                WeaponItem.MaxRangeP = ReverseBytes(NumericUpDown4.Value)
                WeaponItem.MaxRangeS = ReverseBytes(NumericUpDown5.Value)
                WeaponItem.MinST = ReverseBytes(NumericUpDown6.Value)
                WeaponItem.MPCostP = ReverseBytes(NumericUpDown7.Value)
                WeaponItem.MPCostS = ReverseBytes(NumericUpDown8.Value)
                WeaponItem.CritFail = ReverseBytes(NumericUpDown9.Value)
                WeaponItem.Rounds = ReverseBytes(NumericUpDown10.Value)
                WeaponItem.MaxAmmo = ReverseBytes(NumericUpDown11.Value)
                WeaponItem.AnimCode = ReverseBytes(ComboBox4.SelectedIndex)
                WeaponItem.DmgType = ReverseBytes(ComboBox5.SelectedIndex)
                WeaponItem.wSoundID = Asc(ComboBox6.Text)
                If ComboBox10.SelectedIndex > 0 Then WeaponItem.ProjPID = ReverseBytes(ComboBox10.SelectedIndex + &H5000000) Else WeaponItem.ProjPID = &HFFFFFFFF
                If ComboBox11.SelectedIndex > 0 Then WeaponItem.Perk = ReverseBytes(ComboBox11.SelectedIndex - 1) Else WeaponItem.Perk = &HFFFFFFFF
                WeaponItem.Caliber = ReverseBytes(ComboBox12.SelectedIndex)
                If ComboBox13.SelectedIndex > 0 Then WeaponItem.AmmoPID = ReverseBytes(AmmoPID(ComboBox13.SelectedIndex - 1)) Else WeaponItem.AmmoPID = &HFFFFFFFF
                flags = ReverseBytes(CommonItem.FalgsExt) And &HFFFFFF00
                flags = flags Or ComboBox15.SelectedIndex << 4 Or ComboBox14.SelectedIndex
                If CheckBox21.Checked Then flags = flags Or &H100 Else flags = flags And (Not &H100)
                If CheckBox22.Checked Then flags = flags Or &H200 Else flags = flags And (Not &H200)
                CommonItem.FalgsExt = ReverseBytes(flags)
                'Save
                SubSave_Pro(0)
            Case 0 '"Armor"
                Items_LST(f_LST_Index, 1) = "Armor"
                ArmorItem.DTNormal = ReverseBytes(NumericUpDown56.Value)
                ArmorItem.DTLaser = ReverseBytes(NumericUpDown57.Value)
                ArmorItem.DTFire = ReverseBytes(NumericUpDown58.Value)
                ArmorItem.DTPlasma = ReverseBytes(NumericUpDown59.Value)
                ArmorItem.DTElectrical = ReverseBytes(NumericUpDown60.Value)
                ArmorItem.DTEMP = ReverseBytes(NumericUpDown61.Value)
                ArmorItem.DTExplode = ReverseBytes(NumericUpDown62.Value)
                ArmorItem.DRNormal = ReverseBytes(NumericUpDown71.Value)
                ArmorItem.DRLaser = ReverseBytes(NumericUpDown70.Value)
                ArmorItem.DRFire = ReverseBytes(NumericUpDown69.Value)
                ArmorItem.DRPlasma = ReverseBytes(NumericUpDown68.Value)
                ArmorItem.DRElectrical = ReverseBytes(NumericUpDown67.Value)
                ArmorItem.DREMP = ReverseBytes(NumericUpDown65.Value)
                ArmorItem.DRExplode = ReverseBytes(NumericUpDown63.Value)
                ArmorItem.AC = ReverseBytes(NumericUpDown12.Value)
                ArmorItem.MaleFID = ReverseBytes(ComboBox16.SelectedIndex + &H1000000)
                ArmorItem.FemaleFID = ReverseBytes(ComboBox17.SelectedIndex + &H1000000)
                If ComboBox18.SelectedIndex > 0 Then ArmorItem.Perk = ReverseBytes(ComboBox18.SelectedIndex - 1) Else ArmorItem.Perk = &HFFFFFFFF
                'Save
                SubSave_Pro(1)
            Case 2 '"Drugs"
                Items_LST(f_LST_Index, 1) = "Drugs"
                If ComboBox19.SelectedIndex > 1 Then DrugItem.Stat0 = ReverseBytes(ComboBox19.SelectedIndex - 2) Else DrugItem.Stat0 = ReverseBytes(&HFFFFFFFE + ComboBox19.SelectedIndex)
                If ComboBox20.SelectedIndex > 1 Then DrugItem.Stat1 = ReverseBytes(ComboBox20.SelectedIndex - 2) Else DrugItem.Stat1 = ReverseBytes(&HFFFFFFFE + ComboBox20.SelectedIndex)
                If ComboBox21.SelectedIndex > 1 Then DrugItem.Stat2 = ReverseBytes(ComboBox21.SelectedIndex - 2) Else DrugItem.Stat2 = ReverseBytes(&HFFFFFFFE + ComboBox21.SelectedIndex)
                DrugItem.iAmount0 = ReverseBytes(NumericUpDown13.Value)
                DrugItem.iAmount1 = ReverseBytes(NumericUpDown14.Value)
                DrugItem.iAmount2 = ReverseBytes(NumericUpDown15.Value)
                DrugItem.fAmount0 = ReverseBytes(NumericUpDown16.Value)
                DrugItem.fAmount1 = ReverseBytes(NumericUpDown17.Value)
                DrugItem.fAmount2 = ReverseBytes(NumericUpDown18.Value)
                DrugItem.sAmount0 = ReverseBytes(NumericUpDown19.Value)
                DrugItem.sAmount1 = ReverseBytes(NumericUpDown20.Value)
                DrugItem.sAmount2 = ReverseBytes(NumericUpDown21.Value)
                DrugItem.Duration1 = ReverseBytes(NumericUpDown22.Value)
                DrugItem.Duration2 = ReverseBytes(NumericUpDown23.Value)
                If ComboBox22.SelectedIndex > 0 Then DrugItem.W_Effect = ReverseBytes(ComboBox22.SelectedIndex - 1) Else DrugItem.W_Effect = &HFFFFFFFF
                DrugItem.AddictionRate = ReverseBytes(NumericUpDown24.Value)
                DrugItem.W_Onset = ReverseBytes(NumericUpDown25.Value)
                'Save
                SubSave_Pro(2)
            Case 4 '"Ammo"
                Items_LST(f_LST_Index, 1) = "Ammo"
                AmmoItem.Caliber = ReverseBytes(ComboBox23.SelectedIndex)
                AmmoItem.Quantity = ReverseBytes(NumericUpDown26.Value)
                AmmoItem.ACAdjust = ReverseBytes(NumericUpDown27.Value)
                AmmoItem.DRAdjust = ReverseBytes(NumericUpDown28.Value)
                AmmoItem.DamMult = ReverseBytes(NumericUpDown29.Value)
                AmmoItem.DamDiv = ReverseBytes(NumericUpDown30.Value)
                'Save
                SubSave_Pro(3)
            Case 5 '"Misc"
                Items_LST(f_LST_Index, 1) = "Misc"
                If NumericUpDown31.Value <> -1 Then MiscItem.Charges = ReverseBytes(NumericUpDown31.Value)
                If ComboBox24.SelectedIndex > 0 Then MiscItem.PowerPID = ReverseBytes(AmmoPID(ComboBox24.SelectedIndex - 1)) Else MiscItem.PowerPID = &HFFFFFFFF
                If ComboBox25.SelectedIndex <> -1 Then MiscItem.PowerType = ReverseBytes(ComboBox25.SelectedIndex)
                'Save
                SubSave_Pro(4)
            Case 1 '"Container"
                Items_LST(f_LST_Index, 1) = "Container"
                ContanerItem.MaxSize = ReverseBytes(NumericUpDown32.Value)
                If CheckBox15.Checked Then ContanerItem.OpenFlags = ReverseBytes(&H1) Else ContanerItem.OpenFlags = 0
                'Save
                SubSave_Pro(5)
            Case Else 'Key
                Items_LST(f_LST_Index, 1) = "Key"
                'Save
                SubSave_Pro(6)
        End Select
        '
        Dim indx As UShort = LW_SearhItemIndex(f_LST_Index, Main_Form.ListView2)
        If indx <> Nothing Then
            If Main_Form.ListView2.Items(indx).SubItems(2).Text <> ComboBox7.Text Then Main_Form.ListView2.Items(indx).SubItems(2).Text = ComboBox7.Text
            Main_Form.ListView2.Items(indx).Text = "* " & Items_NAME(Main_Form.ListView2.Items(indx).Tag)
        End If
        Button6.Enabled = False
        cPath = Check_File("proto\items\" & Items_LST(f_LST_Index, 0))
    End Sub

    Private Sub SubSave_Pro(ByVal type As Byte)
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\items") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\items")
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0)) Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0), IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        IO.File.Delete(SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0))
        Dim fFile As Byte = FreeFile()
        FileOpen(fFile, SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0), OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, CommonItem)
        Select Case type
            Case 0 '"Weapon"
                FilePut(fFile, WeaponItem)
            Case 1 '"Armor"
                FilePut(fFile, ArmorItem)
            Case 2 '"Drugs"
                FilePut(fFile, DrugItem)
            Case 3 '"Ammo"
                FilePut(fFile, AmmoItem)
            Case 4 '"Misc"
                FilePut(fFile, MiscItem)
            Case 5 '"Container"
                FilePut(fFile, ContanerItem)
            Case Else 'Key
                FilePut(fFile, KeyItem)
        End Select
        FileClose(fFile)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0), IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        'Log
        Main_Form.TextBox1.Text = "Save: " & SaveMOD_Path & "\proto\items\" & Items_LST(f_LST_Index, 0) & vbCrLf & Main_Form.TextBox1.Text
    End Sub

    Private Sub SaveItemMsg(ByVal str As String, Optional ByRef Desc As Boolean = False)
        Current_Path = Check_File("Text\English\Game\pro_item.msg")
        If txtWin Then MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\pro_item.msg", System.Text.Encoding.Default) _
        Else MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\pro_item.msg", System.Text.Encoding.GetEncoding("cp866"))
        '
        If txtLvCp Then CodingToLevCorp(str)
        Dim ID As Integer = NumericUpDown64.Value 'DescID
        If Add_ProMSG(str, ID, Desc) Then GoTo MsgError
        'Save
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\Text\English\Game") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\Text\English\Game")
        If txtWin Then IO.File.WriteAllLines(SaveMOD_Path & "\Text\English\Game\pro_item.msg", MSG_DATATEXT, System.Text.Encoding.Default) _
        Else IO.File.WriteAllLines(SaveMOD_Path & "\Text\English\Game\pro_item.msg", MSG_DATATEXT, System.Text.Encoding.GetEncoding("cp866"))
        'Log
        Main_Form.TextBox1.Text = "Update: " & SaveMOD_Path & "\Text\English\Game\pro_item.msg" & vbCrLf & Main_Form.TextBox1.Text
        'Update Name Item List
        If Not (Desc) Then
            Items_NAME(f_LST_Index) = str
            Dim indx As UShort = LW_SearhItemIndex(f_LST_Index, Main_Form.ListView2)
            If indx <> Nothing Then Main_Form.ListView2.Items(indx).Text = "! " & str
        End If
        Exit Sub
MsgError:
        MsgBox("You can not add value in the Msg file." & vbCrLf & "Not found line: " & ID, MsgBoxStyle.SystemModal And MsgBoxStyle.Critical, "Error")
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If fReady And CheckBox6.Focused Then
            fReady = False
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            Button6.Enabled = True
            fReady = True
        End If
    End Sub

    Public Sub SaveEnable(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged, NumericUpDown9.ValueChanged, _
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
        If fReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, _
        ComboBox4.SelectedIndexChanged, ComboBox15.SelectedIndexChanged, ComboBox14.SelectedIndexChanged, ComboBox13.SelectedIndexChanged, ComboBox12.SelectedIndexChanged, _
        ComboBox11.SelectedIndexChanged, ComboBox10.SelectedIndexChanged, ComboBox18.SelectedIndexChanged, ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, _
        ComboBox3.SelectedIndexChanged, ComboBox25.SelectedIndexChanged, ComboBox24.SelectedIndexChanged, ComboBox23.SelectedIndexChanged, ComboBox22.SelectedIndexChanged, _
        ComboBox21.SelectedIndexChanged, ComboBox20.SelectedIndexChanged, ComboBox19.SelectedIndexChanged
        If fReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox21.CheckedChanged, CheckBox22.CheckedChanged, CheckBox9.CheckedChanged, _
        CheckBox8.CheckedChanged, CheckBox7.CheckedChanged, CheckBox5.CheckedChanged, CheckBox4.CheckedChanged, CheckBox3.CheckedChanged, CheckBox24.CheckedChanged, CheckBox2.CheckedChanged, _
        CheckBox15.CheckedChanged, CheckBox13.CheckedChanged, CheckBox11.CheckedChanged, CheckBox1.CheckedChanged
        If fReady Then Button6.Enabled = True
    End Sub

    Private Sub SaveEnable4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged, RadioButton5.CheckedChanged, RadioButton4.CheckedChanged, RadioButton3.CheckedChanged, RadioButton1.CheckedChanged
        If fReady Then Button6.Enabled = True : CheckBox6.Checked = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim tPatch As String = cPath
        If UnDat_File("proto\items\" & Items_LST(f_LST_Index, 0)) Then
            Button6.Enabled = True
            fReady = False
            ReloadPro = True
            cPath = Cache_Patch
            Ini_ItemsForm(f_LST_Index)
            Items_Form_Load(sender, e)
            ReloadPro = False
        Else
            MsgBox("Prototype this number does not exist in Master.dat.", MsgBoxStyle.Exclamation)
        End If
        cPath = tPatch
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Ini_ItemsForm(f_LST_Index)
        fReady = False
        ReloadPro = True
        Items_Form_Load(sender, e)
        ReloadPro = False
        Button2.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        If (fReady AndAlso MsgBox("Change the subtype of the object?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Or ReloadPro Then
            TabControl1.Visible = False
            If TabControl1.TabCount > 1 Then TabControl1.SelectTab(1) : TabControl1.TabPages.RemoveAt(0)
            Select Case ComboBox7.SelectedIndex
                Case 0 '"Armor"
                    TabControl1.TabPages.Insert(0, TabPage2)
                    CommonItem.ObjType = 0
                    ArmorItem.FemaleFID = ReverseBytes(&H1000000)
                    ArmorItem.MaleFID = ReverseBytes(&H1000000)
                Case 2 '"Drugs"
                    TabControl1.TabPages.Insert(0, TabPage4)
                    CommonItem.ObjType = ReverseBytes(2)
                Case 3 '"Weapon"
                    TabControl1.TabPages.Insert(0, TabPage1)
                    CommonItem.ObjType = ReverseBytes(3)
                    WeaponItem.ProjPID = ReverseBytes(&H5000000)
                    WeaponItem.AmmoPID = -1
                    ComboBox6.SelectedIndex = 0
                Case 4 '"Ammo"
                    TabControl1.TabPages.Insert(0, TabPage5)
                    GroupBox23.Enabled = True : GroupBox24.Enabled = False
                    CommonItem.ObjType = ReverseBytes(4)
                    AmmoItem.DamMult = ReverseBytes(1)
                    AmmoItem.DamDiv = ReverseBytes(1)
                Case 5 'Misc
                    TabControl1.TabPages.Insert(0, TabPage5)
                    GroupBox23.Enabled = False : GroupBox24.Enabled = True
                    CommonItem.ObjType = ReverseBytes(5)
                    MiscItem.PowerPID = -1
            End Select
            CommonItem.FalgsExt = CommonItem.FalgsExt And &HCFFFFF
            Dim Btmp As Boolean = fReady 'сохраняем 
            fReady = False
            Select Case ComboBox7.SelectedIndex
                Case 3 '"Weapon"
                    SetWeaponValue_Form()
                Case 0 '"Armor"
                    SetArmorValue_Form()
                Case 4 '"Ammo"
                    SetAmmoValue_Form()
                Case 1 '"Container"
                    SetContanerValue_Form()
                Case 2 '"Drugs"
                    SetDrugsValue_Form()
                Case 5 '"Misc"
                    SetMiscValue_Form()
            End Select
            fReady = Btmp 'возвращаем
            TabControl1.SelectTab(0)
            Button6.Enabled = True
            TabControl1.Visible = True
        ElseIf fReady Then
            fReady = False
            ComboBox7.SelectedIndex = ReverseBytes(CommonItem.ObjType)
            fReady = True
        End If
    End Sub

    Private Sub TextBox29_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox29.TextChanged
        If fReady Then Button4.Enabled = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SaveItemMsg(TextBox29.Text)
        Button4.Enabled = False
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        If fReady Then Button5.Enabled = True
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        SaveItemMsg(TextBox30.Text, True)
        Button5.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Items_Form_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If Button6.Enabled Then
            Dim btn As MsgBoxResult = MsgBox("Save changes to the Pro-File?", MsgBoxStyle.YesNoCancel, "Attention!")
            If btn = MsgBoxResult.Yes Then
                Save_Pro(sender, e)
            ElseIf btn = MsgBoxResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Items_Form_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Main_Form.Focus()
        Me.Dispose()
    End Sub

    Private Sub Items_Form_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If fReady Then Main_Form.ToolStripStatusLabel1.Text = cPath & "\proto\items\" & Items_LST(f_LST_Index, 0)
    End Sub

    'Private Sub Items_Form_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
    'For Each upDown As NumericUpDown In Me.Controls.OfType(Of NumericUpDown)()
    'AddHandler upDown.ValueChanged, AddressOf SaveEnable
    'Next
    'End Sub

    Private Sub Button6_EnabledChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.EnabledChanged
        If Button6.Enabled Then Button2.Enabled = True
    End Sub

End Class