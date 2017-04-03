'Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.Reflection
Imports Prototypes

Friend Class Critter_Form
    Private fCritterPro As CritPro
    Private TabView0, TabView1, TabView2 As Boolean
    Private fLW_Index As UShort
    Private cPath As String = Nothing

    Friend Sub Ini_CritterForm(ByRef Lw_Index As UShort)
        'Dim b() As Integer = My.Resources.ResourceManager.GetObject("critter")
        Dim CrttrProData(103) As Integer
        If cPath = Nothing Then cPath = Check_File("proto\critters\" & Critter_LST(Lw_Index))
        Dim fFile As Byte = FreeFile()
        FileOpen(fFile, cPath & "\proto\critters\" & Critter_LST(Lw_Index), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        On Error GoTo BadFormat
        FileGet(fFile, CrttrProData)
        On Error GoTo 0
        FileClose(fFile)
        For n = 0 To CrttrProData.Length - 1
            CrttrProData(n) = ReverseBytes(CrttrProData(n))
        Next
        fCritterPro = fnBytesToStruct(CrttrProData, fCritterPro.GetType)
        '
        If Not (Me.Visible) Then fLW_Index = Lw_Index
        '
        'ListView1.FocusedItem.Index(0)
        Exit Sub
BadFormat:
        FileClose()
        MsgBox("The file " & Critter_LST(Lw_Index) & " does not have the correct format: Size <> 416 Bytes", MsgBoxStyle.Critical)
        Main_Form.Focus()
        Me.Dispose()
    End Sub

    Private Sub Critter_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim x As SByte = 1
        For n = 0 To UBound(Items_LST)
            If Items_LST(n, 1) = "Armor" Then ListView1.Items.Add(Items_NAME(n)) : ListView1.Items(x).Tag = Items_LST(n, 0) : x += 1
        Next
        SetValueForm_Tab1()
        CalcSpecialParam()
    End Sub

    'Возвращает Имя или Описание криттера из msg файла
    Private Function GetNameCritterMsg(ByVal NameID As Integer, Optional ByRef Desc As Boolean = False) As String
        If Desc Then NameID += 1
        GetMsgData("pro_crit.msg")
        'Dim msg_Path As String = Check_File("Text\English\Game\pro_crit.msg")
        'If txtWin Then MSG_DATATEXT = IO.File.ReadAllLines(msg_Path & "\Text\English\Game\pro_crit.msg", System.Text.Encoding.Default) _
        'Else MSG_DATATEXT = IO.File.ReadAllLines(msg_Path & "\Text\English\Game\pro_crit.msg", System.Text.Encoding.GetEncoding("cp866"))
        'If txtLvCp Then EncodingLevCorp()
        Return GetNameCritter(NameID)
    End Function

    Private Sub TabControl1_Selected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected
        If e.TabPageIndex = 1 And Not (TabView1) Then
            SetValueForm_Tab2()
        ElseIf e.TabPageIndex = 2 And Not (TabView2) Then
            SetValueForm_Tab3()
        End If
    End Sub

    Private Sub SetValueForm_Tab1()
        TextBox29.Text = GetNameCritterMsg(fCritterPro.DescID)
        Me.Text = TextBox29.Text & " [" & Critter_LST(fLW_Index) & "]"
        TextBox33.Text = fCritterPro.ProtoID
        ComboBox1.SelectedIndex = fCritterPro.FrmID - &H1000000I
        'special
        NumericUpDown1.Value = fCritterPro.Strength
        NumericUpDown2.Value = fCritterPro.Perception
        NumericUpDown3.Value = fCritterPro.Endurance
        NumericUpDown4.Value = fCritterPro.Charisma
        NumericUpDown5.Value = fCritterPro.Intelligence
        NumericUpDown6.Value = fCritterPro.Agility
        NumericUpDown7.Value = fCritterPro.Luck
        'Skills
        NumericUpDown8.Value = fCritterPro.SmallGuns
        NumericUpDown9.Value = fCritterPro.BigGuns
        NumericUpDown10.Value = fCritterPro.EnergyGun
        NumericUpDown11.Value = fCritterPro.Melee
        NumericUpDown12.Value = fCritterPro.Unarmed
        NumericUpDown13.Value = fCritterPro.Throwing
        NumericUpDown14.Value = fCritterPro.FirstAid
        NumericUpDown15.Value = fCritterPro.Doctor
        NumericUpDown16.Value = fCritterPro.Outdoorsman
        NumericUpDown17.Value = fCritterPro.Sneak
        NumericUpDown18.Value = fCritterPro.Lockpick
        NumericUpDown19.Value = fCritterPro.Steal
        NumericUpDown20.Value = fCritterPro.Traps
        NumericUpDown21.Value = fCritterPro.Science
        NumericUpDown22.Value = fCritterPro.Repair
        NumericUpDown23.Value = fCritterPro.Speech
        NumericUpDown24.Value = fCritterPro.Barter
        NumericUpDown25.Value = fCritterPro.Gambling
        '
        NumericUpDown26.Value = fCritterPro.b_HP
        NumericUpDown27.Value = fCritterPro.b_AC
        NumericUpDown28.Value = fCritterPro.b_AP
        NumericUpDown29.Value = fCritterPro.b_Weight
        NumericUpDown30.Value = fCritterPro.b_MeleeDmg
        NumericUpDown31.Value = fCritterPro.b_Sequence
        NumericUpDown32.Value = fCritterPro.b_Healing
        NumericUpDown33.Value = fCritterPro.b_Critical
        NumericUpDown34.Value = fCritterPro.b_Better
        NumericUpDown35.Value = fCritterPro.b_UnarmedDmg
        '
        'TextBox19.Text = ReverseBytes(fCritterPro.HP)
        'TextBox20.Text = ReverseBytes(fCritterPro.AC)
        'TextBox21.Text = ReverseBytes(fCritterPro.AP)
        'TextBox22.Text = ReverseBytes(fCritterPro.Weight)
        'TextBox23.Text = ReverseBytes(fCritterPro.MeleeDmg)
        'TextBox24.Text = ReverseBytes(fCritterPro.Sequence)
        'TextBox25.Text = ReverseBytes(fCritterPro.Healing)
        'TextBox26.Text = ReverseBytes(fCritterPro.Critical)
        '
        TextBox27.Text = fCritterPro.Better
        TextBox28.Text = fCritterPro.UnarmedDmg
        '
        'TextBox31.Text = ReverseBytes(fCritterPro.DRPoison)
        'TextBox32.Text = ReverseBytes(fCritterPro.DRRadiation)
        NumericUpDown55.Value = fCritterPro.b_DRPoison
        NumericUpDown54.Value = fCritterPro.b_DRRadiation
        '
        TabView0 = True
    End Sub

    Private Sub CalcSpecialParam()
        TextBox1.Text = (5 + (4 * NumericUpDown6.Value)) + NumericUpDown8.Value
        TextBox2.Text = (2 * NumericUpDown6.Value) + NumericUpDown9.Value
        TextBox3.Text = (2 * NumericUpDown6.Value) + NumericUpDown10.Value
        TextBox5.Text = (30 + (NumericUpDown6.Value + NumericUpDown1.Value) * 2) + NumericUpDown12.Value
        TextBox4.Text = (20 + (NumericUpDown6.Value + NumericUpDown1.Value) * 2) + NumericUpDown11.Value
        TextBox6.Text = (4 * NumericUpDown6.Value) + NumericUpDown13.Value
        TextBox7.Text = ((NumericUpDown2.Value + NumericUpDown5.Value) * 2) + NumericUpDown14.Value
        TextBox8.Text = (5 + NumericUpDown2.Value + NumericUpDown5.Value) + NumericUpDown15.Value
        TextBox9.Text = ((NumericUpDown3.Value + NumericUpDown5.Value) * 2) + NumericUpDown16.Value
        TextBox10.Text = (5 + (3 * NumericUpDown6.Value)) + NumericUpDown17.Value
        TextBox11.Text = (10 + NumericUpDown2.Value + NumericUpDown6.Value) + NumericUpDown18.Value
        TextBox12.Text = (3 * NumericUpDown6.Value) + NumericUpDown19.Value
        TextBox13.Text = (10 + NumericUpDown2.Value + NumericUpDown6.Value) + NumericUpDown20.Value
        TextBox14.Text = (4 * NumericUpDown5.Value) + NumericUpDown21.Value
        TextBox15.Text = (3 * NumericUpDown5.Value) + NumericUpDown22.Value
        TextBox16.Text = (5 * NumericUpDown4.Value) + NumericUpDown23.Value
        TextBox17.Text = (4 * NumericUpDown4.Value) + NumericUpDown24.Value
        TextBox18.Text = (5 * NumericUpDown7.Value) + NumericUpDown25.Value
        '
        TextBox19.Text = (15 + NumericUpDown1.Value + (NumericUpDown3.Value * 2)) + NumericUpDown26.Value
        TextBox20.Text = NumericUpDown6.Value + NumericUpDown27.Value
        TextBox21.Text = Fix(5 + (NumericUpDown6.Value / 2)) + NumericUpDown28.Value
        TextBox22.Text = (25 + (25 * NumericUpDown1.Value)) + NumericUpDown29.Value
        TextBox23.Text = Math.Max(1, (NumericUpDown1.Value - 5)) + NumericUpDown30.Value
        TextBox24.Text = (2 * NumericUpDown2.Value) + NumericUpDown31.Value
        TextBox25.Text = Math.Max(1, Int((NumericUpDown3.Value / 3))) + NumericUpDown32.Value
        TextBox26.Text = NumericUpDown7.Value + NumericUpDown33.Value
        '
        TextBox31.Text = (5 * NumericUpDown3.Value) + NumericUpDown55.Value 'DRPoison
        TextBox32.Text = (2 * NumericUpDown3.Value) + NumericUpDown54.Value 'DRRadiation
    End Sub

    Private Sub SetValueForm_Tab2()
        NumericUpDown40.Value = fCritterPro.b_DTNormal
        NumericUpDown41.Value = fCritterPro.b_DTLaser
        NumericUpDown42.Value = fCritterPro.b_DTFire
        NumericUpDown43.Value = fCritterPro.b_DTPlasma
        NumericUpDown44.Value = fCritterPro.b_DTElectrical
        NumericUpDown45.Value = fCritterPro.b_DTEMP
        NumericUpDown46.Value = fCritterPro.b_DTExplode
        '
        NumericUpDown56.Value = fCritterPro.DTNormal ' + ReverseBytes(fCritterPro.b_DTNormal)
        NumericUpDown57.Value = fCritterPro.DTLaser ' + ReverseBytes(fCritterPro.b_DTLaser)
        NumericUpDown58.Value = fCritterPro.DTFire '+ ReverseBytes(fCritterPro.b_DTFire)
        NumericUpDown59.Value = fCritterPro.DTPlasma '+ ReverseBytes(fCritterPro.b_DTPlasma)
        NumericUpDown60.Value = fCritterPro.DTElectrical '+ ReverseBytes(fCritterPro.b_DTElectrical)
        NumericUpDown61.Value = fCritterPro.DTEMP '+ ReverseBytes(fCritterPro.b_DTEMP)
        NumericUpDown62.Value = fCritterPro.DTExplode '+ ReverseBytes(fCritterPro.b_DTExplode)
        '
        NumericUpDown47.Value = fCritterPro.b_DRNormal
        NumericUpDown48.Value = fCritterPro.b_DRLaser
        NumericUpDown49.Value = fCritterPro.b_DRFire
        NumericUpDown50.Value = fCritterPro.b_DRPlasma
        NumericUpDown51.Value = fCritterPro.b_DRElectrical
        NumericUpDown52.Value = fCritterPro.b_DREMP
        NumericUpDown53.Value = fCritterPro.b_DRExplode
        '
        NumericUpDown71.Value = fCritterPro.DRNormal
        NumericUpDown70.Value = fCritterPro.DRLaser
        NumericUpDown69.Value = fCritterPro.DRFire
        NumericUpDown68.Value = fCritterPro.DRPlasma
        NumericUpDown67.Value = fCritterPro.DRElectrical
        NumericUpDown65.Value = fCritterPro.DREMP
        NumericUpDown63.Value = fCritterPro.DRExplode
        '
        TabView1 = True
    End Sub

    Private Sub SetValueForm_Tab3()
        TextBox30.Text = GetNameCritterMsg(fCritterPro.DescID, True)
        Button4.Enabled = False : Button5.Enabled = False
        NumericUpDown64.Value = fCritterPro.DescID
        '
        If fCritterPro.ScriptID <> -1 Then ComboBox9.SelectedIndex = 1 + (fCritterPro.ScriptID - &H4000000I) Else ComboBox9.SelectedIndex = 0
        ComboBox8.SelectedIndex = fCritterPro.Gender
        ComboBox2.SelectedIndex = fCritterPro.AIPacket
        ComboBox3.SelectedIndex = fCritterPro.TeamNum
        ComboBox4.SelectedIndex = fCritterPro.BodyType
        ComboBox5.SelectedIndex = fCritterPro.DamageType
        ComboBox6.SelectedIndex = fCritterPro.KillType
        '
        NumericUpDown36.Value = fCritterPro.LightDis
        NumericUpDown37.Value = Math.Round(fCritterPro.LightInt * 100 / &HFFFF)
        NumericUpDown38.Value = fCritterPro.ExpVal
        NumericUpDown39.Value = fCritterPro.Age
        'Flags
        CheckBox1.Checked = fCritterPro.Falgs And &H8
        CheckBox2.Checked = fCritterPro.Falgs And &H10
        CheckBox3.Checked = fCritterPro.Falgs And &H800
        CheckBox4.Checked = fCritterPro.Falgs And &H80000000
        CheckBox5.Checked = fCritterPro.Falgs And &H20000000
        'CheckBox12.Checked = ReverseBytes(fCritterPro.Falgs) And &H20
        '
        CheckBox6.Checked = fCritterPro.Falgs And &H8000
        If Not CheckBox6.Checked Then
            RadioButton1.Checked = fCritterPro.Falgs And &H10000
            RadioButton4.Checked = fCritterPro.Falgs And &H20000
            RadioButton2.Checked = fCritterPro.Falgs And &H40000
            RadioButton5.Checked = fCritterPro.Falgs And &H80000
            RadioButton3.Checked = fCritterPro.Falgs And &H4000
        End If
        '
        'CheckBox7.Checked = ReverseBytes(fCritterPro.FalgsExt) And &H800
        'CheckBox8.Checked = ReverseBytes(fCritterPro.FalgsExt) And &H1000
        CheckBox9.Checked = fCritterPro.FalgsExt And &H2000
        CheckBox10.Checked = fCritterPro.FalgsExt And &H4000
        'CheckBox11.Checked = ReverseBytes(fCritterPro.FalgsExt) And &H8000
        '
        CheckBox13.Checked = fCritterPro.CritterFlags And &H20
        CheckBox14.Checked = fCritterPro.CritterFlags And &H40
        CheckBox15.Checked = fCritterPro.CritterFlags And &H80
        CheckBox16.Checked = fCritterPro.CritterFlags And &H2000
        CheckBox17.Checked = fCritterPro.CritterFlags And &H200
        CheckBox18.Checked = fCritterPro.CritterFlags And &H100
        CheckBox19.Checked = fCritterPro.CritterFlags And &H800
        CheckBox20.Checked = fCritterPro.CritterFlags And &H1000
        CheckBox21.Checked = fCritterPro.CritterFlags And &H4000
        CheckBox22.Checked = fCritterPro.CritterFlags And &H400
        CheckBox23.Checked = fCritterPro.CritterFlags And &H2
        '
        TabView2 = True
    End Sub

    Private Sub Save_CritterPro()
        'Tab 1 Special
        fCritterPro.Strength = NumericUpDown1.Value
        fCritterPro.Perception = NumericUpDown2.Value
        fCritterPro.Endurance = NumericUpDown3.Value
        fCritterPro.Charisma = NumericUpDown4.Value
        fCritterPro.Intelligence = NumericUpDown5.Value
        fCritterPro.Agility = NumericUpDown6.Value
        fCritterPro.Luck = NumericUpDown7.Value
        'Skills
        fCritterPro.SmallGuns = NumericUpDown8.Value
        fCritterPro.BigGuns = NumericUpDown9.Value
        fCritterPro.EnergyGun = NumericUpDown10.Value
        fCritterPro.Melee = NumericUpDown11.Value
        fCritterPro.Unarmed = NumericUpDown12.Value
        fCritterPro.Throwing = NumericUpDown13.Value
        fCritterPro.FirstAid = NumericUpDown14.Value
        fCritterPro.Doctor = NumericUpDown15.Value
        fCritterPro.Outdoorsman = NumericUpDown16.Value
        fCritterPro.Sneak = NumericUpDown17.Value
        fCritterPro.Lockpick = NumericUpDown18.Value
        fCritterPro.Steal = NumericUpDown19.Value
        fCritterPro.Traps = NumericUpDown20.Value
        fCritterPro.Science = NumericUpDown21.Value
        fCritterPro.Repair = NumericUpDown22.Value
        fCritterPro.Speech = NumericUpDown23.Value
        fCritterPro.Barter = NumericUpDown24.Value
        fCritterPro.Gambling = NumericUpDown25.Value
        '
        fCritterPro.b_HP = NumericUpDown26.Value
        fCritterPro.b_AC = NumericUpDown27.Value
        fCritterPro.b_AP = NumericUpDown28.Value
        fCritterPro.b_Weight = NumericUpDown29.Value
        fCritterPro.b_MeleeDmg = NumericUpDown30.Value
        fCritterPro.b_Sequence = NumericUpDown31.Value
        fCritterPro.b_Healing = NumericUpDown32.Value
        fCritterPro.b_Critical = NumericUpDown33.Value
        fCritterPro.b_Better = NumericUpDown34.Value
        fCritterPro.b_UnarmedDmg = NumericUpDown35.Value
        '        
        fCritterPro.HP = Val(TextBox19.Text) - NumericUpDown26.Value
        fCritterPro.AC = Val(TextBox20.Text) - NumericUpDown27.Value
        fCritterPro.AP = Val(TextBox21.Text) - NumericUpDown28.Value
        fCritterPro.Weight = Val(TextBox22.Text) - NumericUpDown29.Value
        fCritterPro.MeleeDmg = Val(TextBox23.Text) - NumericUpDown30.Value
        fCritterPro.Sequence = Val(TextBox24.Text) - NumericUpDown31.Value
        fCritterPro.Healing = Val(TextBox25.Text) - NumericUpDown32.Value
        fCritterPro.Critical = Val(TextBox26.Text) - NumericUpDown33.Value
        '*************
        'fCritterPro.Better = ReverseBytes(TextBox27.Text)
        'fCritterPro.UnarmedDmg = ReverseBytes(TextBox28.Text)
        '*************
        'Tab 2 Defence
        fCritterPro.b_DTNormal = NumericUpDown40.Value
        fCritterPro.b_DTLaser = NumericUpDown41.Value
        fCritterPro.b_DTFire = NumericUpDown42.Value
        fCritterPro.b_DTPlasma = NumericUpDown43.Value
        fCritterPro.b_DTElectrical = NumericUpDown44.Value
        fCritterPro.b_DTEMP = NumericUpDown45.Value
        fCritterPro.b_DTExplode = NumericUpDown46.Value
        '
        fCritterPro.DTNormal = NumericUpDown56.Value
        fCritterPro.DTLaser = NumericUpDown57.Value
        fCritterPro.DTFire = NumericUpDown58.Value
        fCritterPro.DTPlasma = NumericUpDown59.Value
        fCritterPro.DTElectrical = NumericUpDown60.Value
        fCritterPro.DTEMP = NumericUpDown61.Value
        fCritterPro.DTExplode = NumericUpDown62.Value
        '
        fCritterPro.b_DRNormal = NumericUpDown47.Value
        fCritterPro.b_DRLaser = NumericUpDown48.Value
        fCritterPro.b_DRFire = NumericUpDown49.Value
        fCritterPro.b_DRPlasma = NumericUpDown50.Value
        fCritterPro.b_DRElectrical = NumericUpDown51.Value
        fCritterPro.b_DREMP = NumericUpDown52.Value
        fCritterPro.b_DRExplode = NumericUpDown53.Value
        fCritterPro.b_DRPoison = NumericUpDown55.Value
        fCritterPro.b_DRRadiation = NumericUpDown54.Value
        '
        fCritterPro.DRNormal = NumericUpDown71.Value
        fCritterPro.DRLaser = NumericUpDown70.Value
        fCritterPro.DRFire = NumericUpDown69.Value
        fCritterPro.DRPlasma = NumericUpDown68.Value
        fCritterPro.DRElectrical = NumericUpDown67.Value
        fCritterPro.DREMP = NumericUpDown65.Value
        fCritterPro.DRExplode = NumericUpDown63.Value
        fCritterPro.DRPoison = Val(TextBox31.Text) - NumericUpDown55.Value
        fCritterPro.DRRadiation = Val(TextBox32.Text) - NumericUpDown54.Value
        '
        'Tab 3
        fCritterPro.DescID = NumericUpDown64.Value
        If ComboBox9.SelectedIndex <> 0 Then fCritterPro.ScriptID = (ComboBox9.SelectedIndex - 1) + &H4000000I Else fCritterPro.ScriptID = &HFFFFFFFF
        fCritterPro.Gender = ComboBox8.SelectedIndex
        fCritterPro.AIPacket = ComboBox2.SelectedIndex
        fCritterPro.TeamNum = ComboBox3.SelectedIndex
        fCritterPro.BodyType = ComboBox4.SelectedIndex
        fCritterPro.DamageType = ComboBox5.SelectedIndex
        fCritterPro.KillType = ComboBox6.SelectedIndex
        '
        fCritterPro.LightDis = NumericUpDown36.Value
        fCritterPro.LightInt = Math.Round((NumericUpDown37.Value * &HFFFF) / 100)
        fCritterPro.ExpVal = NumericUpDown38.Value
        fCritterPro.Age = NumericUpDown39.Value
        '
        Dim flags As Integer = fCritterPro.Falgs
        If CheckBox1.Checked Then flags = flags Or &H8 Else flags = flags And (Not &H8)
        If CheckBox2.Checked Then flags = flags Or &H10 Else flags = flags And (Not &H10)
        If CheckBox3.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox4.Checked Then flags = flags Or &H80000000 Else flags = flags And (Not &H80000000)
        If CheckBox5.Checked Then flags = flags Or &H20000000 Else flags = flags And (Not &H20000000)
        'CheckBox12.Checked = And &H20
        '
        flags = flags And &HFFF0BFFF
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
        fCritterPro.Falgs = flags
        '
        flags = fCritterPro.FalgsExt
        'CheckBox7.Checked =  And &H800
        'CheckBox8.Checked =  And &H1000
        If CheckBox9.Checked Then flags = flags Or &H2000 Else flags = flags And (Not &H2000)
        If CheckBox10.Checked Then flags = flags Or &H4000 Else flags = flags And (Not &H4000)
        'CheckBox11.Checked = And &H8000
        fCritterPro.FalgsExt = flags
        '
        flags = fCritterPro.CritterFlags
        If CheckBox13.Checked Then flags = flags Or &H20 Else flags = flags And (Not &H20)
        If CheckBox14.Checked Then flags = flags Or &H40 Else flags = flags And (Not &H40)
        If CheckBox15.Checked Then flags = flags Or &H80 Else flags = flags And (Not &H80)
        If CheckBox16.Checked Then flags = flags Or &H2000 Else flags = flags And (Not &H2000)
        If CheckBox17.Checked Then flags = flags Or &H200 Else flags = flags And (Not &H200)
        If CheckBox18.Checked Then flags = flags Or &H100 Else flags = flags And (Not &H100)
        If CheckBox19.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox20.Checked Then flags = flags Or &H1000 Else flags = flags And (Not &H1000)
        If CheckBox21.Checked Then flags = flags Or &H4000 Else flags = flags And (Not &H4000)
        If CheckBox22.Checked Then flags = flags Or &H400 Else flags = flags And (Not &H400)
        If CheckBox23.Checked Then flags = flags Or &H2 Else flags = flags And (Not &H2)
        fCritterPro.CritterFlags = flags
        '
        fCritterPro.FrmID = ComboBox1.SelectedIndex + &H1000000I
        '
        'Save to Pro
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\critters") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\critters")
        If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index)) Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        Dim fFile As Byte = FreeFile()
        Dim CrttrProData(103) As Integer
        CrttrProData = fnStructToBytes(fCritterPro)
        For n = 0 To CrttrProData.Length - 1
            CrttrProData(n) = ReverseBytes(CrttrProData(n))
        Next
        FileOpen(fFile, SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, CrttrProData)
        FileClose(fFile)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        'Log
        Main_Form.TextBox1.Text = "Save: " & SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index) & vbCrLf & Main_Form.TextBox1.Text
    End Sub

    Private Sub SaveCritterMsg(ByVal str As String, Optional ByRef Desc As Boolean = False)
        Current_Path = Check_File("Text\English\Game\pro_crit.msg")
        If txtWin Then MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\pro_crit.msg", System.Text.Encoding.Default) _
        Else MSG_DATATEXT = IO.File.ReadAllLines(Current_Path & "\Text\English\Game\pro_crit.msg", System.Text.Encoding.GetEncoding("cp866"))
        '
        If txtLvCp Then CodingToLevCorp(str)
        Dim ID As Integer = NumericUpDown64.Value 'ReverseBytes(fCritterPro.DescID) 
        If Add_ProMSG(str, ID, Desc) Then GoTo MsgError
        'Save
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\Text\English\Game") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\Text\English\Game")
        If txtWin Then IO.File.WriteAllLines(SaveMOD_Path & "\Text\English\Game\pro_crit.msg", MSG_DATATEXT, System.Text.Encoding.Default) _
        Else IO.File.WriteAllLines(SaveMOD_Path & "\Text\English\Game\pro_crit.msg", MSG_DATATEXT, System.Text.Encoding.GetEncoding("cp866"))
        'Log
        Main_Form.TextBox1.Text = "Update: " & SaveMOD_Path & "\Text\English\Game\pro_crit.msg" & vbCrLf & Main_Form.TextBox1.Text
        Exit Sub
MsgError:
        MsgBox("You cannot add a value to Msg file." & vbCrLf & "Not found line: " & ID, MsgBoxStyle.SystemModal And MsgBoxStyle.Critical, "Error")
    End Sub

    Private Sub ValueChanged_Tab1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged, NumericUpDown2.ValueChanged, _
    NumericUpDown3.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown7.ValueChanged, _
    NumericUpDown8.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown12.ValueChanged, _
    NumericUpDown13.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown17.ValueChanged, _
    NumericUpDown18.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown22.ValueChanged, _
    NumericUpDown23.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown27.ValueChanged, _
    NumericUpDown28.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown32.ValueChanged, NumericUpDown33.ValueChanged
        If TabView0 Then CalcSpecialParam() : Button6.Enabled = True
    End Sub

    Private Sub ValueChanged_Tab2a(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown54.ValueChanged, NumericUpDown55.ValueChanged
        If TabView0 = True Then
            TextBox31.Text = (5 * NumericUpDown3.Value) + NumericUpDown55.Value 'DRPoison
            TextBox32.Text = (2 * NumericUpDown3.Value) + NumericUpDown54.Value 'DRRadiation
            Button6.Enabled = True
        End If
    End Sub

    Private Sub ValueChanged_Tab2b(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown56.ValueChanged, NumericUpDown57.ValueChanged, _
    NumericUpDown59.ValueChanged, NumericUpDown60.ValueChanged, NumericUpDown61.ValueChanged, NumericUpDown62.ValueChanged, NumericUpDown63.ValueChanged, NumericUpDown58.ValueChanged, _
    NumericUpDown65.ValueChanged, NumericUpDown67.ValueChanged, NumericUpDown68.ValueChanged, NumericUpDown69.ValueChanged, NumericUpDown70.ValueChanged, NumericUpDown71.ValueChanged, _
    NumericUpDown53.ValueChanged, NumericUpDown52.ValueChanged, NumericUpDown51.ValueChanged, NumericUpDown50.ValueChanged, NumericUpDown49.ValueChanged, NumericUpDown48.ValueChanged, _
    NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, _
    NumericUpDown41.ValueChanged, NumericUpDown40.ValueChanged
        If TabView1 Then Button6.Enabled = True
    End Sub

    Private Sub ValueChanged_Tab3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown64.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown38.ValueChanged, _
    NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged, CheckBox4.CheckedChanged, CheckBox5.CheckedChanged, _
    CheckBox9.CheckedChanged, CheckBox10.CheckedChanged, CheckBox13.CheckedChanged, CheckBox14.CheckedChanged, CheckBox15.CheckedChanged, CheckBox16.CheckedChanged, ComboBox9.SelectedIndexChanged, _
    CheckBox17.CheckedChanged, CheckBox18.CheckedChanged, CheckBox19.CheckedChanged, CheckBox20.CheckedChanged, CheckBox21.CheckedChanged, CheckBox22.CheckedChanged, CheckBox23.CheckedChanged, _
    ComboBox4.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox2.SelectedIndexChanged
        If TabView2 Then Button6.Enabled = True
    End Sub

    Private Sub ValueChanged_Tab3a(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, _
    RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        If TabView2 Then
            Button6.Enabled = True
            TabView2 = False
            CheckBox6.Checked = False
            TabView2 = True
        End If
    End Sub

    Private Sub ValueChanged_Tab3b(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If TabView2 Then
            TabView2 = False
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            TabView2 = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        'ListView1.SelectedItems.Count > 0
        If ListView1.FocusedItem IsNot Nothing Then
            If ListView1.FocusedItem.Index > 0 Then
                Dim cItem As CmItemPro
                Dim aItem As ArItemPro
                Dim ProFile As String = ListView1.FocusedItem.Tag
                Dim iFID As Integer
                Dim fFile As Byte = FreeFile()
                Current_Path = Check_File("proto\items\" & ProFile)
                FileOpen(fFile, Current_Path & "\proto\items\" & ProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
                'FileGet(fFile, iFID, &H35)
                FileGet(fFile, cItem)
                FileGet(fFile, aItem)
                FileClose(fFile)
                If ListView1.FocusedItem.ToolTipText = Nothing Then
                    Dim tip As String
                    tip = "Normal DT: " & ReverseBytes(aItem.DTNormal) & " DR: " & ReverseBytes(aItem.DRNormal) & vbCrLf
                    tip &= "Laser DT: " & ReverseBytes(aItem.DTLaser) & " DR: " & ReverseBytes(aItem.DRLaser) & vbCrLf
                    tip &= "Fire DT: " & ReverseBytes(aItem.DTFire) & " DR: " & ReverseBytes(aItem.DRFire) & vbCrLf
                    tip &= "Plasma DT: " & ReverseBytes(aItem.DTPlasma) & " DR: " & ReverseBytes(aItem.DRPlasma) & vbCrLf
                    tip &= "Electr DT: " & ReverseBytes(aItem.DTElectrical) & " DR: " & ReverseBytes(aItem.DRElectrical) & vbCrLf
                    tip &= "Emp DT: " & ReverseBytes(aItem.DTEMP) & " DR: " & ReverseBytes(aItem.DREMP) & vbCrLf
                    tip &= "Expl DT: " & ReverseBytes(aItem.DTExplode) & " DR: " & ReverseBytes(aItem.DRExplode)
                    ListView1.FocusedItem.ToolTipText = tip
                End If
                If cItem.InvFID = -1 Then
                    PictureBox2.Image = Nothing
                Else
                    iFID = ReverseBytes(cItem.InvFID) - &H7000000
                    ProFile = Strings.Left(Iven_FRM(iFID), RTrim(Iven_FRM(iFID)).Length - 4)
                    If My.Computer.FileSystem.FileExists(Cache_Patch & "\art\inven\" & ProFile & ".gif") = False Then
                        ItemFrmGif("inven\", ProFile)
                    End If
                    PictureBox2.Image = Image.FromFile(Cache_Patch & "\art\inven\" & ProFile & ".gif")
                End If
            Else
                PictureBox2.Image = Nothing
            End If
            Button7.Enabled = True
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim cItem As CmItemPro
        Dim aItem As ArItemPro
        Dim ProFile As String = ListView1.FocusedItem.Tag
        If ProFile = Nothing Then GoTo SetVal
        Dim pth As String = Check_File("proto\items\" & ProFile)
        Dim fFile As Byte = FreeFile()
        FileOpen(fFile, pth & "\proto\items\" & ProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        FileGet(fFile, cItem)
        FileGet(fFile, aItem)
        FileClose(fFile)
SetVal:
        NumericUpDown40.Value = ReverseBytes(aItem.DTNormal)
        NumericUpDown41.Value = ReverseBytes(aItem.DTLaser)
        NumericUpDown42.Value = ReverseBytes(aItem.DTFire)
        NumericUpDown43.Value = ReverseBytes(aItem.DTPlasma)
        NumericUpDown44.Value = ReverseBytes(aItem.DTElectrical)
        NumericUpDown45.Value = ReverseBytes(aItem.DTEMP)
        NumericUpDown46.Value = ReverseBytes(aItem.DTExplode)
        '
        NumericUpDown47.Value = ReverseBytes(aItem.DRNormal)
        NumericUpDown48.Value = ReverseBytes(aItem.DRLaser)
        NumericUpDown49.Value = ReverseBytes(aItem.DRFire)
        NumericUpDown50.Value = ReverseBytes(aItem.DRPlasma)
        NumericUpDown51.Value = ReverseBytes(aItem.DRElectrical)
        NumericUpDown52.Value = ReverseBytes(aItem.DREMP)
        NumericUpDown53.Value = ReverseBytes(aItem.DRExplode)
    End Sub

    Private Sub ComboBox1_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim s As String = ComboBox1.SelectedItem
        Dim n As Byte = InStr(s, ",", CompareMethod.Text)
        Label44.Text = "Frm ID: " & (ComboBox1.SelectedIndex + &H1000000)
        If n <> 0 Then s = Strings.Left(s, n - 1)
        If TabView0 Then Button6.Enabled = True
        On Error GoTo BadFrm
        If My.Computer.FileSystem.FileExists(Cache_Patch & "\art\critters\" & s & "aa.gif") = False Then
            Frm_to_Gif(s)
        End If
        Dim img As Image = Image.FromFile(Cache_Patch & "\art\critters\" & s & "aa.gif")
        If img.Width > PictureBox1.Size.Width Or img.Size.Height > PictureBox1.Size.Height Then PictureBox1.SizeMode = PictureBoxSizeMode.Zoom Else PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
        PictureBox1.Image = img
        Exit Sub
BadFrm:
        PictureBox1.Image = My.Resources.RESERVAA 'My.Resources.ResourceManager.GetObject(UCase(s) & "AA")
    End Sub

    Private Sub Reload(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Ini_CritterForm(fLW_Index)
        TabView0 = False
        TabView1 = False
        TabView2 = False
        SetValueForm_Tab1()
        CalcSpecialParam()
        SetValueForm_Tab2()
        SetValueForm_Tab3()
        Button2.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub Restore(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim tPatch As String = cPath
        If UnDat_File("proto\critters\" & Critter_LST(fLW_Index)) Then
            cPath = Cache_Patch
            Ini_CritterForm(fLW_Index)
            TabView0 = False
            TabView1 = False
            TabView2 = False
            SetValueForm_Tab1()
            CalcSpecialParam()
            SetValueForm_Tab2()
            SetValueForm_Tab3()
            Button6.Enabled = True
        Else
            MsgBox("The prototype this number does not exist in Master.dat.", MsgBoxStyle.Exclamation)
        End If
        cPath = tPatch
    End Sub

    Private Sub Button_Save(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If TabView1 = False Then SetValueForm_Tab2()
        If TabView2 = False Then SetValueForm_Tab3()
        Save_CritterPro()
        Button6.Enabled = False
        cPath = Check_File("proto\critters\" & Critter_LST(fLW_Index))
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SaveCritterMsg(TextBox29.Text)
        Button4.Enabled = False
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        SaveCritterMsg(TextBox30.Text, True)
        Button5.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Critter_Form_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Main_Form.Focus()
        Me.Dispose()
    End Sub

    Private Sub TextBox29_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox29.TextChanged
        Button4.Enabled = True
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Button5.Enabled = True
    End Sub

    Private Sub Critter_Form_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If TabView0 Then Main_Form.ToolStripStatusLabel1.Text = cPath & "\proto\critters\" & Critter_LST(fLW_Index)
    End Sub

    Private Sub Critter_Form_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If Button6.Enabled Then
            Dim btn As MsgBoxResult = MsgBox("Save changes to the Pro-File?", MsgBoxStyle.YesNoCancel, "Attention!")
            If btn = MsgBoxResult.Yes Then
                Button_Save(sender, e)
            ElseIf btn = MsgBoxResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Button6_EnabledChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.EnabledChanged
        If Button6.Enabled Then Button2.Enabled = True
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Create_AIEditForm(fCritterPro.AIPacket)
    End Sub

End Class