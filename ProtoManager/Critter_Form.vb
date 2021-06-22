Imports System.Drawing
Imports System.IO
Imports System.Text

Imports Prototypes
Imports Enums

Friend Class Critter_Form

    Private CritterPro As CritPro

    Private TabStatsView, TabDefenceView, TabMiscView As Boolean
    Private cPath As String = Nothing

    Private ReadOnly cLST_Index As Integer

    Sub New(ByVal cLST_Index As Integer)
        InitializeComponent()
        Me.cLST_Index = cLST_Index
    End Sub

    Friend Function LoadProData() As Boolean
        Dim proFile As String = PROTO_CRITTERS & Critter_LST(cLST_Index).proFile
        If cPath = Nothing Then cPath = DatFiles.CheckFile(proFile, False)

        If ProFiles.LoadCritterProData(cPath & proFile, CritterPro) Then
            'BadFormat
            MsgBox("The pro file " & proFile & " does not have the correct format.", MsgBoxStyle.Critical)
            Return True ' error
        End If

        Return False
    End Function

    Private Sub Critter_Form_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Activated
        If TabStatsView Then Main_Form.ToolStripStatusLabel1.Text = cPath & PROTO_CRITTERS & Critter_LST(cLST_Index).proFile
    End Sub

    Private Sub Critter_Form_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Button6.Enabled Then
            Dim btn As MsgBoxResult = MsgBox("Save changes to the Pro-File?", MsgBoxStyle.YesNoCancel, "Attention!")
            If btn = MsgBoxResult.Yes Then
                Button_Save(sender, e)
            ElseIf btn = MsgBoxResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub Critter_Form_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        If pbCritterFID.Image IsNot Nothing Then pbCritterFID.Image.Dispose()
        If PictureBox2.Image IsNot Nothing Then PictureBox2.Image.Dispose()
        Me.Dispose()
        Main_Form.Focus()
    End Sub

    Private Sub Critter_Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim x As Byte = 1

        For n = 0 To UBound(Items_LST)
            If Items_LST(n).itemType = ItemType.Armor Then
                ListView1.Items.Add(Items_LST(n).itemName)
                ListView1.Items(x).Tag = Items_LST(n).proFile
                x += 1
            End If
        Next

        SetStatsValue_Tab()
        CalcSpecialParam()
    End Sub

    'Возвращает Имя или Описание криттера из msg файла
    Private Function GetNameCritterMsg(ByVal NameID As Integer, Optional ByVal Desc As Boolean = False) As String
        If Desc Then NameID += 1
        Messages.GetMsgData("pro_crit.msg")

        Return Messages.GetNameObject(NameID)
    End Function

    Private Sub TabControl1_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) Handles TabControl1.Selected
        If e.TabPageIndex = 1 And Not (TabDefenceView) Then
            SetDefenceValue_Tab()
        ElseIf e.TabPageIndex = 2 And Not (TabMiscView) Then
            SetMiscValue_Tab()
        End If
    End Sub

    Private Sub SetStatsValue_Tab()
        On Error Resume Next

        TextBox29.Text = GetNameCritterMsg(CritterPro.DescID)
        Me.Text = TextBox29.Text & " [" & Critter_LST(cLST_Index).proFile & "]"

        TextBox33.Text = CritterPro.ProtoID
        ComboBox1.SelectedIndex = CritterPro.FrmID - &H1000000I
        'special
        NumericUpDown1.Value = CritterPro.Strength
        NumericUpDown2.Value = CritterPro.Perception
        NumericUpDown3.Value = CritterPro.Endurance
        NumericUpDown4.Value = CritterPro.Charisma
        NumericUpDown5.Value = CritterPro.Intelligence
        NumericUpDown6.Value = CritterPro.Agility
        NumericUpDown7.Value = CritterPro.Luck
        'Skills
        NumericUpDown8.Value = CritterPro.SmallGuns
        NumericUpDown9.Value = CritterPro.BigGuns
        NumericUpDown10.Value = CritterPro.EnergyGun
        NumericUpDown11.Value = CritterPro.Melee
        NumericUpDown12.Value = CritterPro.Unarmed
        NumericUpDown13.Value = CritterPro.Throwing
        NumericUpDown14.Value = CritterPro.FirstAid
        NumericUpDown15.Value = CritterPro.Doctor
        NumericUpDown16.Value = CritterPro.Outdoorsman
        NumericUpDown17.Value = CritterPro.Sneak
        NumericUpDown18.Value = CritterPro.Lockpick
        NumericUpDown19.Value = CritterPro.Steal
        NumericUpDown20.Value = CritterPro.Traps
        NumericUpDown21.Value = CritterPro.Science
        NumericUpDown22.Value = CritterPro.Repair
        NumericUpDown23.Value = CritterPro.Speech
        NumericUpDown24.Value = CritterPro.Barter
        NumericUpDown25.Value = CritterPro.Gambling

        NumericUpDown26.Value = CritterPro.b_HP
        NumericUpDown27.Value = CritterPro.b_AC
        NumericUpDown28.Value = CritterPro.b_AP
        NumericUpDown29.Value = CritterPro.b_Weight
        NumericUpDown30.Value = CritterPro.b_MeleeDmg
        NumericUpDown31.Value = CritterPro.b_Sequence
        NumericUpDown32.Value = CritterPro.b_Healing
        NumericUpDown33.Value = CritterPro.b_Critical
        NumericUpDown34.Value = CritterPro.b_Better
        NumericUpDown35.Value = CritterPro.b_UnarmedDmg

        'TextBox19.Text = fCritterPro.HP
        'TextBox20.Text = fCritterPro.AC
        'TextBox21.Text = fCritterPro.AP
        'TextBox22.Text = fCritterPro.Weight
        'TextBox23.Text = fCritterPro.MeleeDmg
        'TextBox24.Text = fCritterPro.Sequence
        'TextBox25.Text = fCritterPro.Healing
        'TextBox26.Text = fCritterPro.Critical

        TextBox27.Text = CritterPro.Better.ToString
        TextBox28.Text = CritterPro.UnarmedDmg.ToString

        'TextBox31.Text = fCritterPro.DRPoison
        'TextBox32.Text = fCritterPro.DRRadiation
        NumericUpDown55.Value = CritterPro.b_DRPoison
        NumericUpDown54.Value = CritterPro.b_DRRadiation

        TabStatsView = True
    End Sub

    Private Sub CalcSpecialParam()
        On Error Resume Next

        tbSmallGun.Text = CalcStats.SmallGun_Skill(NumericUpDown6.Value) + NumericUpDown8.Value
        TextBox2.Text = CalcStats.BigEnergyGun_Skill(NumericUpDown6.Value) + NumericUpDown9.Value
        TextBox3.Text = CalcStats.EnergyGun_Skill(NumericUpDown6.Value) + NumericUpDown10.Value
        TextBox5.Text = CalcStats.Unarmed_Skill(NumericUpDown6.Value, NumericUpDown1.Value) + NumericUpDown12.Value
        TextBox4.Text = CalcStats.Melee_Skill(NumericUpDown6.Value, NumericUpDown1.Value) + NumericUpDown11.Value
        TextBox6.Text = CalcStats.Throwing_Skill(NumericUpDown6.Value) + NumericUpDown13.Value
        TextBox7.Text = CalcStats.FirstAid_Skill(NumericUpDown2.Value, NumericUpDown5.Value) + NumericUpDown14.Value
        TextBox8.Text = CalcStats.Doctor_Skill(NumericUpDown2.Value, NumericUpDown5.Value) + NumericUpDown15.Value
        TextBox9.Text = CalcStats.Outdoorsman_Skill(NumericUpDown3.Value, NumericUpDown5.Value) + NumericUpDown16.Value
        TextBox10.Text = CalcStats.Sneak_Skill(NumericUpDown6.Value) + NumericUpDown17.Value
        TextBox11.Text = CalcStats.Lockpick_Skill(NumericUpDown2.Value, NumericUpDown6.Value) + NumericUpDown18.Value
        TextBox12.Text = CalcStats.Steal_Skill(NumericUpDown6.Value) + NumericUpDown19.Value
        TextBox13.Text = CalcStats.Trap_Skill(NumericUpDown2.Value, NumericUpDown6.Value) + NumericUpDown20.Value
        TextBox14.Text = CalcStats.Science_Skill(NumericUpDown5.Value) + NumericUpDown21.Value
        TextBox15.Text = CalcStats.Repair_Skill(NumericUpDown5.Value) + NumericUpDown22.Value
        TextBox16.Text = CalcStats.Speech_Skill(NumericUpDown4.Value) + NumericUpDown23.Value
        TextBox17.Text = CalcStats.Barter_Skill(NumericUpDown4.Value) + NumericUpDown24.Value
        TextBox18.Text = CalcStats.Gamblings_Skill(NumericUpDown7.Value) + NumericUpDown25.Value

        TextBox19.Text = CalcStats.Health_Point(NumericUpDown1.Value, NumericUpDown3.Value) + NumericUpDown26.Value
        TextBox20.Text = NumericUpDown6.Value + NumericUpDown27.Value
        TextBox21.Text = CalcStats.Action_Point(NumericUpDown6.Value) + NumericUpDown28.Value
        TextBox22.Text = CalcStats.Carry_Weight(NumericUpDown1.Value) + NumericUpDown29.Value
        TextBox23.Text = CalcStats.Melee_Damage(NumericUpDown1.Value) + NumericUpDown30.Value
        TextBox24.Text = CalcStats.Sequence(NumericUpDown2.Value) + NumericUpDown31.Value
        TextBox25.Text = CalcStats.Healing_Rate(NumericUpDown3.Value) + NumericUpDown32.Value
        TextBox26.Text = NumericUpDown7.Value + NumericUpDown33.Value
        TextBox27.Text = CritterPro.Better + NumericUpDown34.Value

        TextBox31.Text = CalcStats.Poison(NumericUpDown3.Value) + NumericUpDown55.Value 'DRPoison
        TextBox32.Text = CalcStats.Radiation(NumericUpDown3.Value) + NumericUpDown54.Value 'DRRadiation
    End Sub

    Private Sub SetDefenceValue_Tab()
        On Error Resume Next

        NumericUpDown56.Value = CritterPro.DTNormal ' + fCritterPro.b_DTNormal
        NumericUpDown57.Value = CritterPro.DTLaser ' + fCritterPro.b_DTLaser
        NumericUpDown58.Value = CritterPro.DTFire '+ fCritterPro.b_DTFire
        NumericUpDown59.Value = CritterPro.DTPlasma '+ fCritterPro.b_DTPlasma
        NumericUpDown60.Value = CritterPro.DTElectrical '+ fCritterPro.b_DTElectrical
        NumericUpDown61.Value = CritterPro.DTEMP '+ fCritterPro.b_DTEMP
        NumericUpDown62.Value = CritterPro.DTExplode '+ fCritterPro.b_DTExplode

        NumericUpDown71.Value = CritterPro.DRNormal
        NumericUpDown70.Value = CritterPro.DRLaser
        NumericUpDown69.Value = CritterPro.DRFire
        NumericUpDown68.Value = CritterPro.DRPlasma
        NumericUpDown67.Value = CritterPro.DRElectrical
        NumericUpDown65.Value = CritterPro.DREMP
        NumericUpDown63.Value = CritterPro.DRExplode

        SetBonusDefenceValue()
    End Sub

    Private Sub SetBonusDefenceValue()
        On Error Resume Next

        NumericUpDown40.Value = CritterPro.b_DTNormal
        NumericUpDown41.Value = CritterPro.b_DTLaser
        NumericUpDown42.Value = CritterPro.b_DTFire
        NumericUpDown43.Value = CritterPro.b_DTPlasma
        NumericUpDown44.Value = CritterPro.b_DTElectrical
        NumericUpDown45.Value = CritterPro.b_DTEMP
        NumericUpDown46.Value = CritterPro.b_DTExplode

        NumericUpDown47.Value = CritterPro.b_DRNormal
        NumericUpDown48.Value = CritterPro.b_DRLaser
        NumericUpDown49.Value = CritterPro.b_DRFire
        NumericUpDown50.Value = CritterPro.b_DRPlasma
        NumericUpDown51.Value = CritterPro.b_DRElectrical
        NumericUpDown52.Value = CritterPro.b_DREMP
        NumericUpDown53.Value = CritterPro.b_DRExplode

        TabDefenceView = True
    End Sub

    Private Sub SetMiscValue_Tab()
        On Error Resume Next

        TextBox30.Text = GetNameCritterMsg(CritterPro.DescID, True)
        Button4.Enabled = False
        Button5.Enabled = False

        NumericUpDown64.Value = CritterPro.DescID
        If CritterPro.ScriptID <> -1 Then
            ComboBox9.SelectedIndex = 1 + (CritterPro.ScriptID - &H4000000I)
        Else
            ComboBox9.SelectedIndex = 0
        End If
        ComboBox8.SelectedIndex = CritterPro.Gender
        ComboBox2.SelectedIndex = PacketAI.IndexOfValue(CritterPro.AIPacket)
        ComboBox3.SelectedIndex = CritterPro.TeamNum
        ComboBox4.SelectedIndex = CritterPro.BodyType
        ComboBox5.SelectedIndex = CritterPro.DamageType
        ComboBox6.SelectedIndex = CritterPro.KillType

        NumericUpDown36.Value = CritterPro.LightDis
        NumericUpDown37.Value = Math.Round((CritterPro.LightInt * 100) / &HFFFF)
        NumericUpDown38.Value = CritterPro.ExpVal
        NumericUpDown39.Value = CritterPro.Age
        'Flags
        CheckBox1.Checked = CritterPro.Falgs And &H8
        CheckBox2.Checked = CritterPro.Falgs And &H10
        CheckBox3.Checked = CritterPro.Falgs And &H800
        CheckBox4.Checked = CritterPro.Falgs And &H80000000
        CheckBox5.Checked = CritterPro.Falgs And &H20000000
        'CheckBox12.Checked = fCritterPro.Falgs And &H20

        CheckBox6.Checked = CritterPro.Falgs And &H8000
        If Not CheckBox6.Checked Then
            RadioButton1.Checked = CritterPro.Falgs And &H10000
            RadioButton4.Checked = CritterPro.Falgs And &H20000
            RadioButton2.Checked = CritterPro.Falgs And &H40000
            RadioButton5.Checked = CritterPro.Falgs And &H80000
            RadioButton3.Checked = CritterPro.Falgs And &H4000
        End If

        'CheckBox7.Checked = fCritterPro.FalgsExt And &H800
        'CheckBox8.Checked = fCritterPro.FalgsExt And &H1000
        CheckBox9.Checked = CritterPro.FalgsExt And &H2000
        CheckBox10.Checked = CritterPro.FalgsExt And &H4000
        'CheckBox11.Checked = fCritterPro.FalgsExt And &H8000

        CheckBox13.Checked = CritterPro.CritterFlags And &H20
        CheckBox14.Checked = CritterPro.CritterFlags And &H40
        CheckBox15.Checked = CritterPro.CritterFlags And &H80
        CheckBox16.Checked = CritterPro.CritterFlags And &H2000
        CheckBox17.Checked = CritterPro.CritterFlags And &H200
        CheckBox18.Checked = CritterPro.CritterFlags And &H100
        CheckBox19.Checked = CritterPro.CritterFlags And &H800
        CheckBox20.Checked = CritterPro.CritterFlags And &H1000
        CheckBox21.Checked = CritterPro.CritterFlags And &H4000
        CheckBox22.Checked = CritterPro.CritterFlags And &H400
        CheckBox23.Checked = CritterPro.CritterFlags And &H2

        TabMiscView = True
    End Sub

    Private Sub Save_CritterPro()
        'Tab Special
        CritterPro.Strength = NumericUpDown1.Value
        CritterPro.Perception = NumericUpDown2.Value
        CritterPro.Endurance = NumericUpDown3.Value
        CritterPro.Charisma = NumericUpDown4.Value
        CritterPro.Intelligence = NumericUpDown5.Value
        CritterPro.Agility = NumericUpDown6.Value
        CritterPro.Luck = NumericUpDown7.Value
        'Skills
        CritterPro.SmallGuns = NumericUpDown8.Value
        CritterPro.BigGuns = NumericUpDown9.Value
        CritterPro.EnergyGun = NumericUpDown10.Value
        CritterPro.Melee = NumericUpDown11.Value
        CritterPro.Unarmed = NumericUpDown12.Value
        CritterPro.Throwing = NumericUpDown13.Value
        CritterPro.FirstAid = NumericUpDown14.Value
        CritterPro.Doctor = NumericUpDown15.Value
        CritterPro.Outdoorsman = NumericUpDown16.Value
        CritterPro.Sneak = NumericUpDown17.Value
        CritterPro.Lockpick = NumericUpDown18.Value
        CritterPro.Steal = NumericUpDown19.Value
        CritterPro.Traps = NumericUpDown20.Value
        CritterPro.Science = NumericUpDown21.Value
        CritterPro.Repair = NumericUpDown22.Value
        CritterPro.Speech = NumericUpDown23.Value
        CritterPro.Barter = NumericUpDown24.Value
        CritterPro.Gambling = NumericUpDown25.Value

        CritterPro.b_HP = NumericUpDown26.Value
        CritterPro.b_AC = NumericUpDown27.Value
        CritterPro.b_AP = NumericUpDown28.Value
        CritterPro.b_Weight = NumericUpDown29.Value
        CritterPro.b_MeleeDmg = NumericUpDown30.Value
        CritterPro.b_Sequence = NumericUpDown31.Value
        CritterPro.b_Healing = NumericUpDown32.Value
        CritterPro.b_Critical = NumericUpDown33.Value
        CritterPro.b_Better = NumericUpDown34.Value
        CritterPro.b_UnarmedDmg = NumericUpDown35.Value

        CritterPro.HP = Val(TextBox19.Text) - NumericUpDown26.Value
        CritterPro.AC = Val(TextBox20.Text) - NumericUpDown27.Value
        CritterPro.AP = Val(TextBox21.Text) - NumericUpDown28.Value
        CritterPro.Weight = Val(TextBox22.Text) - NumericUpDown29.Value
        CritterPro.MeleeDmg = Val(TextBox23.Text) - NumericUpDown30.Value
        CritterPro.Sequence = Val(TextBox24.Text) - NumericUpDown31.Value
        CritterPro.Healing = Val(TextBox25.Text) - NumericUpDown32.Value
        CritterPro.Critical = Val(TextBox26.Text) - NumericUpDown33.Value
        '*************
        CritterPro.Better = Val(TextBox27.Text) - NumericUpDown34.Value
        'fCritterPro.UnarmedDmg = TextBox28.Text
        '*************
        'Tab Defence
        CritterPro.b_DTNormal = NumericUpDown40.Value
        CritterPro.b_DTLaser = NumericUpDown41.Value
        CritterPro.b_DTFire = NumericUpDown42.Value
        CritterPro.b_DTPlasma = NumericUpDown43.Value
        CritterPro.b_DTElectrical = NumericUpDown44.Value
        CritterPro.b_DTEMP = NumericUpDown45.Value
        CritterPro.b_DTExplode = NumericUpDown46.Value

        CritterPro.DTNormal = NumericUpDown56.Value
        CritterPro.DTLaser = NumericUpDown57.Value
        CritterPro.DTFire = NumericUpDown58.Value
        CritterPro.DTPlasma = NumericUpDown59.Value
        CritterPro.DTElectrical = NumericUpDown60.Value
        CritterPro.DTEMP = NumericUpDown61.Value
        CritterPro.DTExplode = NumericUpDown62.Value

        CritterPro.b_DRNormal = NumericUpDown47.Value
        CritterPro.b_DRLaser = NumericUpDown48.Value
        CritterPro.b_DRFire = NumericUpDown49.Value
        CritterPro.b_DRPlasma = NumericUpDown50.Value
        CritterPro.b_DRElectrical = NumericUpDown51.Value
        CritterPro.b_DREMP = NumericUpDown52.Value
        CritterPro.b_DRExplode = NumericUpDown53.Value
        CritterPro.b_DRPoison = NumericUpDown55.Value
        CritterPro.b_DRRadiation = NumericUpDown54.Value

        CritterPro.DRNormal = NumericUpDown71.Value
        CritterPro.DRLaser = NumericUpDown70.Value
        CritterPro.DRFire = NumericUpDown69.Value
        CritterPro.DRPlasma = NumericUpDown68.Value
        CritterPro.DRElectrical = NumericUpDown67.Value
        CritterPro.DREMP = NumericUpDown65.Value
        CritterPro.DRExplode = NumericUpDown63.Value
        CritterPro.DRPoison = Val(TextBox31.Text) - NumericUpDown55.Value
        CritterPro.DRRadiation = Val(TextBox32.Text) - NumericUpDown54.Value

        'Tab Misc
        CritterPro.DescID = NumericUpDown64.Value
        If ComboBox9.SelectedIndex <> 0 Then CritterPro.ScriptID = (ComboBox9.SelectedIndex - 1) + &H4000000I Else CritterPro.ScriptID = &HFFFFFFFF
        CritterPro.Gender = ComboBox8.SelectedIndex
        CritterPro.AIPacket = PacketAI.Item(ComboBox2.SelectedItem.ToString)
        CritterPro.TeamNum = ComboBox3.SelectedIndex
        CritterPro.BodyType = ComboBox4.SelectedIndex
        CritterPro.DamageType = ComboBox5.SelectedIndex
        CritterPro.KillType = ComboBox6.SelectedIndex

        CritterPro.LightDis = NumericUpDown36.Value
        CritterPro.LightInt = Math.Round((NumericUpDown37.Value * &HFFFF) / 100)
        CritterPro.ExpVal = NumericUpDown38.Value
        CritterPro.Age = NumericUpDown39.Value

        Dim flags As Integer = CritterPro.Falgs
        If CheckBox1.Checked Then flags = flags Or &H8 Else flags = flags And (Not &H8)
        If CheckBox2.Checked Then flags = flags Or &H10 Else flags = flags And (Not &H10)
        If CheckBox3.Checked Then flags = flags Or &H800 Else flags = flags And (Not &H800)
        If CheckBox4.Checked Then flags = flags Or &H80000000 Else flags = flags And (Not &H80000000)
        If CheckBox5.Checked Then flags = flags Or &H20000000 Else flags = flags And (Not &H20000000)
        'CheckBox12.Checked = And &H20

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
        CritterPro.Falgs = flags

        flags = CritterPro.FalgsExt
        'CheckBox7.Checked =  And &H800
        'CheckBox8.Checked =  And &H1000
        If CheckBox9.Checked Then flags = flags Or &H2000 Else flags = flags And (Not &H2000)
        If CheckBox10.Checked Then flags = flags Or &H4000 Else flags = flags And (Not &H4000)
        'CheckBox11.Checked = And &H8000
        CritterPro.FalgsExt = flags

        flags = CritterPro.CritterFlags
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
        CritterPro.CritterFlags = flags

        CritterPro.FrmID = ComboBox1.SelectedIndex + &H1000000I

        'Save data to Pro file
        Dim proFile As String = SaveMOD_Path & PROTO_CRITTERS
        If Not Directory.Exists(proFile) Then Directory.CreateDirectory(proFile)
        proFile &= Critter_LST(cLST_Index).proFile
        ProFiles.SaveCritterProData(proFile, CritterPro)

        'Log
        Main.PrintLog("Save Pro: " & proFile)
    End Sub

    Private Sub SaveCritterMsg(ByVal str As String, Optional ByRef Desc As Boolean = False)
        Dim ID As Integer = NumericUpDown64.Value 'fCritterPro.DescID 

        Messages.GetMsgData("pro_crit.msg", False)
        If Messages.AddTextMSG(str, ID, Desc) Then
            MsgBox("You can not add value to the Msg file." & vbLf & "Not found msg line #:" & ID, MsgBoxStyle.SystemModal And MsgBoxStyle.Critical, "Error: pro_crit.msg")
            Exit Sub
        End If
        'Save
        Messages.SaveMSGFile("pro_crit.msg")

        'Update Name List
        If Not (Desc) Then
            Button4.Enabled = False
            Critter_LST(cLST_Index).crtName = str
            For Each item As ListViewItem In Main_Form.ListView1.Items
                If CInt(item.Tag) = cLST_Index Then
                    Main_Form.ListView1.Items(item.Index).SubItems(0).Text = str
                    Exit For
                End If
            Next
        Else
            Button5.Enabled = False
        End If
    End Sub

    Private Sub ValueChanged_Tab1(ByVal sender As Object, ByVal e As EventArgs) Handles NumericUpDown1.ValueChanged, NumericUpDown2.ValueChanged, _
    NumericUpDown3.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown7.ValueChanged, _
    NumericUpDown8.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown12.ValueChanged, _
    NumericUpDown13.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown17.ValueChanged, _
    NumericUpDown18.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown22.ValueChanged, _
    NumericUpDown23.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown27.ValueChanged, _
    NumericUpDown28.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown32.ValueChanged, _
    NumericUpDown33.ValueChanged, NumericUpDown34.ValueChanged
        '
        If TabStatsView Then
            CalcSpecialParam()
            Button6.Enabled = True
        Else
            Dim cntl As NumericUpDown = sender
            If cntl.Value <> -100 AndAlso cntl.BackColor <> Color.Black Then
                cntl.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub ValueChanged_Tab2a(ByVal sender As Object, ByVal e As EventArgs) Handles NumericUpDown54.ValueChanged, NumericUpDown55.ValueChanged
        If TabStatsView = True Then
            TextBox31.Text = (5 * NumericUpDown3.Value) + NumericUpDown55.Value 'DRPoison
            TextBox32.Text = (2 * NumericUpDown3.Value) + NumericUpDown54.Value 'DRRadiation
            Button6.Enabled = True
        Else
            Dim cntl As NumericUpDown = sender
            If cntl.Value <> -100 Then
                cntl.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub ValueChanged_Tab2b(ByVal sender As Object, ByVal e As EventArgs) Handles NumericUpDown56.ValueChanged, NumericUpDown57.ValueChanged, _
    NumericUpDown59.ValueChanged, NumericUpDown60.ValueChanged, NumericUpDown61.ValueChanged, NumericUpDown62.ValueChanged, NumericUpDown63.ValueChanged, NumericUpDown58.ValueChanged, _
    NumericUpDown65.ValueChanged, NumericUpDown67.ValueChanged, NumericUpDown68.ValueChanged, NumericUpDown69.ValueChanged, NumericUpDown70.ValueChanged, NumericUpDown71.ValueChanged, _
    NumericUpDown53.ValueChanged, NumericUpDown52.ValueChanged, NumericUpDown51.ValueChanged, NumericUpDown50.ValueChanged, NumericUpDown49.ValueChanged, NumericUpDown48.ValueChanged, _
    NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown42.ValueChanged, _
    NumericUpDown41.ValueChanged, NumericUpDown40.ValueChanged
        '
        If TabDefenceView Then
            Button6.Enabled = True
        Else
            Dim cntl As NumericUpDown = sender
            If cntl.Value <> -999 Then
                cntl.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub ValueChanged_Tab3(ByVal sender As Object, ByVal e As EventArgs) Handles NumericUpDown64.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown38.ValueChanged, _
    NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged, CheckBox4.CheckedChanged, CheckBox5.CheckedChanged, _
    CheckBox9.CheckedChanged, CheckBox10.CheckedChanged, CheckBox13.CheckedChanged, CheckBox14.CheckedChanged, CheckBox15.CheckedChanged, CheckBox16.CheckedChanged, ComboBox9.SelectedIndexChanged, _
    CheckBox17.CheckedChanged, CheckBox18.CheckedChanged, CheckBox19.CheckedChanged, CheckBox20.CheckedChanged, CheckBox21.CheckedChanged, CheckBox22.CheckedChanged, CheckBox23.CheckedChanged, _
    ComboBox4.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox2.SelectedIndexChanged
        '
        If TabMiscView Then Button6.Enabled = True
    End Sub

    Private Sub ValueChanged_Tab3a(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, _
    RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        '
        If TabMiscView Then
            Button6.Enabled = True
            TabMiscView = False
            CheckBox6.Checked = False
            TabMiscView = True
        End If
    End Sub

    Private Sub ValueChanged_Tab3b(ByVal sender As Object, ByVal e As EventArgs) Handles CheckBox6.CheckedChanged
        If TabMiscView Then
            TabMiscView = False
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            RadioButton3.Checked = False
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            TabMiscView = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.FocusedItem IsNot Nothing Then
            If ListView1.FocusedItem.Index > 0 Then
                Dim aItem As ArItemPro
                Dim InvFID As Integer
                Dim fFile As Integer = FreeFile()
                Dim ProFile As String = ListView1.FocusedItem.Tag

                Dim filePath = DatFiles.CheckFile(PROTO_ITEMS & ProFile)
                FileOpen(fFile, filePath, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
                FileGet(fFile, InvFID, Prototypes.offsetInvenFID + 1)
                FileGet(fFile, aItem, Prototypes.ArmorBlock)
                FileClose(fFile)

                If ListView1.FocusedItem.ToolTipText IsNot Nothing Then
                    Dim tip As StringBuilder = New StringBuilder("Armor characteristic:")
                    tip.AppendLine()
                    tip.AppendLine("Normal DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTNormal), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRNormal))
                    tip.AppendLine(" Laser DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTLaser), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRLaser))
                    tip.AppendLine("  Fire DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTFire), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRFire))
                    tip.AppendLine("Plasma DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTPlasma), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRPlasma))
                    tip.AppendLine("Electr DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTElectrical), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRElectrical))
                    tip.AppendLine("   Emp DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTEMP), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DREMP))
                    tip.Append("  Expl DT:" & Strings.RSet(ProFiles.ReverseBytes(aItem.DTExplode), 3) & " | DR: " & ProFiles.ReverseBytes(aItem.DRExplode))
                    ListView1.FocusedItem.ToolTipText = tip.ToString
                End If

                If InvFID = -1 Then
                    PictureBox2.Image = Nothing
                Else
                    Main.GetItemsLstFRM()
                    InvFID = ProFiles.ReverseBytes(InvFID) - &H7000000
                    ProFile = Strings.Left(Iven_FRM(InvFID), RTrim(Iven_FRM(InvFID)).Length - 4)
                    If Not File.Exists(Cache_Patch & ART_INVEN & ProFile & ".gif") Then
                        DatFiles.ItemFrmToGif("inven\", ProFile)
                    End If
                    PictureBox2.Image = Image.FromFile(Cache_Patch & ART_INVEN & ProFile & ".gif")
                End If
            Else
                PictureBox2.Image = Nothing
            End If
            Button7.Enabled = True
        End If
    End Sub

    Private Sub ArmorApply(ByVal sender As Object, ByVal e As EventArgs) Handles Button7.Click
        Dim aItem As ArItemPro
        Dim proFile As String = ListView1.FocusedItem.Tag

        If proFile IsNot Nothing Then
            Dim fFile As Integer = FreeFile()
            proFile = DatFiles.CheckFile(PROTO_ITEMS & proFile)
            FileOpen(fFile, proFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            FileGet(fFile, aItem, Prototypes.ArmorBlock)
            FileClose(fFile)
        Else
            SetBonusDefenceValue()
            Exit Sub
        End If

        NumericUpDown40.Value = ProFiles.ReverseBytes(aItem.DTNormal)
        NumericUpDown41.Value = ProFiles.ReverseBytes(aItem.DTLaser)
        NumericUpDown42.Value = ProFiles.ReverseBytes(aItem.DTFire)
        NumericUpDown43.Value = ProFiles.ReverseBytes(aItem.DTPlasma)
        NumericUpDown44.Value = ProFiles.ReverseBytes(aItem.DTElectrical)
        NumericUpDown45.Value = ProFiles.ReverseBytes(aItem.DTEMP)
        NumericUpDown46.Value = ProFiles.ReverseBytes(aItem.DTExplode)

        NumericUpDown47.Value = ProFiles.ReverseBytes(aItem.DRNormal)
        NumericUpDown48.Value = ProFiles.ReverseBytes(aItem.DRLaser)
        NumericUpDown49.Value = ProFiles.ReverseBytes(aItem.DRFire)
        NumericUpDown50.Value = ProFiles.ReverseBytes(aItem.DRPlasma)
        NumericUpDown51.Value = ProFiles.ReverseBytes(aItem.DRElectrical)
        NumericUpDown52.Value = ProFiles.ReverseBytes(aItem.DREMP)
        NumericUpDown53.Value = ProFiles.ReverseBytes(aItem.DRExplode)
    End Sub

    Private Sub ComboBox1_Changed(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        tbFrmID.Text = (ComboBox1.SelectedIndex + &H1000000)

        Dim frm As String = ComboBox1.SelectedItem
        Dim fileFrm As String = Cache_Patch & ART_CRITTERS_PATH & frm & "aa.gif"
        If Not File.Exists(fileFrm) Then DatFiles.CritterFrmToGif(frm)

        Dim img As Bitmap = My.Resources.RESERVAA 'BadFrm
        If File.Exists(fileFrm) Then
            img = New Bitmap(fileFrm)
            If img.Width > pbCritterFID.Size.Width OrElse img.Size.Height > pbCritterFID.Size.Height Then
                If img.GetFrameCount(Imaging.FrameDimension.Time) = 1 Then img.MakeTransparent(Color.White)
                pbCritterFID.SizeMode = PictureBoxSizeMode.Zoom
            Else
                pbCritterFID.SizeMode = PictureBoxSizeMode.CenterImage
            End If
            If TabStatsView Then Button6.Enabled = True
        End If

        pbCritterFID.Image = img
    End Sub

    Private Sub InitFormData()
        LoadProData()
        TabStatsView = False
        TabDefenceView = False
        TabMiscView = False
        SetStatsValue_Tab()
        CalcSpecialParam()
        SetDefenceValue_Tab()
        SetMiscValue_Tab()
    End Sub

    Private Sub Reload(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        InitFormData()
        Button2.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub Restore(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim tempPath As String = cPath

        If DatFiles.UnDatFile(PROTO_CRITTERS & Critter_LST(cLST_Index).proFile, 416) Then
            cPath = Cache_Patch
            InitFormData()
            Button6.Enabled = True
        Else
            MsgBox("The prototype this number does not exist in Master.dat.", MsgBoxStyle.Exclamation)
        End If

        cPath = tempPath
    End Sub

    Private Sub Button_Save(ByVal sender As Object, ByVal e As EventArgs) Handles Button6.Click
        If TabDefenceView = False Then SetDefenceValue_Tab()
        If TabMiscView = False Then SetMiscValue_Tab()
        Save_CritterPro()
        For Each item As ListViewItem In Main_Form.ListView1.Items
            If CInt(item.Tag) = cLST_Index Then
                Main_Form.ListView1.Items(item.Index).ForeColor = Color.DarkBlue
                Main_Form.ListView1.Items(item.Index).SubItems(2).Text = If(Settings.proRO, "R/O", String.Empty)
                Exit For
            End If
        Next
        Button6.Enabled = False
        cPath = DatFiles.CheckFile(PROTO_CRITTERS & Critter_LST(cLST_Index).proFile, False)
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        SaveCritterMsg(TextBox29.Text)
    End Sub

    Private Sub Button5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button5.Click
        SaveCritterMsg(TextBox30.Text, True)
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub TextBox29_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox29.TextChanged
        Button4.Enabled = True
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox30.TextChanged
        Button5.Enabled = True
    End Sub

    Private Sub Button6_EnabledChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Button6.EnabledChanged
        If Button6.Enabled Then Button2.Enabled = True
    End Sub

    Private Sub Button8_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button8.Click
        Main.Create_AIEditForm(CritterPro.AIPacket)
    End Sub

    Private Sub NumericUpDown_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles NumericUpDown71.KeyPress, NumericUpDown70.KeyPress, NumericUpDown69.KeyPress,
        NumericUpDown68.KeyPress, NumericUpDown67.KeyPress, NumericUpDown65.KeyPress, NumericUpDown64.KeyPress, NumericUpDown63.KeyPress, NumericUpDown62.KeyPress,
        NumericUpDown61.KeyPress, NumericUpDown60.KeyPress, NumericUpDown59.KeyPress, NumericUpDown58.KeyPress, NumericUpDown57.KeyPress, NumericUpDown56.KeyPress,
        NumericUpDown53.KeyPress, NumericUpDown52.KeyPress, NumericUpDown51.KeyPress, NumericUpDown50.KeyPress, NumericUpDown49.KeyPress, NumericUpDown48.KeyPress,
        NumericUpDown47.KeyPress, NumericUpDown46.KeyPress, NumericUpDown45.KeyPress, NumericUpDown44.KeyPress, NumericUpDown43.KeyPress, NumericUpDown42.KeyPress,
        NumericUpDown41.KeyPress, NumericUpDown40.KeyPress, NumericUpDown39.KeyPress, NumericUpDown38.KeyPress, NumericUpDown37.KeyPress, NumericUpDown36.KeyPress
        '
        If Char.IsDigit(e.KeyChar) Then Button6.Enabled = True
    End Sub

End Class