Public Class TxtEdit_Form
    Private LocIndex As Integer
    Private fLW_Index As Integer

    Private CrttrProData(103) As Integer
    Private ItemProData(13) As Integer '1+byte
    Private ArmItmProData(17) As Integer
    Private DrgItmProData(16) As Integer
    Private WpnItmProData(14) As Integer '1+byte
    Private AmmItmProData(5) As Integer
    Private MscItmProData(2) As Integer
    Private CntItmProData(1) As Integer
    Private KeyItmProData As Integer
    Private SndProData As Byte
    Private wSndProData As Byte

    Private CrtNamePro() As String = {
   "pid",
   "msg_id",
   "fid",
   "light_distance",
   "light_intensity",
   "flags",
   "flags_ext",
   "script_id",
   "head_fid",
   "ai_packet",
   "team_num",
   "critter_flags",
   "base_stat_srength",
   "base_stat_prception",
   "base_stat_endurance",
   "base_stat_charisma",
   "base_stat_intelligence",
   "base_stat_agility",
   "base_stat_luck",
   "base_stat_hp",
   "base_stat_ap",
   "base_stat_ac",
   "base_stat_unarmed_damage",
   "base_stat_melee_damage",
   "base_stat_carry_weight",
   "base_stat_sequence",
   "base_stat_healing_rate",
   "base_stat_critical_chance",
   "base_stat_better_criticals",
   "base_dt_normal",
   "base_dt_laser",
   "base_dt_fire",
   "base_dt_plasma",
   "base_dt_electrical",
   "base_dt_emp",
   "base_dt_explode",
   "base_dr_normal",
   "base_dr_laser",
   "base_dr_fire",
   "base_dr_plasma",
   "base_dr_electrical",
   "base_dr_emp",
   "base_dr_explode",
   "base_dr_radiation",
   "base_dr_poison",
   "base_age",
   "base_gender",
   "bonus_stat_srength",
   "bonus_stat_prception",
   "bonus_stat_endurance",
   "bonus_stat_charisma",
   "bonus_stat_intelligence",
   "bonus_stat_agility",
   "bonus_stat_luck",
   "bonus_stat_hp",
   "bonus_stat_ap",
   "bonus_stat_ac",
   "bonus_stat_unarmed_damage",
   "bonus_stat_melee_damage",
   "bonus_stat_carry_weight",
   "bonus_stat_sequence",
   "bonus_stat_healing_rate",
   "bonus_stat_critical_chance",
   "bonus_stat_better_criticals",
   "bonus_dt_normal",
   "bonus_dt_laser",
   "bonus_dt_fire",
   "bonus_dt_plasma",
   "bonus_dt_electrical",
   "bonus_dt_emp",
   "bonus_dt_explode",
   "bonus_dr_normal",
   "bonus_dr_laser",
   "bonus_dr_fire",
   "bonus_dr_plasma",
   "bonus_dr_electrical",
   "bonus_dr_emp",
   "bonus_dr_explode",
   "bonus_dr_radiation",
   "bonus_dr_poison",
   "bonus_age",
   "bonus_gender",
   "skill_small_guns",
   "skill_big_guns",
   "skill_energy_weapons",
   "skill_unarmed",
   "skill_melee",
   "skill_throwing",
   "skill_first_aid",
   "skill_doctor",
   "skill_sneak",
   "skill_lockpick",
   "skill_steal",
   "skill_traps",
   "skill_science",
   "skill_repair",
   "skill_speech",
   "skill_barter",
   "skill_gambling",
   "skill_outdoorsman",
   "body_type",
   "exp_val",
   "kill_type",
   "damage_type"}

    Private Const SndNamePro As String = "sound_id"
    Private Const wSndNamePro As String = "weapon_sound_id"
    Private Const KeyNamePro As String = "unknown"

    Private ItmNamePro() As String = {
       "pid",
       "msg_id",
       "fid",
       "light_distance",
       "light_intensity",
       "flags",
       "flags_ext",
       "script_id",
       "obj_subtype",
       "material_id",
       "size",
       "weight",
       "cost",
       "inventory_fid"}

    Private ArmNamePro() As String = {
       "armor_class",
       "armor_dt_normal",
       "armor_dt_laser",
       "armor_dt_fire",
       "armor_dt_plasma",
       "armor_dt_electrical",
       "armor_dt_emp",
       "armor_dt_explode",
       "armor_dr_normal",
       "armor_dr_laser",
       "armor_dr_fire",
       "armor_dr_plasma",
       "armor_dr_electrical",
       "armor_dr_emp",
       "armor_dr_explode",
       "armor_perk_id",
       "armor_male_fid",
       "armor_female_fid"}

    Private DrgNamePro() As String = {
     "modify_stat0",
     "modify_stat1",
     "modify_stat2",
     "instant_amount0",
     "instant_amount1",
     "instant_amount2",
     "first_duration",
     "first_amount0",
     "first_amount1",
     "first_amount2",
     "second_duration",
     "second_amount0",
     "second_amount1",
     "second_amount2",
     "addiction_rate",
     "addiction_effect",
     "addiction_onset_time"}

    Private WpnNamePro() As String = {
     "weapon_anim_code",
     "min_dmg",
     "max_dmg",
     "dmg_type",
     "primary_attack_max_range",
     "secondary_attack_max_range",
     "proj_pid",
     "min_str",
     "primary_attack_ap_cost",
     "secondary_attack_ap_cost",
     "critical_fail",
     "weapon_perk",
     "weapon_rounds",
     "weapon_caliber",
     "weapon_ammo_pid",
     "weapon_sound_id"}

    Private AmmNamePro() As String = {
     "ammo_caliber",
     "ammo_quantity",
     "ammo_ac_adjust",
     "ammo_dr_adjust",
     "ammo_damage_mult",
     "ammo_damage_div"}

    Private MscNamePro() As String = {
     "misc_power_pid",
     "misc_power_type",
     "misc_charges"}

    Private CntNamePro() As String = {
       "container_max_size",
       "container_open_flags"}

   
    Friend Sub Ini_Form(ByRef Lw_Index As UShort, ByRef Type As Byte)
        Dim n, m As Integer
        Dim cPath As String
        Dim fFile As Byte = FreeFile()

        On Error GoTo BadFormat
        If Type = 0 Then
            cPath = Check_File("proto\critters\" & Critter_LST(Lw_Index))
            FileOpen(fFile, cPath & "\proto\critters\" & Critter_LST(Lw_Index), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            FileGet(fFile, CrttrProData)
            FileClose(fFile)
            For n = 0 To CrttrProData.Length - 1
                CrttrProData(n) = ReverseBytes(CrttrProData(n))
            Next
            If Not Me.Visible Then Me.Text = GetNameMsg(CrttrProData(1), False) & Me.Text
            For n = 0 To CrttrProData.Length - 1
                ListView1.Items.Add(CrtNamePro(n))                  'param
                ListView1.Items(n).SubItems.Add(CrttrProData(n))    'value
                ListView1.Items(n).SubItems.Add("0x" + Hex(CrttrProData(n)))    'hex
                If (n > 63 And n < 80) Or (n > 28 And n < 45) Then
                    ListView1.Items(n).Group = ListView1.Groups.Item(3)  'arror
                ElseIf (n > 11 And n < 29) Or (n >= 47 And n < 64) Then
                    ListView1.Items(n).Group = ListView1.Groups.Item(1) 'stat
                ElseIf (n > 81 And n < 100) Then
                    ListView1.Items(n).Group = ListView1.Groups.Item(2) 'SKILLS
                Else
                    ListView1.Items(n).Group = ListView1.Groups.Item(0)
                End If
            Next
        Else
            cPath = Check_File("proto\items\" & Items_LST(Lw_Index, 0))
            FileOpen(fFile, cPath & "\proto\items\" & Items_LST(Lw_Index, 0), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            FileGet(fFile, ItemProData)
            FileGet(fFile, SndProData)
            For n = 0 To ItemProData.Length - 1
                ItemProData(n) = ReverseBytes(ItemProData(n))
                ListView1.Items.Add(ItmNamePro(n))                          'param
                ListView1.Items(n).SubItems.Add(ItemProData(n))             'value
                ListView1.Items(n).SubItems.Add("0x" + Hex(ItemProData(n))) 'hex
                ListView1.Items(n).Group = ListView1.Groups.Item(0)
            Next
            ListView1.Items.Add(SndNamePro)                          'param
            ListView1.Items(n).SubItems.Add(SndProData)             'value
            ListView1.Items(n).SubItems.Add("0x" + Hex(SndProData)) 'hex
            ListView1.Items(n).Group = ListView1.Groups.Item(0)
            m = n + 1
            If Not Me.Visible Then Me.Text = GetNameMsg(ItemProData(1), True) & Me.Text
            Select Case ItemProData(8)
                Case 3 '"Weapon"
                    FileGet(fFile, WpnItmProData)
                    FileGet(fFile, wSndProData)
                    For n = 0 To WpnItmProData.Length - 1
                        WpnItmProData(n) = ReverseBytes(WpnItmProData(n))
                        ListView1.Items.Add(WpnNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(WpnItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(WpnItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(4)
                    Next
                    ListView1.Items.Add(wSndNamePro)                  'param
                    ListView1.Items(m + n).SubItems.Add(wSndProData)    'value
                    ListView1.Items(m + n).SubItems.Add("0x" + Hex(wSndProData))    'hex
                    ListView1.Items(m + n).Group = ListView1.Groups.Item(4)
                Case 0 '"Armor"
                    FileGet(fFile, ArmItmProData)
                    For n = 0 To ArmItmProData.Length - 1
                        ArmItmProData(n) = ReverseBytes(ArmItmProData(n))
                        ListView1.Items.Add(ArmNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(ArmItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(ArmItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(6)
                    Next
                Case 4 '"Ammo"
                    FileGet(fFile, AmmItmProData)
                    For n = 0 To AmmItmProData.Length - 1
                        AmmItmProData(n) = ReverseBytes(AmmItmProData(n))
                        ListView1.Items.Add(AmmNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(AmmItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(AmmItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(5)
                    Next
                Case 1 '"Container"
                    FileGet(fFile, CntItmProData)
                    For n = 0 To CntItmProData.Length - 1
                        CntItmProData(n) = ReverseBytes(CntItmProData(n))
                        ListView1.Items.Add(CntNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(CntItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(CntItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(8)
                    Next
                Case 2 '"Drugs"
                    FileGet(fFile, DrgItmProData)
                    For n = 0 To DrgItmProData.Length - 1
                        DrgItmProData(n) = ReverseBytes(DrgItmProData(n))
                        ListView1.Items.Add(DrgNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(DrgItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(DrgItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(7)
                    Next
                Case 5 '"Misc"
                    FileGet(fFile, MscItmProData)
                    For n = 0 To MscItmProData.Length - 1
                        MscItmProData(n) = ReverseBytes(MscItmProData(n))
                        ListView1.Items.Add(MscNamePro(n))                  'param
                        ListView1.Items(m + n).SubItems.Add(MscItmProData(n))    'value
                        ListView1.Items(m + n).SubItems.Add("0x" + Hex(MscItmProData(n)))    'hex
                        ListView1.Items(m + n).Group = ListView1.Groups.Item(8)
                    Next
                Case Else '"Key"
                    FileGet(fFile, KeyItmProData)
                    KeyItmProData = ReverseBytes(KeyItmProData)
                    ListView1.Items.Add(KeyNamePro)                  'param
                    ListView1.Items(m).SubItems.Add(KeyItmProData)    'value
                    ListView1.Items(m).SubItems.Add("0x" + Hex(KeyItmProData))    'hex
                    ListView1.Items(m).Group = ListView1.Groups.Item(8)
            End Select
            FileClose(fFile)
        End If
        On Error GoTo 0
        If Not (Me.Visible) Then
            fLW_Index = Lw_Index
            Me.Tag = Type
        End If
        Me.Show()
        Exit Sub
BadFormat:
        FileClose()
        MsgBox("Error reading Pro file, maybe he does have the not correct format.", MsgBoxStyle.Critical)
        Main_Form.Focus()
        Me.Dispose()
    End Sub

    Private Function GetNameMsg(ByRef NameID As Integer, ByRef msg As Boolean) As String
        If msg Then GetMsgData("pro_item.msg") Else GetMsgData("pro_crit.msg")
        Return GetNameCritter(NameID)
    End Function

    Private Sub ListView1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        LocIndex = ListView1.FocusedItem.Index
        TextBox1.Enabled = True
        TextBox1.Text = ListView1.Items.Item(LocIndex).SubItems(1).Text
        If CheckBox1.Checked Then TextBox1.Text = Hex(TextBox1.Text)
        TextBox1.Tag = TextBox1.Text
        If ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.MistyRose Then
            ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.Pink
        Else
            ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.LightPink
        End If
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.Leave
        Dim TxtChangEn As Boolean = False
        On Error Resume Next
        TextBox1.Enabled = False
        If TextBox1.Text <> TextBox1.Tag And TextBox1.Text.Length > 0 Then TxtChangEn = True

        If TxtChangEn Or ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.Pink Then
            ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.MistyRose
        Else
            ListView1.Items.Item(LocIndex).BackColor = Drawing.Color.White
        End If
        If TxtChangEn Then
            If CheckBox1.Checked Then TextBox1.Text = CStr(Integer.Parse(TextBox1.Text, Globalization.NumberStyles.HexNumber))
            ListView1.Items.Item(LocIndex).SubItems(1).Text = TextBox1.Text
            ListView1.Items.Item(LocIndex).SubItems(2).Text = "0x" & Convert.ToString(CInt(TextBox1.Text), 16).ToUpper
            If Me.Tag = 0 Then
                CrttrProData(LocIndex) = CInt(TextBox1.Text)
            Else
                If LocIndex <= 13 Then
                    ItemProData(LocIndex) = CInt(TextBox1.Text)
                Else
                    If LocIndex = 14 Then
                        SndProData = CInt(TextBox1.Text)
                    Else
                        LocIndex -= 15
                        Select Case ItemProData(8)
                            Case 3 '"Weapon"
                                If LocIndex > 14 Then
                                    wSndProData = CInt(TextBox1.Text)
                                Else
                                     WpnItmProData(LocIndex) = CInt(TextBox1.Text)
                                End If
                            Case 0 '"Armor"
                                ArmItmProData(LocIndex) = CInt(TextBox1.Text)
                            Case 4 '"Ammo"
                                AmmItmProData(LocIndex) = CInt(TextBox1.Text)
                            Case 1 '"Container"
                                CntItmProData(LocIndex) = CInt(TextBox1.Text)
                            Case 2 '"Drugs"
                                DrgItmProData(LocIndex) = CInt(TextBox1.Text)
                            Case 5 '"Misc"
                                MscItmProData(LocIndex) = CInt(TextBox1.Text)
                            Case Else '"Key"
                                KeyItmProData = CInt(TextBox1.Text)
                        End Select
                    End If
                End If
            End If
        End If
        TextBox1.Text = Nothing
    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        ListView1.ShowGroups = Not (ListView1.ShowGroups)
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ListView1.Focus()
        ElseIf e.KeyCode = Keys.Escape Then
            TextBox1.Text = TextBox1.Tag
            ListView1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not (e.KeyChar = ChrW(Keys.Back) Or e.KeyChar = ChrW(Keys.Delete) Or e.KeyChar = "-") Then
            If CheckBox1.Checked = False Then
                If Not (Char.IsDigit(e.KeyChar)) Then e.Handled = True
            Else
                e.KeyChar = Char.ToUpper(e.KeyChar)
                If Not (Char.IsDigit(e.KeyChar)) And Not (e.KeyChar >= "A" And e.KeyChar <= "F") Then e.Handled = True
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ListView1.Items.Clear()
        ListView1.ShowGroups = True
        Ini_Form(fLW_Index, Me.Tag)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fFile As Byte = FreeFile()
        If Me.Tag = 0 Then
            'Save to Pro critter
            For n = 0 To CrttrProData.Length - 1
                CrttrProData(n) = ReverseBytes(CrttrProData(n))
            Next
            If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\critters") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\critters")
            If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index)) Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            FileOpen(fFile, SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
            FilePut(fFile, CrttrProData)
            FileClose(fFile)
            If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index), IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            'log
            Main_Form.TextBox1.Text = "Save: " & SaveMOD_Path & "\proto\critters\" & Critter_LST(fLW_Index) & vbCrLf & Main_Form.TextBox1.Text
        Else
            'Save to Pro item
            Dim Type As Integer = ItemProData(8)
            For n = 0 To ItemProData.Length - 1
                ItemProData(n) = ReverseBytes(ItemProData(n))
            Next
            'SndProData = ReverseBytes(SndProData)
            If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\items") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\items")
            If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0)) Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0), IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            FileOpen(fFile, SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0), OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
            FilePut(fFile, ItemProData)
            FilePut(fFile, SndProData)
            Select Case Type
                Case 3 '"Weapon"
                    For n = 0 To WpnItmProData.Length - 1
                        WpnItmProData(n) = ReverseBytes(WpnItmProData(n))
                    Next
                    'wSndProData = ReverseBytes(wSndProData)
                    FilePut(fFile, WpnItmProData)
                    FilePut(fFile, wSndProData)
                Case 0 '"Armor"
                    For n = 0 To ArmItmProData.Length - 1
                        ArmItmProData(n) = ReverseBytes(ArmItmProData(n))
                    Next
                    FilePut(fFile, ArmItmProData)
                Case 4 '"Ammo"
                    For n = 0 To AmmItmProData.Length - 1
                        AmmItmProData(n) = ReverseBytes(AmmItmProData(n))
                    Next
                    FilePut(fFile, AmmItmProData)
                Case 1 '"Container"
                    For n = 0 To CntItmProData.Length - 1
                        CntItmProData(n) = ReverseBytes(CntItmProData(n))
                    Next
                    FilePut(fFile, CntItmProData)
                Case 2 '"Drugs"
                    For n = 0 To DrgItmProData.Length - 1
                        DrgItmProData(n) = ReverseBytes(DrgItmProData(n))
                    Next
                    FilePut(fFile, DrgItmProData)
                Case 5 '"Misc"
                    For n = 0 To MscItmProData.Length - 1
                        MscItmProData(n) = ReverseBytes(MscItmProData(n))
                    Next
                    FilePut(fFile, MscItmProData)
                Case 6 '"Key"
                    FilePut(fFile, KeyItmProData)
                Case Else
                    MsgBox("Invalid object type!", MsgBoxStyle.Critical, "Error Pro")
                    FileClose(fFile)
                    My.Computer.FileSystem.DeleteFile(SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0))
                    'log
                    Main_Form.TextBox1.Text = "Delete: " & SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0) & vbCrLf & Main_Form.TextBox1.Text
                    Exit Sub
            End Select
            FileClose(fFile)
            If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0), IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            'log
            Main_Form.TextBox1.Text = "Save: " & SaveMOD_Path & "\proto\items\" & Items_LST(fLW_Index, 0) & vbCrLf & Main_Form.TextBox1.Text
        End If
    End Sub

End Class