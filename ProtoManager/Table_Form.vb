Public Class Table_Form
    Dim CheckedList() As String
    Dim Table() As String

    Private Const spr As String = ";"

    Private CommonItem As CmItemPro
    Private WeaponItem As WpItemPro
    Private ArmorItem As ArItemPro
    Private AmmoItem As AmItemPro
    Private DrugItem As DgItemPro
    Private MiscItem As McItemPro

    Private Sub CheckAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckAllToolStripMenuItem.Click
        DeSelItemsAll(True)
    End Sub

    Private Sub DeselecAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeselecAllToolStripMenuItem.Click
        DeSelItemsAll(False)
    End Sub

    Private Sub DeSelItemsAll(ByVal value As Boolean)
        Select Case TabControl1.SelectedIndex
            Case 0 'Critter
                For n = 0 To CheckedListBox6.Items.Count - 1
                    CheckedListBox6.SetItemChecked(n, value)
                Next
            Case 1 'Weapon
                For n = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SetItemChecked(n, value)
                Next
            Case 2 'Ammo
                For n = 0 To CheckedListBox2.Items.Count - 1
                    CheckedListBox2.SetItemChecked(n, value)
                Next
            Case 3 'Armor
                For n = 0 To CheckedListBox3.Items.Count - 1
                    CheckedListBox3.SetItemChecked(n, value)
                Next
            Case 4 'Drugs
                For n = 0 To CheckedListBox4.Items.Count - 1
                    CheckedListBox4.SetItemChecked(n, value)
                Next
            Case Else
                For n = 0 To CheckedListBox5.Items.Count - 1
                    CheckedListBox5.SetItemChecked(n, value)
                Next
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Select Case TabControl1.SelectedIndex
            Case 0 'Critter
                Dim count As Byte = CheckedListBox6.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox6.CheckedItems(n).ToString
                Next
            Case 1 'Weapon
                Dim count As Byte = CheckedListBox1.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox1.CheckedItems(n).ToString
                Next
            Case 2 'Ammo
                Dim count As Byte = CheckedListBox2.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox2.CheckedItems(n).ToString
                Next
            Case 3 'Armor
                Dim count As Byte = CheckedListBox3.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox3.CheckedItems(n).ToString
                Next
            Case 4 'Drugs
                Dim count As Byte = CheckedListBox4.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox4.CheckedItems(n).ToString
                Next
            Case Else
                Dim count As Byte = CheckedListBox5.CheckedIndices.Count - 1
                ReDim CheckedList(count)
                For n = 0 To count
                    CheckedList(n) = CheckedListBox5.CheckedItems(n).ToString
                Next
        End Select
        CreateTable(TabControl1.SelectedTab.Text)
    End Sub

    Private Sub CreateTable(ByVal iType As String)
        Dim x As UShort, n As Integer, m As Byte
        ReDim Table(0)
        Table(x) = "Import" & spr & "ProFILE" & spr & "NAME"
        If iType <> "Critter" Then
            For n = 0 To UBound(Items_LST)
                If Items_LST(n, 1) = iType Then
                    x += 1
                    ReDim Preserve Table(x)
                    Table(x) = Items_LST(n, 0)
                End If
            Next
            '
            Dim Read As Boolean = False
            GetMsgData("pro_item.msg")
            Dim fFile As Byte
            For n = 1 To UBound(Table)
                Current_Path = Check_File("proto\items\" & Table(n))
                fFile = FreeFile()
                FileOpen(fFile, Current_Path & "\proto\items\" & Table(n), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
                FileGet(fFile, CommonItem)
                If Current_Path = SaveMOD_Path Then
                    Table(n) = spr & Table(n)
                Else
                    Table(n) = "#" & spr & Table(n)
                End If
                Table(n) &= spr & GetNameCritter(ReverseBytes(CommonItem.DescID))
                For m = 0 To UBound(CheckedList)
                    If n = 1 Then Table(0) &= spr & CheckedList(m)
                    Select Case iType
                        Case "Weapon"
                            If Not Read Then FileGet(fFile, WeaponItem) : FileClose(fFile) : Read = True
                            CreateTable_Weapon(n, m)
                        Case "Ammo"
                            If Not Read Then FileGet(fFile, AmmoItem) : FileClose(fFile) : Read = True
                            CreateTable_Ammo(n, m)
                        Case "Armor"
                            If Not Read Then FileGet(fFile, ArmorItem) : FileClose(fFile) : Read = True
                            CreateTable_Armor(n, m)
                        Case "Drugs"
                            If Not Read Then FileGet(fFile, DrugItem) : FileClose(fFile) : Read = True
                            CreateTable_Drugs(n, m)
                        Case Else
                            If Not Read Then FileGet(fFile, MiscItem) : FileClose(fFile) : Read = True
                            CreateTable_Misc(n, m)
                    End Select
                Next
                Read = False
            Next
        Else ' "Critter table"
            If Critter_LST Is Nothing Then CreateCritterList()
            ReDim Preserve Table(UBound(Critter_LST) + 1)
            GetMsgData("pro_crit.msg")
            Dim fFile As Byte
            For n = 1 To UBound(Table)
                Current_Path = Check_File("proto\critters\" & Critter_LST(n - 1))
                If My.Computer.FileSystem.GetFileInfo(Current_Path & "\proto\critters\" & Critter_LST(n - 1)).Length < 416 Then GoTo FNext
                If Current_Path = SaveMOD_Path Then
                    Table(n) = spr & Critter_LST(n - 1)
                Else
                    Table(n) = "#" & spr & Critter_LST(n - 1)
                End If
                fFile = FreeFile()
                FileOpen(fFile, Current_Path & "\proto\critters\" & Critter_LST(n - 1), OpenMode.Binary, OpenAccess.Read, OpenShare.LockWrite)
                FileGet(fFile, CritterPro)
                FileClose(fFile)
                Table(n) &= spr & GetNameCritter(ReverseBytes(CritterPro.DescID))
                For m = 0 To UBound(CheckedList)
                    'создаем строку с параметрами
                    If n = 1 Then Table(0) &= spr & CheckedList(m)
                    CreateTable_Critter(n, m)
                Next
FNext:
            Next
            FileClose()
        End If
        '
        SaveFileDialog1.FileName = iType
SaveCancel:
        If SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
        iType = SaveFileDialog1.FileName
SaveRetry:
        On Error GoTo SError
        IO.File.WriteAllLines(iType, Table, System.Text.ASCIIEncoding.Default) '& ".csv"
        If MsgBox("Open saved table file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Process.Start(SaveFileDialog1.FileName)
        Exit Sub
SError:
        On Error GoTo -1
        If MsgBox("Error save table file!", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then GoTo SaveRetry
        GoTo SaveCancel
    End Sub

    Private Sub CreateTable_Weapon(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Cost"
                Table(n) &= spr & ReverseBytes(CommonItem.Cost)
            Case "Weight"
                Table(n) &= spr & ReverseBytes(CommonItem.Weight)
            Case "Min Strength"
                Table(n) &= spr & ReverseBytes(WeaponItem.MinST)
            Case "Damage Type"
                Table(n) &= spr & DmgType(ReverseBytes(WeaponItem.DmgType))
            Case "Min Damage"
                Table(n) &= spr & ReverseBytes(WeaponItem.MinDmg)
            Case "Max Damage"
                Table(n) &= spr & ReverseBytes(WeaponItem.MaxDmg)
            Case "Range Primary Attack"
                Table(n) &= spr & ReverseBytes(WeaponItem.MaxRangeP)
            Case "Range Secondary Attack"
                Table(n) &= spr & ReverseBytes(WeaponItem.MaxRangeS)
            Case "AP Cost Primary Attack"
                Table(n) &= spr & ReverseBytes(WeaponItem.MPCostP)
            Case "AP Cost Secondary Attack"
                Table(n) &= spr & ReverseBytes(WeaponItem.MPCostS)
            Case "Max Ammo"
                Table(n) &= spr & ReverseBytes(WeaponItem.MaxAmmo)
            Case "Rounds Brust"
                Table(n) &= spr & ReverseBytes(WeaponItem.Rounds)
            Case "Caliber"
                If WeaponItem.Caliber <> &HFFFFFFFF Then
                    Table(n) &= spr & CaliberNAME(ReverseBytes(WeaponItem.Caliber))
                Else : Table(n) &= spr : End If
            Case "Ammo PID"
                If WeaponItem.AmmoPID <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_NAME(ReverseBytes(WeaponItem.AmmoPID) - 1) & " [" & ReverseBytes(WeaponItem.AmmoPID) & "]"
                Else : Table(n) &= spr : End If
            Case "Critical Fail"
                Table(n) &= spr & ReverseBytes(WeaponItem.CritFail)
            Case "Perk"
                If WeaponItem.Perk <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(ReverseBytes(WeaponItem.Perk)) & " [" & ReverseBytes(WeaponItem.Perk) & "]"
                Else : Table(n) &= spr : End If
            Case "Size"
                Table(n) &= spr & ReverseBytes(CommonItem.Size)
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Ammo(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Dam Div"
                Table(n) &= spr & ReverseBytes(AmmoItem.DamDiv)
            Case "Dam Mult"
                Table(n) &= spr & ReverseBytes(AmmoItem.DamMult)
            Case "AC Adjust"
                Table(n) &= spr & ReverseBytes(AmmoItem.ACAdjust)
            Case "DR Adjust"
                Table(n) &= spr & ReverseBytes(AmmoItem.DRAdjust)
            Case "Quantity"
                Table(n) &= spr & ReverseBytes(AmmoItem.Quantity)
            Case "Caliber"
                If AmmoItem.Caliber <> &HFFFFFFFF Then
                    Table(n) &= spr & CaliberNAME(ReverseBytes(AmmoItem.Caliber))
                Else : Table(n) &= spr : End If
            Case "Cost"
                Table(n) &= spr & ReverseBytes(CommonItem.Cost)
            Case "Weight"
                Table(n) &= spr & ReverseBytes(CommonItem.Weight)
            Case "Size"
                Table(n) &= spr & ReverseBytes(CommonItem.Size)
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Armor(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Cost"
                Table(n) &= spr & ReverseBytes(CommonItem.Cost)
            Case "Weight"
                Table(n) &= spr & ReverseBytes(CommonItem.Weight)
            Case "Armor Class"
                Table(n) &= spr & ReverseBytes(ArmorItem.AC)
            Case "Normal DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTNormal) & "|" & ReverseBytes(ArmorItem.DRNormal)
            Case "Laser DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTLaser) & "|" & ReverseBytes(ArmorItem.DRLaser)
            Case "Fire DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTFire) & "|" & ReverseBytes(ArmorItem.DRFire)
            Case "Plasma DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTPlasma) & "|" & ReverseBytes(ArmorItem.DRPlasma)
            Case "Electrical DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTElectrical) & "|" & ReverseBytes(ArmorItem.DRElectrical)
            Case "EMP DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTEMP) & "|" & ReverseBytes(ArmorItem.DREMP)
            Case "Explosion DT|DR"
                Table(n) &= spr & ReverseBytes(ArmorItem.DTExplode) & "|" & ReverseBytes(ArmorItem.DRExplode)
            Case "Perk"
                If ArmorItem.Perk <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(ReverseBytes(ArmorItem.Perk)) & " [" & ReverseBytes(ArmorItem.Perk) & "]"
                Else : Table(n) &= spr : End If
            Case "Size"
                Table(n) &= spr & ReverseBytes(CommonItem.Size)
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Drugs(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Cost"
                Table(n) &= spr & ReverseBytes(CommonItem.Cost)
            Case "Weight"
                Table(n) &= spr & ReverseBytes(CommonItem.Weight)
            Case "Modify Stat 0"
                If DrugItem.Stat0 <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_Form.ComboBox19.Items(2 + (ReverseBytes(DrugItem.Stat0))).ToString & " [" & ReverseBytes(DrugItem.Stat0) & "]"
                Else : Table(n) &= spr : End If
            Case "Modify Stat 1"
                If DrugItem.Stat1 <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_Form.ComboBox19.Items(2 + (ReverseBytes(DrugItem.Stat1))).ToString & " [" & ReverseBytes(DrugItem.Stat1) & "]"
                Else : Table(n) &= spr : End If
            Case "Modify Stat 2"
                If DrugItem.Stat2 <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_Form.ComboBox19.Items(2 + (ReverseBytes(DrugItem.Stat2))).ToString & " [" & ReverseBytes(DrugItem.Stat2) & "]"
                Else : Table(n) &= spr : End If
            Case "Instant Amount 0"
                Table(n) &= spr & ReverseBytes(DrugItem.iAmount0)
            Case "Instant Amount 1"
                Table(n) &= spr & ReverseBytes(DrugItem.iAmount1)
            Case "Instant Amount 2"
                Table(n) &= spr & ReverseBytes(DrugItem.iAmount2)
            Case "First Amount 0"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount0)
            Case "First Amount 1"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount1)
            Case "First Amount 2"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount2)
            Case "First Duration Time"
                Table(n) &= spr & ReverseBytes(DrugItem.Duration1)
            Case "Second Amount 0"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount0)
            Case "Second Amount 1"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount1)
            Case "Second Amount 2"
                Table(n) &= spr & ReverseBytes(DrugItem.fAmount2)
            Case "Second Duration Time"
                Table(n) &= spr & ReverseBytes(DrugItem.Duration2)
            Case "Addiction Effect"
                If DrugItem.W_Effect <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(ReverseBytes(DrugItem.W_Effect)) & " [" & ReverseBytes(DrugItem.W_Effect) & "]"
                Else : Table(n) &= spr : End If
            Case "Addiction Onset Time"
                Table(n) &= spr & ReverseBytes(DrugItem.W_Onset)
            Case "Addiction Rate"
                Table(n) &= spr & ReverseBytes(DrugItem.AddictionRate)
            Case "Size"
                Table(n) &= spr & ReverseBytes(CommonItem.Size)
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Misc(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Power PID"
                If MiscItem.PowerPID <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_NAME(ReverseBytes(MiscItem.PowerPID) - 1) & " [" & ReverseBytes(MiscItem.PowerPID) & "]"
                Else : Table(n) &= spr : End If
            Case "Power Type"
                If ReverseBytes(MiscItem.PowerType) < UBound(CaliberNAME) Then
                    Table(n) &= spr & CaliberNAME(ReverseBytes(MiscItem.PowerType))
                Else : Table(n) &= spr : End If
            Case "Charges"
                Table(n) &= spr & ReverseBytes(MiscItem.Charges)
            Case "Cost"
                Table(n) &= spr & ReverseBytes(CommonItem.Cost)
            Case "Weight"
                Table(n) &= spr & ReverseBytes(CommonItem.Weight)
            Case "Size"
                Table(n) &= spr & ReverseBytes(CommonItem.Size)
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(ReverseBytes(CommonItem.Falgs) And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Critter(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList(m)
            Case "Strength"
                Table(n) &= spr & ReverseBytes(CritterPro.Strength)
            Case "Perception"
                Table(n) &= spr & ReverseBytes(CritterPro.Perception)
            Case "Endurance"
                Table(n) &= spr & ReverseBytes(CritterPro.Endurance)
            Case "Charisma"
                Table(n) &= spr & ReverseBytes(CritterPro.Charisma)
            Case "Intelligence"
                Table(n) &= spr & ReverseBytes(CritterPro.Intelligence)
            Case "Agility"
                Table(n) &= spr & ReverseBytes(CritterPro.Agility)
            Case "Luck"
                Table(n) &= spr & ReverseBytes(CritterPro.Luck)
                '
            Case "Health Point"
                Table(n) &= spr & (ReverseBytes(CritterPro.HP) + ReverseBytes(CritterPro.b_HP))
            Case "Action Point"
                Table(n) &= spr & (ReverseBytes(CritterPro.AP) + ReverseBytes(CritterPro.b_AP))
            Case "Armor Class"
                Table(n) &= spr & (ReverseBytes(CritterPro.AC) + ReverseBytes(CritterPro.b_AC))
            Case "Melee Damage"
                Table(n) &= spr & (ReverseBytes(CritterPro.MeleeDmg) + ReverseBytes(CritterPro.b_MeleeDmg))
            Case "Damage Type"
                Table(n) &= spr & DmgType(ReverseBytes(CritterPro.DamageType))
            Case "Critical Chance"
                Table(n) &= spr & (ReverseBytes(CritterPro.Critical) + ReverseBytes(CritterPro.b_Critical))
            Case "Sequence"
                Table(n) &= spr & (ReverseBytes(CritterPro.Sequence) + ReverseBytes(CritterPro.b_Sequence))
            Case "Healing Rate"
                Table(n) &= spr & (ReverseBytes(CritterPro.Healing) + ReverseBytes(CritterPro.b_Healing))
            Case "Exp Value"
                Table(n) &= spr & ReverseBytes(CritterPro.ExpVal)
                '
            Case "Small Guns [Skill]"
                Table(n) &= spr & CStr(c_SmallGun_Skill() + ReverseBytes(CritterPro.SmallGuns))
            Case "Big Guns [Skill]"
                Table(n) &= spr & CStr(c_BigEnergyGun_Skill() + ReverseBytes(CritterPro.BigGuns))
            Case "Energy Weapons [Skill]"
                Table(n) &= spr & CStr(c_BigEnergyGun_Skill() + ReverseBytes(CritterPro.EnergyGun))
            Case "Unarmed [Skill]"
                Table(n) &= spr & CStr(c_Unarmed_Skill() + ReverseBytes(CritterPro.Unarmed))
            Case "Melee [Skill]"
                Table(n) &= spr & CStr(c_Melee_Skill() + ReverseBytes(CritterPro.Melee))
            Case "Throwing [Skill]"
                Table(n) &= spr & CStr(c_Throwing_Skill() + ReverseBytes(CritterPro.Throwing))
                '
            Case "Resistance Radiation"
                Table(n) &= spr & (ReverseBytes(CritterPro.DRRadiation) + ReverseBytes(CritterPro.b_DRRadiation))
            Case "Resistance Poison"
                Table(n) &= spr & (ReverseBytes(CritterPro.DRPoison) + ReverseBytes(CritterPro.b_DRPoison))
                '
            Case "Normal DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTNormal) & "|" & ReverseBytes(CritterPro.b_DRNormal)
            Case "Laser DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTLaser) & "|" & ReverseBytes(CritterPro.b_DRLaser)
            Case "Fire DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTFire) & "|" & ReverseBytes(CritterPro.b_DRFire)
            Case "Plasma DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTPlasma) & "|" & ReverseBytes(CritterPro.b_DRPlasma)
            Case "Electrical DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTElectrical) & "|" & ReverseBytes(CritterPro.b_DRElectrical)
            Case "EMP DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTEMP) & "|" & ReverseBytes(CritterPro.b_DREMP)
            Case "Explosion DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.b_DTExplode) & "|" & ReverseBytes(CritterPro.b_DRExplode)
                '
            Case "Base Normal DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTNormal) & "|" & ReverseBytes(CritterPro.DRNormal)
            Case "Base Laser DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTLaser) & "|" & ReverseBytes(CritterPro.DRLaser)
            Case "Base Fire DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTFire) & "|" & ReverseBytes(CritterPro.DRFire)
            Case "Base Plasma DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTPlasma) & "|" & ReverseBytes(CritterPro.DRPlasma)
            Case "Base Electrical DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTElectrical) & "|" & ReverseBytes(CritterPro.DRElectrical)
            Case "Base EMP DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTEMP) & "|" & ReverseBytes(CritterPro.DREMP)
            Case "Base Explosion DT|DR"
                Table(n) &= spr & ReverseBytes(CritterPro.DTExplode) & "|" & ReverseBytes(CritterPro.DRExplode)
        End Select
    End Sub

    Friend Sub Import_Table(ByVal file As String)
        Dim n As Integer, m As Byte, z As UShort, t_Path As String, ProFile As String
        Dim CommonItem As CmItemPro, WeaponItem As WpItemPro, ArmorItem As ArItemPro
        Dim AmmoItem As AmItemPro, DrugItem As DgItemPro, MiscItem As McItemPro
        Dim Table_File() As String = IO.File.ReadAllLines(file, System.Text.Encoding.Default)
        Dim Table_Param() As String = Split(Table_File(0), ";", -1, CompareMethod.Binary)
        Dim Table_Value(UBound(Table_File) - 1, UBound(Table_Param)) As String
        Dim strSplit() As String
        'разделить
        For n = 1 To UBound(Table_File)
            Dim tLine() As String = Split(Table_File(n), ";", -1, CompareMethod.Binary)
            If tLine(0) <> Nothing OrElse tLine.Length < Table_Param.Length Then
                If tLine(0) <> Nothing Then
                    TableLog_Form.ListBox1.Items.Add("Skip: Table Line " & (n + 1) & " - Used ""#"" symbol in begin line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning: Table Line " & (n + 1) & " - Error count value param.")
                End If
                GoTo lNext
            End If
            For m = 0 To UBound(Table_Param)
                Table_Value(n - 1, m) = tLine(m)
            Next
lNext:
        Next
        'открыть профайл
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\items") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\items")
        Dim fFile As Byte = FreeFile()
        On Error GoTo tError
        For n = 0 To UBound(Table_Value)
            ProFile = Table_Value(n, 1)
            If ProFile = Nothing Then GoTo pNext
            If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\items\" & ProFile) = False Then
                t_Path = Check_File("proto\items\" & ProFile)
                My.Computer.FileSystem.CopyFile(t_Path & "\proto\items\" & ProFile, SaveMOD_Path & "\proto\items\" & ProFile, FileIO.UIOption.AllDialogs)
            End If
            IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & ProFile, IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            FileOpen(fFile, SaveMOD_Path & "\proto\items\" & ProFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.Shared)
            FileGet(fFile, CommonItem)
            Dim fInfo As IO.FileInfo = New IO.FileInfo(SaveMOD_Path & "\proto\items\" & ProFile)
            Select Case fInfo.Length
                Case 69 'Misc
                    FileGet(fFile, MiscItem)
                Case 122 'Оружие 
                    FileGet(fFile, WeaponItem)
                Case 125 'Наркотик
                    FileGet(fFile, DrugItem)
                Case 81 'Патрон
                    FileGet(fFile, AmmoItem)
                Case 129 'Броня
                    FileGet(fFile, ArmorItem)
            End Select
            'изменить значения 
            For m = 3 To UBound(Table_Param)
                Select Case Table_Param(m)
                    'Common
                    Case "Cost"
                        CommonItem.Cost = ReverseBytes(Table_Value(n, m))
                    Case "Weight"
                        CommonItem.Weight = ReverseBytes(Table_Value(n, m))
                    Case "Size"
                        CommonItem.Size = ReverseBytes(Table_Value(n, m))
                    Case "Shoot Thru [Flag]"
                        If Table_Value(n, m) = "True" Then CommonItem.Falgs = CommonItem.Falgs Or ReverseBytes(&H80000000) Else CommonItem.Falgs = CommonItem.Falgs And ReverseBytes(Not (&H80000000))
                    Case "Light Thru [Flag]"
                        If Table_Value(n, m) = "True" Then CommonItem.Falgs = CommonItem.Falgs Or ReverseBytes(&H20000000) Else CommonItem.Falgs = CommonItem.Falgs And ReverseBytes(Not (&H20000000))
                        'weapon
                    Case "Min Strength"
                        WeaponItem.MinST = ReverseBytes(Table_Value(n, m))
                    Case "Damage Type"
                        For z = 0 To UBound(DmgType)
                            If Table_Value(n, m) = DmgType(z) Then WeaponItem.DmgType = ReverseBytes(z) : Exit For
                        Next
                    Case "Min Damage"
                        WeaponItem.MinDmg = ReverseBytes(Table_Value(n, m))
                    Case "Max Damage"
                        WeaponItem.MaxDmg = ReverseBytes(Table_Value(n, m))
                    Case "Range Primary Attack"
                        WeaponItem.MaxRangeP = ReverseBytes(Table_Value(n, m))
                    Case "Range Secondary Attack"
                        WeaponItem.MaxRangeS = ReverseBytes(Table_Value(n, m))
                    Case "AP Cost Primary Attack"
                        WeaponItem.MPCostP = ReverseBytes(Table_Value(n, m))
                    Case "AP Cost Secondary Attack"
                        WeaponItem.MPCostS = ReverseBytes(Table_Value(n, m))
                    Case "Max Ammo"
                        WeaponItem.MaxAmmo = ReverseBytes(Table_Value(n, m))
                    Case "Rounds Brust"
                        WeaponItem.Rounds = ReverseBytes(Table_Value(n, m))
                    Case "Caliber"
                        If fInfo.Length = 122 Then 'weapon
                            If Table_Value(n, m) <> Nothing Then
                                For z = 0 To UBound(CaliberNAME)
                                    If Table_Value(n, m) = CaliberNAME(z) Then WeaponItem.Caliber = ReverseBytes(z) : Exit For
                                Next
                            End If
                        Else 'ammo
                            If Table_Value(n, m) <> Nothing Then
                                For z = 0 To UBound(CaliberNAME)
                                    If Table_Value(n, m) = CaliberNAME(z) Then AmmoItem.Caliber = ReverseBytes(z) : Exit For
                                Next
                            End If
                        End If
                    Case "Ammo PID"
                        WeaponItem.AmmoPID = GetTable_Param(Table_Value(n, m))
                    Case "Critical Fail"
                        WeaponItem.CritFail = ReverseBytes(Table_Value(n, m))
                    Case "Perk"
                        If fInfo.Length = 122 Then 'weapon
                            WeaponItem.Perk = GetTable_Param(Table_Value(n, m))
                        Else 'armor
                            ArmorItem.Perk = GetTable_Param(Table_Value(n, m))
                        End If
                        'Ammo
                    Case "Dam Div"
                        AmmoItem.DamDiv = ReverseBytes(Table_Value(n, m))
                    Case "Dam Mult"
                        AmmoItem.DamMult = ReverseBytes(Table_Value(n, m))
                    Case "AC Adjust"
                        AmmoItem.ACAdjust = ReverseBytes(Table_Value(n, m))
                    Case "DR Adjust"
                        AmmoItem.DRAdjust = ReverseBytes(Table_Value(n, m))
                    Case "Quantity"
                        AmmoItem.Quantity = ReverseBytes(Table_Value(n, m))
                        'Armor
                    Case "Armor Class"
                        ArmorItem.AC = ReverseBytes(Table_Value(n, m))
                    Case "Normal DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTNormal = ReverseBytes(strSplit(0))
                        ArmorItem.DRNormal = ReverseBytes(strSplit(1))
                    Case "Laser DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTLaser = ReverseBytes(strSplit(0))
                        ArmorItem.DRLaser = ReverseBytes(strSplit(1))
                    Case "Fire DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTFire = ReverseBytes(strSplit(0))
                        ArmorItem.DRFire = ReverseBytes(strSplit(1))
                    Case "Plasma DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTPlasma = ReverseBytes(strSplit(0))
                        ArmorItem.DRPlasma = ReverseBytes(strSplit(1))
                    Case "Electrical DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTElectrical = ReverseBytes(strSplit(0))
                        ArmorItem.DRElectrical = ReverseBytes(strSplit(1))
                    Case "EMP DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTEMP = ReverseBytes(strSplit(0))
                        ArmorItem.DREMP = ReverseBytes(strSplit(1))
                    Case "Explosion DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        ArmorItem.DTExplode = ReverseBytes(strSplit(0))
                        ArmorItem.DRExplode = ReverseBytes(strSplit(1))
                        'Drug
                    Case "Modify Stat 0"
                        DrugItem.Stat0 = GetTable_Param(Table_Value(n, m))
                    Case "Modify Stat 1"
                        DrugItem.Stat1 = GetTable_Param(Table_Value(n, m))
                    Case "Modify Stat 2"
                        DrugItem.Stat2 = GetTable_Param(Table_Value(n, m))
                    Case "Instant Amount 0"
                        DrugItem.iAmount0 = ReverseBytes(Table_Value(n, m))
                    Case "Instant Amount 1"
                        DrugItem.iAmount1 = ReverseBytes(Table_Value(n, m))
                    Case "Instant Amount 2"
                        DrugItem.iAmount2 = ReverseBytes(Table_Value(n, m))
                    Case "First Amount 0"
                        DrugItem.fAmount0 = ReverseBytes(Table_Value(n, m))
                    Case "First Amount 1"
                        DrugItem.fAmount1 = ReverseBytes(Table_Value(n, m))
                    Case "First Amount 2"
                        DrugItem.fAmount2 = ReverseBytes(Table_Value(n, m))
                    Case "First Duration Time"
                        DrugItem.Duration1 = ReverseBytes(Table_Value(n, m))
                    Case "Second Amount 0"
                        DrugItem.fAmount0 = ReverseBytes(Table_Value(n, m))
                    Case "Second Amount 1"
                        DrugItem.fAmount1 = ReverseBytes(Table_Value(n, m))
                    Case "Second Amount 2"
                        DrugItem.fAmount2 = ReverseBytes(Table_Value(n, m))
                    Case "Second Duration Time"
                        DrugItem.Duration2 = ReverseBytes(Table_Value(n, m))
                    Case "Addiction Effect"
                        DrugItem.W_Effect = GetTable_Param(Table_Value(n, m))
                    Case "Addiction Onset Time"
                        DrugItem.W_Onset = ReverseBytes(Table_Value(n, m))
                    Case "Addiction Rate"
                        DrugItem.AddictionRate = ReverseBytes(Table_Value(n, m))
                        'Misc
                    Case "Power PID"
                        MiscItem.PowerPID = GetTable_Param(Table_Value(n, m))
                    Case "Power Type"
                        If Table_Value(n, m) <> Nothing Then
                            For z = 0 To UBound(CaliberNAME)
                                If Table_Value(n, m) = CaliberNAME(z) Then MiscItem.PowerType = ReverseBytes(z) : Exit For
                            Next
                        End If
                    Case "Charges"
                        If Table_Value(n, m) <> Nothing Then MiscItem.Charges = ReverseBytes(Table_Value(n, m))
                End Select
            Next
            'сохранить профайл, и перейти к следующему профайлу
            FilePut(fFile, CommonItem, 1)
            Select Case fInfo.Length
                Case 69 'Misc
                    FilePut(fFile, MiscItem)
                Case 122 'Оружие 
                    FilePut(fFile, WeaponItem)
                Case 125 'Наркотик
                    FilePut(fFile, DrugItem)
                Case 81 'Патрон
                    FilePut(fFile, AmmoItem)
                Case 129 'Броня
                    FilePut(fFile, ArmorItem)
            End Select
            FileClose(fFile)
            If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & ProFile, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
pNext:
        Next
        On Error GoTo 0
        If TableLog_Form.ListBox1.Items.Count > 0 Then TableLog_Form.Show()
        MsgBox("Done!")
        Exit Sub
tError:
        FileClose(fFile)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\items\" & ProFile, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        MsgBox("Param: " & UCase(Table_Param(m)) & " PRO Line: " & Table_Value(n, 0), MsgBoxStyle.Critical, "Error Import")
    End Sub

    Private Function GetTable_Param(ByRef tParam As String) As Integer
        If tParam <> Nothing Then
            Dim y As Byte = InStr(tParam, "[", CompareMethod.Binary)
            Return ReverseBytes(tParam.Substring(y, tParam.Length - (y + 1)))
        End If
        Return &HFFFFFFFF
    End Function

    Friend Sub Import_CR_Table(ByVal file As String)
        Dim z As UShort, n As Integer, m As Integer
        Dim ProFile, t_Path As String 'CritterPro As CritPro,
        Dim strSplit() As String

        Dim Table_File() As String = IO.File.ReadAllLines(file, System.Text.Encoding.Default)
        Dim Table_Param() As String = Split(Table_File(0), ";", -1, CompareMethod.Binary)
        Dim Table_Value(UBound(Table_File) - 1, UBound(Table_Param)) As String
        'разделить
        For n = 1 To UBound(Table_File)
            Dim tLine() As String = Split(Table_File(n), ";", -1, CompareMethod.Binary)
            If tLine(0) <> Nothing OrElse tLine.Length < Table_Param.Length Then
                If tLine(0) <> Nothing Then
                    TableLog_Form.ListBox1.Items.Add("Skip: Table Line " & (n + 1) & " - Used ""#"" symbol in begin line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning: Table Line " & (n + 1) & " - Error count value param.")
                End If
                GoTo lNext
            End If
            For m = 0 To UBound(Table_Param)
                Table_Value(n - 1, m) = tLine(m)
            Next
lNext:
        Next
        'открыть профайл
        If My.Computer.FileSystem.DirectoryExists(SaveMOD_Path & "\proto\critters") = False Then My.Computer.FileSystem.CreateDirectory(SaveMOD_Path & "\proto\critters")
        Dim fFile As Byte = FreeFile()
        On Error GoTo tError
        For n = 0 To UBound(Table_Value)
            ProFile = Table_Value(n, 1)
            If ProFile = Nothing Then GoTo pNext
            If My.Computer.FileSystem.FileExists(SaveMOD_Path & "\proto\critters\" & ProFile) = False Then
                t_Path = Check_File("proto\critters\" & ProFile)
                My.Computer.FileSystem.CopyFile(t_Path & "\proto\critters\" & ProFile, SaveMOD_Path & "\proto\critters\" & ProFile, FileIO.UIOption.AllDialogs)
            End If
            IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & ProFile, IO.FileAttributes.Normal Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
            FileOpen(fFile, SaveMOD_Path & "\proto\critters\" & ProFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.Shared)
            FileGet(fFile, CritterPro)
            'изменить значения
            'Common pass 1
            For m = 3 To UBound(Table_Param)
                Select Case Table_Param(m)
                    Case "Strength"
                        CritterPro.Strength = ReverseBytes(Table_Value(n, m))
                    Case "Perception"
                        CritterPro.Perception = ReverseBytes(Table_Value(n, m))
                    Case "Endurance"
                        CritterPro.Endurance = ReverseBytes(Table_Value(n, m))
                    Case "Charisma"
                        CritterPro.Charisma = ReverseBytes(Table_Value(n, m))
                    Case "Intelligence"
                        CritterPro.Intelligence = ReverseBytes(Table_Value(n, m))
                    Case "Agility"
                        CritterPro.Agility = ReverseBytes(Table_Value(n, m))
                    Case "Luck"
                        CritterPro.Luck = ReverseBytes(Table_Value(n, m))
                    Case "Exp Value"
                        CritterPro.ExpVal = ReverseBytes(Table_Value(n, m))
                    Case "Damage Type"
                        For z = 0 To UBound(DmgType)
                            If Table_Value(n, m) = DmgType(z) Then CritterPro.DamageType = ReverseBytes(z) : Exit For
                        Next
                        'Armor
                    Case "Normal DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTNormal = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRNormal = ReverseBytes(CInt(strSplit(1)))
                    Case "Laser DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTLaser = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRLaser = ReverseBytes(CInt(strSplit(1)))
                    Case "Fire DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTFire = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRFire = ReverseBytes(CInt(strSplit(1)))
                    Case "Plasma DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTPlasma = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRPlasma = ReverseBytes(CInt(strSplit(1)))
                    Case "Electrical DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTElectrical = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRElectrical = ReverseBytes(CInt(strSplit(1)))
                    Case "EMP DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTEMP = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DREMP = ReverseBytes(CInt(strSplit(1)))
                    Case "Explosion DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.b_DTExplode = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.b_DRExplode = ReverseBytes(CInt(strSplit(1)))
                        '
                    Case "Base Normal DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTNormal = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRNormal = ReverseBytes(CInt(strSplit(1)))
                    Case "Base Laser DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTLaser = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRLaser = ReverseBytes(CInt(strSplit(1)))
                    Case "Base Fire DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTFire = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRFire = ReverseBytes(CInt(strSplit(1)))
                    Case "Base Plasma DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTPlasma = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRPlasma = ReverseBytes(CInt(strSplit(1)))
                    Case "Base Electrical DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTElectrical = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRElectrical = ReverseBytes(CInt(strSplit(1)))
                    Case "Base EMP DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTEMP = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DREMP = ReverseBytes(CInt(strSplit(1)))
                    Case "Base Explosion DT|DR"
                        strSplit = Table_Value(n, m).Split("|")
                        CritterPro.DTExplode = ReverseBytes(CInt(strSplit(0)))
                        CritterPro.DRExplode = ReverseBytes(CInt(strSplit(1)))
                End Select
            Next
            'Calculate pass 2
            For m = 3 To UBound(Table_Param)
                Select Case Table_Param(m)
                    Case "Action Point"
                        z = c_Action_Point()
                        CritterPro.AP = ReverseBytes(z)
                        CritterPro.b_AP = ReverseBytes(Table_Value(n, m) - z)
                    Case "Armor Class"
                        CritterPro.AC = CritterPro.Agility
                        CritterPro.b_AC = ReverseBytes(Table_Value(n, m) - ReverseBytes(CritterPro.Agility))
                    Case "Health Point"
                        z = c_Health_Point()
                        CritterPro.HP = ReverseBytes(z)
                        CritterPro.b_HP = ReverseBytes(Table_Value(n, m) - z)
                    Case "Healing Rate"
                        z = c_Healing_Rate()
                        CritterPro.Healing = ReverseBytes(z)
                        CritterPro.b_Healing = ReverseBytes(Table_Value(n, m) - z)
                    Case "Melee Damage"
                        z = c_Melee_Damage()
                        CritterPro.MeleeDmg = ReverseBytes(z)
                        CritterPro.b_MeleeDmg = ReverseBytes(Table_Value(n, m) - z)
                    Case "Critical Chance"
                        CritterPro.Critical = CritterPro.Luck
                        CritterPro.b_Critical = ReverseBytes(Table_Value(n, m) - ReverseBytes(CritterPro.Luck))
                    Case "Sequence"
                        z = c_Sequence()
                        CritterPro.Sequence = ReverseBytes(z)
                        CritterPro.b_Sequence = ReverseBytes(Table_Value(n, m) - z)
                    Case "Resistance Radiation"
                        z = c_Radiation()
                        CritterPro.DRRadiation = ReverseBytes(z)
                        CritterPro.b_DRRadiation = ReverseBytes(Table_Value(n, m) - z)
                    Case "Resistance Poison"
                        z = c_Poison()
                        CritterPro.DRPoison = ReverseBytes(z)
                        CritterPro.b_DRPoison = ReverseBytes(Table_Value(n, m) - z)
                        'Skill
                    Case "Small Guns [Skill]"
                        CritterPro.SmallGuns = ReverseBytes(Table_Value(n, m) - c_SmallGun_Skill())
                    Case "Big Guns [Skill]"
                        CritterPro.BigGuns = ReverseBytes(Table_Value(n, m) - c_BigEnergyGun_Skill())
                    Case "Energy Weapons [Skill]"
                        CritterPro.EnergyGun = ReverseBytes(Table_Value(n, m) - c_BigEnergyGun_Skill())
                    Case "Unarmed [Skill]"
                        CritterPro.Unarmed = ReverseBytes(Table_Value(n, m) - c_Unarmed_Skill())
                    Case "Melee [Skill]"
                        CritterPro.Melee = ReverseBytes(Table_Value(n, m) - c_Melee_Skill())
                    Case "Throwing [Skill]"
                        CritterPro.Throwing = ReverseBytes(Table_Value(n, m) - c_Throwing_Skill())
                End Select
            Next
            'сохранить профайл, и перейти к следующему профайлу
            FilePut(fFile, CritterPro, 1)
            FileClose(fFile)
            If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & ProFile, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
pNext:
        Next
        On Error GoTo 0
        If TableLog_Form.ListBox1.Items.Count > 0 Then TableLog_Form.Show()
        MsgBox("Done!")
        Exit Sub
tError:
        FileClose(fFile)
        If proRO Then IO.File.SetAttributes(SaveMOD_Path & "\proto\critters\" & ProFile, IO.FileAttributes.ReadOnly Or IO.FileAttributes.Archive Or IO.FileAttributes.NotContentIndexed)
        MsgBox("Param: " & UCase(Table_Param(m)) & " PRO Line: " & Table_Value(n, 0), MsgBoxStyle.Critical, "Error Import")
    End Sub


End Class