Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports Prototypes

Public Class Table_Form

    Private Enum TabType As Integer
        Critter
        Weapon
        Ammo
        Armor
        Drugs
        Misc
    End Enum

    Private CheckedList As CheckedListBox.CheckedItemCollection
    Private Table As List(Of String) = New List(Of String)()

    Private Const spr As String = ";"
    Private Const splt As Char = "|"c

    Private ReadOnly DmgType() As String = {"Normal", "Laser", "Fire", "Plasma", "Electrical", "EMP", "Explode"}

    Private ReadOnly DrugEffect() As String = {"Drug Stat (Special)", "None", "Strength", "Perception", "Endurance",
                                               "Charisma", "Intelligence", "Agility", "Luck", "Max.Healing Point",
                                               "Max.Action Point", "Calss Armor", "Unarmed Damage", "Melee Damage",
                                               "Max.Weight", "Sequence", "Healing Rate", "Critical Chance",
                                               "Better Critical", "Normal Tresholds Damage", "Laser Tresholds Damage",
                                               "Fire Tresholds Damage", "Plasma Tresholds Damage",
                                               "Electrical Tresholds Damage", "EMP Tresholds Damage",
                                               "Explode Tresholds Damage", "Normal Damage Resistance",
                                               "Laser Damage Resistance", "Fire Damage Resistance",
                                               "Plasma Damage Resistance", "Electrical Damage Resistance",
                                               "EMP Damage Resistance", "Explode Damage Resistance",
                                               "Radiation Resistance", "Poison Resistance", "Age", "Gender", "Current HP",
                                               "Current Poison Level", "Current Radiation Level"}

    Private CritterPro As CritPro
    Private CommonItem As CmItemPro
    Private WeaponItem As WpItemPro
    Private ArmorItem As ArItemPro
    Private AmmoItem As AmItemPro
    Private DrugItem As DgItemPro
    Private MiscItem As McItemPro

    Private Sub CreateTable()
        Dim fFile As Integer, iType As Integer = TabControl1.SelectedIndex

        Table.Add("Import" & spr & "ProFILE" & spr & "NAME")
        If iType > TabType.Critter Then
            For n = 0 To UBound(Items_LST)
                If Items_LST(n).itemType = Array.IndexOf(ItemTypesName, TabControl1.SelectedTab.Text) Then
                    Table.Add(Items_LST(n).proFile)
                End If
            Next
            '
            Dim Read As Boolean = False
            GetMsgData("pro_item.msg")
            For n = 1 To Table.Count - 1
                Current_Path = DatFiles.CheckFile(PROTO_ITEMS & Table(n))
                Dim cmProDataBuf(Prototypes.ItemCommonLen - 1) As Integer
                fFile = FreeFile()
                FileOpen(fFile, Current_Path & PROTO_ITEMS & Table(n), OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
                FileGet(fFile, cmProDataBuf)
                ProFiles.ReverseLoadData(cmProDataBuf, CommonItem)
                FileGet(fFile, CommonItem.SoundID)
                If Current_Path = SaveMOD_Path Then
                    Table(n) = spr & Table(n)
                Else
                    Table(n) = "#" & spr & Table(n) ' # - ignore mark
                End If
                Table(n) &= spr & Messages.GetNameObject(CommonItem.DescID) ' get name
                For m = 0 To CheckedList.Count - 1
                    If n = 1 Then Table(0) &= spr & CheckedList.Item(m).ToString
                    Select Case iType
                        Case TabType.Weapon
                            If Not Read Then
                                Dim wnProDataBuf(Prototypes.ItemWeaponLen - 1) As Integer
                                FileGet(fFile, wnProDataBuf)
                                ProFiles.ReverseLoadData(wnProDataBuf, WeaponItem)
                                FileGet(fFile, WeaponItem.wSoundID)
                                FileClose(fFile)
                                Read = True
                            End If
                            CreateTable_Weapon(n, m)
                        Case TabType.Ammo
                            If Not Read Then
                                Dim amProDataBuf(Prototypes.ItemAmmoLen - 1) As Integer
                                FileGet(fFile, amProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(amProDataBuf, AmmoItem)
                                Read = True
                            End If
                            CreateTable_Ammo(n, m)
                        Case TabType.Armor
                            If Not Read Then
                                Dim arProDataBuf(Prototypes.ItemArmorLen - 1) As Integer
                                FileGet(fFile, arProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(arProDataBuf, ArmorItem)
                                Read = True
                            End If
                            CreateTable_Armor(n, m)
                        Case TabType.Drugs
                            If Not Read Then
                                Dim drProDataBuf(Prototypes.ItemDrugsLen - 1) As Integer
                                FileGet(fFile, drProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(drProDataBuf, DrugItem)
                                Read = True
                            End If
                            CreateTable_Drugs(n, m)
                        Case Else ' misc
                            If Not Read Then
                                Dim msProDataBuf(Prototypes.ItemMiscLen - 1) As Integer
                                FileGet(fFile, msProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(msProDataBuf, MiscItem)
                                Read = True
                            End If
                            CreateTable_Misc(n, m)
                    End Select
                Next
                Read = False
            Next
        Else
            ' Critter table
            If Critter_LST Is Nothing Then Main.CreateCritterList()

            Progress_Form.ProgressBar1.Value = 0
            Main.ShowProgressBar(UBound(Critter_LST) + 1)

            GetMsgData("pro_crit.msg")
            For n = 1 To UBound(Critter_LST) + 1
                Dim proFile As String = Critter_LST(n - 1).proFile
                Current_Path = DatFiles.CheckFile(PROTO_CRITTERS & proFile)
                Dim pathFile As String = Current_Path & PROTO_CRITTERS & proFile
                If FileSystem.GetFileInfo(pathFile).Length < 416 Then
                    Table.Add("#" & spr & proFile & spr & "<BadFormat>")
                    'Log 
                    Main_Form.TextBox1.Text = "Bad Format: " & pathFile & vbLf & Main_Form.TextBox1.Text
                    Application.DoEvents()
                    Continue For
                End If
                ProFiles.LoadCritterProData(pathFile, CritterPro)

                If Current_Path.ToLower = SaveMOD_Path.ToLower Then
                    Table.Add(spr & proFile)
                Else
                    Table.Add("#" & spr & proFile) ' # - ignore mark
                End If

                Table(n) &= spr & Messages.GetNameObject(CritterPro.DescID)
                For m = 0 To CheckedList.Count - 1
                    'создаем строку с параметрами
                    If n = 1 Then Table(0) &= spr & CheckedList.Item(m).ToString
                    CreateTable_Critter(n, m)
                Next
                Progress_Form.ProgressBar1.Value = n
            Next
        End If
        SaveTable(TabControl1.SelectedTab.Text)
        Progress_Form.Close()
        Table.Clear()
    End Sub

    Private Sub SaveTable(ByVal fileName As String)
        SaveFileDialog1.FileName = fileName
        If SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
        fileName = SaveFileDialog1.FileName
SaveRetry:
        Try
            File.WriteAllLines(fileName, Table, ASCIIEncoding.Default) '& ".csv"   
        Catch
            If MsgBox("Error save table file!", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then GoTo SaveRetry
            SaveTable(fileName)
        End Try
        If MsgBox("Open saved table file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Process.Start(SaveFileDialog1.FileName)
    End Sub

    Private Sub CreateTable_Weapon(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Cost"
                Table(n) &= spr & CommonItem.Cost
            Case "Weight"
                Table(n) &= spr & CommonItem.Weight
            Case "Min Strength"
                Table(n) &= spr & WeaponItem.MinST
            Case "Damage Type"
                Table(n) &= spr & DmgType(WeaponItem.DmgType)
            Case "Min Damage"
                Table(n) &= spr & WeaponItem.MinDmg
            Case "Max Damage"
                Table(n) &= spr & WeaponItem.MaxDmg
            Case "Range Primary Attack"
                Table(n) &= spr & WeaponItem.MaxRangeP
            Case "Range Secondary Attack"
                Table(n) &= spr & WeaponItem.MaxRangeS
            Case "AP Cost Primary Attack"
                Table(n) &= spr & WeaponItem.MPCostP
            Case "AP Cost Secondary Attack"
                Table(n) &= spr & WeaponItem.MPCostS
            Case "Max Ammo"
                Table(n) &= spr & WeaponItem.MaxAmmo
            Case "Rounds Brust"
                Table(n) &= spr & WeaponItem.Rounds
            Case "Caliber"
                If WeaponItem.Caliber <> &HFFFFFFFF Then
                    Table(n) &= spr & CaliberNAME(WeaponItem.Caliber)
                Else : Table(n) &= spr : End If
            Case "Ammo PID"
                If WeaponItem.AmmoPID <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_LST(WeaponItem.AmmoPID - 1).itemName & " [" & WeaponItem.AmmoPID & "]"
                Else : Table(n) &= spr : End If
            Case "Critical Fail"
                Table(n) &= spr & WeaponItem.CritFail
            Case "Perk"
                If WeaponItem.Perk <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(WeaponItem.Perk) & " [" & WeaponItem.Perk & "]"
                Else : Table(n) &= spr : End If
            Case "Size"
                Table(n) &= spr & CommonItem.Size
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Ammo(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Dam Div"
                Table(n) &= spr & AmmoItem.DamDiv
            Case "Dam Mult"
                Table(n) &= spr & AmmoItem.DamMult
            Case "AC Adjust"
                Table(n) &= spr & AmmoItem.ACAdjust
            Case "DR Adjust"
                Table(n) &= spr & AmmoItem.DRAdjust
            Case "Quantity"
                Table(n) &= spr & AmmoItem.Quantity
            Case "Caliber"
                If AmmoItem.Caliber <> &HFFFFFFFF Then
                    Table(n) &= spr & CaliberNAME(AmmoItem.Caliber)
                Else : Table(n) &= spr : End If
            Case "Cost"
                Table(n) &= spr & CommonItem.Cost
            Case "Weight"
                Table(n) &= spr & CommonItem.Weight
            Case "Size"
                Table(n) &= spr & CommonItem.Size
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Armor(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Cost"
                Table(n) &= spr & CommonItem.Cost
            Case "Weight"
                Table(n) &= spr & CommonItem.Weight
            Case "Armor Class"
                Table(n) &= spr & ArmorItem.AC
            Case "Normal DT|DR"
                Table(n) &= spr & ArmorItem.DTNormal & "|" & ArmorItem.DRNormal
            Case "Laser DT|DR"
                Table(n) &= spr & ArmorItem.DTLaser & "|" & ArmorItem.DRLaser
            Case "Fire DT|DR"
                Table(n) &= spr & ArmorItem.DTFire & "|" & ArmorItem.DRFire
            Case "Plasma DT|DR"
                Table(n) &= spr & ArmorItem.DTPlasma & "|" & ArmorItem.DRPlasma
            Case "Electrical DT|DR"
                Table(n) &= spr & ArmorItem.DTElectrical & "|" & ArmorItem.DRElectrical
            Case "EMP DT|DR"
                Table(n) &= spr & ArmorItem.DTEMP & "|" & ArmorItem.DREMP
            Case "Explosion DT|DR"
                Table(n) &= spr & ArmorItem.DTExplode & "|" & ArmorItem.DRExplode
            Case "Perk"
                If ArmorItem.Perk <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(ArmorItem.Perk) & " [" & ArmorItem.Perk & "]"
                Else : Table(n) &= spr : End If
            Case "Size"
                Table(n) &= spr & CommonItem.Size
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Drugs(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Cost"
                Table(n) &= spr & CommonItem.Cost
            Case "Weight"
                Table(n) &= spr & CommonItem.Weight
            Case "Modify Stat 0"
                If DrugItem.Stat0 <> &HFFFFFFFF Then
                    Table(n) &= spr & DrugEffect(2 + (DrugItem.Stat0)).ToString & " [" & DrugItem.Stat0 & "]"
                Else : Table(n) &= spr : End If
            Case "Modify Stat 1"
                If DrugItem.Stat1 <> &HFFFFFFFF Then
                    Table(n) &= spr & DrugEffect(2 + (DrugItem.Stat1)).ToString & " [" & DrugItem.Stat1 & "]"
                Else : Table(n) &= spr : End If
            Case "Modify Stat 2"
                If DrugItem.Stat2 <> &HFFFFFFFF Then
                    Table(n) &= spr & DrugEffect(2 + (DrugItem.Stat2)).ToString & " [" & DrugItem.Stat2 & "]"
                Else : Table(n) &= spr : End If
            Case "Instant Amount 0"
                Table(n) &= spr & DrugItem.iAmount0
            Case "Instant Amount 1"
                Table(n) &= spr & DrugItem.iAmount1
            Case "Instant Amount 2"
                Table(n) &= spr & DrugItem.iAmount2
            Case "First Amount 0"
                Table(n) &= spr & DrugItem.fAmount0
            Case "First Amount 1"
                Table(n) &= spr & DrugItem.fAmount1
            Case "First Amount 2"
                Table(n) &= spr & DrugItem.fAmount2
            Case "First Duration Time"
                Table(n) &= spr & DrugItem.Duration1
            Case "Second Amount 0"
                Table(n) &= spr & DrugItem.fAmount0
            Case "Second Amount 1"
                Table(n) &= spr & DrugItem.fAmount1
            Case "Second Amount 2"
                Table(n) &= spr & DrugItem.fAmount2
            Case "Second Duration Time"
                Table(n) &= spr & DrugItem.Duration2
            Case "Addiction Effect"
                If DrugItem.W_Effect <> &HFFFFFFFF Then
                    Table(n) &= spr & Perk_NAME(DrugItem.W_Effect) & " [" & DrugItem.W_Effect & "]"
                Else : Table(n) &= spr : End If
            Case "Addiction Onset Time"
                Table(n) &= spr & DrugItem.W_Onset
            Case "Addiction Rate"
                Table(n) &= spr & DrugItem.AddictionRate
            Case "Size"
                Table(n) &= spr & CommonItem.Size
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Misc(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Power PID"
                If MiscItem.PowerPID <> &HFFFFFFFF Then
                    Table(n) &= spr & Items_LST(MiscItem.PowerPID - 1).itemName & " [" & MiscItem.PowerPID & "]"
                Else : Table(n) &= spr : End If
            Case "Power Type"
                If MiscItem.PowerType < UBound(CaliberNAME) Then
                    Table(n) &= spr & CaliberNAME(MiscItem.PowerType)
                Else : Table(n) &= spr : End If
            Case "Charges"
                Table(n) &= spr & MiscItem.Charges
            Case "Cost"
                Table(n) &= spr & CommonItem.Cost
            Case "Weight"
                Table(n) &= spr & CommonItem.Weight
            Case "Size"
                Table(n) &= spr & CommonItem.Size
            Case "Shoot Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H80000000)
            Case "Light Thru [Flag]"
                Table(n) &= spr & CBool(CommonItem.Falgs And &H20000000)
        End Select
    End Sub

    Private Sub CreateTable_Critter(ByRef n As Integer, ByRef m As Byte)
        Select Case CheckedList.Item(m).ToString
            Case "Strength"
                Table(n) &= spr & CritterPro.Strength
            Case "Perception"
                Table(n) &= spr & CritterPro.Perception
            Case "Endurance"
                Table(n) &= spr & CritterPro.Endurance
            Case "Charisma"
                Table(n) &= spr & CritterPro.Charisma
            Case "Intelligence"
                Table(n) &= spr & CritterPro.Intelligence
            Case "Agility"
                Table(n) &= spr & CritterPro.Agility
            Case "Luck"
                Table(n) &= spr & CritterPro.Luck
                '
            Case "Health Point"
                Table(n) &= spr & (CritterPro.HP + CritterPro.b_HP)
            Case "Action Point"
                Table(n) &= spr & (CritterPro.AP + CritterPro.b_AP)
            Case "Armor Class"
                Table(n) &= spr & (CritterPro.AC + CritterPro.b_AC)
            Case "Melee Damage"
                Table(n) &= spr & (CritterPro.MeleeDmg + CritterPro.b_MeleeDmg)
            Case "Damage Type"
                Table(n) &= spr & DmgType(CritterPro.DamageType)
            Case "Critical Chance"
                Table(n) &= spr & (CritterPro.Critical + CritterPro.b_Critical)
            Case "Sequence"
                Table(n) &= spr & (CritterPro.Sequence + CritterPro.b_Sequence)
            Case "Healing Rate"
                Table(n) &= spr & (CritterPro.Healing + CritterPro.b_Healing)
            Case "Exp Value"
                Table(n) &= spr & CritterPro.ExpVal
                '
            Case "Small Guns [Skill]"
                Table(n) &= spr & CStr(CalcStats.SmallGun_Skill(CritterPro.Agility) + CritterPro.SmallGuns)
            Case "Big Guns [Skill]"
                Table(n) &= spr & CStr(CalcStats.BigEnergyGun_Skill(CritterPro.Agility) + CritterPro.BigGuns)
            Case "Energy Weapons [Skill]"
                Table(n) &= spr & CStr(CalcStats.BigEnergyGun_Skill(CritterPro.Agility) + CritterPro.EnergyGun)
            Case "Unarmed [Skill]"
                Table(n) &= spr & CStr(CalcStats.Unarmed_Skill(CritterPro.Agility, CritterPro.Strength) + CritterPro.Unarmed)
            Case "Melee [Skill]"
                Table(n) &= spr & CStr(CalcStats.Melee_Skill(CritterPro.Agility, CritterPro.Strength) + CritterPro.Melee)
            Case "Throwing [Skill]"
                Table(n) &= spr & CStr(CalcStats.Throwing_Skill(CritterPro.Agility) + CritterPro.Throwing)
                '
            Case "Resistance Radiation"
                Table(n) &= spr & (CritterPro.DRRadiation + CritterPro.b_DRRadiation)
            Case "Resistance Poison"
                Table(n) &= spr & (CritterPro.DRPoison + CritterPro.b_DRPoison)
                '
            Case "Normal DT|DR"
                Table(n) &= spr & CritterPro.b_DTNormal & "|" & CritterPro.b_DRNormal
            Case "Laser DT|DR"
                Table(n) &= spr & CritterPro.b_DTLaser & "|" & CritterPro.b_DRLaser
            Case "Fire DT|DR"
                Table(n) &= spr & CritterPro.b_DTFire & "|" & CritterPro.b_DRFire
            Case "Plasma DT|DR"
                Table(n) &= spr & CritterPro.b_DTPlasma & "|" & CritterPro.b_DRPlasma
            Case "Electrical DT|DR"
                Table(n) &= spr & CritterPro.b_DTElectrical & "|" & CritterPro.b_DRElectrical
            Case "EMP DT|DR"
                Table(n) &= spr & CritterPro.b_DTEMP & "|" & CritterPro.b_DREMP
            Case "Explosion DT|DR"
                Table(n) &= spr & CritterPro.b_DTExplode & "|" & CritterPro.b_DRExplode
                '
            Case "Base Normal DT|DR"
                Table(n) &= spr & CritterPro.DTNormal & "|" & CritterPro.DRNormal
            Case "Base Laser DT|DR"
                Table(n) &= spr & CritterPro.DTLaser & "|" & CritterPro.DRLaser
            Case "Base Fire DT|DR"
                Table(n) &= spr & CritterPro.DTFire & "|" & CritterPro.DRFire
            Case "Base Plasma DT|DR"
                Table(n) &= spr & CritterPro.DTPlasma & "|" & CritterPro.DRPlasma
            Case "Base Electrical DT|DR"
                Table(n) &= spr & CritterPro.DTElectrical & "|" & CritterPro.DRElectrical
            Case "Base EMP DT|DR"
                Table(n) &= spr & CritterPro.DTEMP & "|" & CritterPro.DREMP
            Case "Base Explosion DT|DR"
                Table(n) &= spr & CritterPro.DTExplode & "|" & CritterPro.DRExplode
        End Select
    End Sub

    Private Function GetTable_Param(ByRef tParam As String) As Integer
        If tParam <> Nothing Then
            Dim y As Byte = InStr(tParam, "[", CompareMethod.Binary)
            Return tParam.Substring(y, tParam.Length - (y + 1))
        End If

        Return &HFFFFFFFF
    End Function

    Friend Sub Items_ImportTable(ByVal tableFile As String)
        Dim CommonItem As CmItemPro, WeaponItem As WpItemPro, ArmorItem As ArItemPro
        Dim AmmoItem As AmItemPro, DrugItem As DgItemPro, MiscItem As McItemPro
        Dim n, m As Integer
        '
        Dim Table_File() As String
        Try
            Table_File = File.ReadAllLines(tableFile, Encoding.Default)
        Catch ex As Exception
            MsgBox("Can not open this table file!", MsgBoxStyle.Critical, "Open error")
            Exit Sub
        End Try
        Dim Table_Param() As String = Split(Table_File(0), spr)
        Dim Table_Value(UBound(Table_File) - 1, UBound(Table_Param)) As String

        'разделить
        For n = 1 To UBound(Table_File)
            Dim tLine() As String = Split(Table_File(n), spr)
            If tLine(0) <> Nothing OrElse tLine.Length < Table_Param.Length Then
                If tLine(0) <> Nothing Then
                    TableLog_Form.ListBox1.Items.Add("Skip Line: Table Line " & (n + 1) & " - Used '#' symbol in begin line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning: Table Line " & (n + 1) & " - Error count value parametr.")
                End If
                Continue For
            End If
            For m = 0 To UBound(Table_Param)
                Table_Value(n - 1, m) = tLine(m)
            Next
        Next

        'открыть профайл
        If Not Directory.Exists(SaveMOD_Path & PROTO_ITEMS) Then Directory.CreateDirectory(SaveMOD_Path & PROTO_ITEMS)
        Dim strSplit() As String
        For n = 0 To UBound(Table_Value)
            Dim ProFile As String = Table_Value(n, 1)
            If ProFile = Nothing Then Continue For
            Dim pPath = SaveMOD_Path & PROTO_ITEMS & ProFile
            If Not File.Exists(pPath) Then
                Dim source As String = DatFiles.CheckFile(PROTO_ITEMS & ProFile) & PROTO_ITEMS & ProFile
                FileSystem.CopyFile(source, pPath, FileIO.UIOption.AllDialogs)
            End If

            Dim iType As Integer
            Select Case FileSystem.GetFileInfo(pPath).Length
                Case 69 'Misc
                    iType = ItemType.Misc
                Case 122 'Оружие 
                    iType = ItemType.Weapon
                Case 125 'Наркотик
                    iType = ItemType.Drugs
                Case 81 'Патрон
                    iType = ItemType.Ammo
                Case 129 'Броня
                    iType = ItemType.Armor
                Case Else
                    TableLog_Form.ListBox1.Items.Add("Error: Pro file '" & ProFile & "' item type does not match the file size.")
                    Continue For
            End Select
            ProFiles.LoadItemProData(pPath, iType, CommonItem, WeaponItem, ArmorItem, AmmoItem, DrugItem, MiscItem)

            Try
                'изменить значения 
                For m = 3 To UBound(Table_Param)
                    Select Case Table_Param(m)
                        'Common
                        Case "Cost"
                            CommonItem.Cost = CInt(Table_Value(n, m))
                        Case "Weight"
                            CommonItem.Weight = CInt(Table_Value(n, m))
                        Case "Size"
                            CommonItem.Size = CInt(Table_Value(n, m))
                        Case "Shoot Thru [Flag]"
                            If Table_Value(n, m) = "True" Then
                                CommonItem.Falgs = CommonItem.Falgs Or &H80000000
                            Else
                                CommonItem.Falgs = CommonItem.Falgs And (Not (&H80000000))
                            End If
                        Case "Light Thru [Flag]"
                            If Table_Value(n, m) = "True" Then
                                CommonItem.Falgs = CommonItem.Falgs Or &H20000000
                            Else
                                CommonItem.Falgs = CommonItem.Falgs And (Not (&H20000000))
                            End If
                            'weapon
                        Case "Min Strength"
                            WeaponItem.MinST = CInt(Table_Value(n, m))
                        Case "Damage Type"
                            For z = 0 To UBound(DmgType)
                                If Table_Value(n, m) = DmgType(z) Then
                                    WeaponItem.DmgType = z
                                    Exit For
                                End If
                            Next
                        Case "Min Damage"
                            WeaponItem.MinDmg = CInt(Table_Value(n, m))
                        Case "Max Damage"
                            WeaponItem.MaxDmg = CInt(Table_Value(n, m))
                        Case "Range Primary Attack"
                            WeaponItem.MaxRangeP = CInt(Table_Value(n, m))
                        Case "Range Secondary Attack"
                            WeaponItem.MaxRangeS = CInt(Table_Value(n, m))
                        Case "AP Cost Primary Attack"
                            WeaponItem.MPCostP = CInt(Table_Value(n, m))
                        Case "AP Cost Secondary Attack"
                            WeaponItem.MPCostS = CInt(Table_Value(n, m))
                        Case "Max Ammo"
                            WeaponItem.MaxAmmo = CInt(Table_Value(n, m))
                        Case "Rounds Brust"
                            WeaponItem.Rounds = CInt(Table_Value(n, m))
                        Case "Caliber"
                            If iType = ItemType.Weapon Then 'weapon
                                If Table_Value(n, m) <> Nothing Then
                                    For z = 0 To UBound(CaliberNAME)
                                        If Table_Value(n, m) = CaliberNAME(z) Then
                                            WeaponItem.Caliber = z
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else 'ammo
                                If Table_Value(n, m) <> Nothing Then
                                    For z = 0 To UBound(CaliberNAME)
                                        If Table_Value(n, m) = CaliberNAME(z) Then
                                            AmmoItem.Caliber = z
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Case "Ammo PID"
                            WeaponItem.AmmoPID = GetTable_Param(Table_Value(n, m))
                        Case "Critical Fail"
                            WeaponItem.CritFail = CInt(Table_Value(n, m))
                        Case "Perk"
                            If iType = ItemType.Weapon Then 'weapon
                                WeaponItem.Perk = GetTable_Param(Table_Value(n, m))
                            Else 'armor
                                ArmorItem.Perk = GetTable_Param(Table_Value(n, m))
                            End If
                            'Ammo
                        Case "Dam Div"
                            AmmoItem.DamDiv = CInt(Table_Value(n, m))
                        Case "Dam Mult"
                            AmmoItem.DamMult = CInt(Table_Value(n, m))
                        Case "AC Adjust"
                            AmmoItem.ACAdjust = CInt(Table_Value(n, m))
                        Case "DR Adjust"
                            AmmoItem.DRAdjust = CInt(Table_Value(n, m))
                        Case "Quantity"
                            AmmoItem.Quantity = CInt(Table_Value(n, m))
                            'Armor
                        Case "Armor Class"
                            ArmorItem.AC = CInt(Table_Value(n, m))
                        Case "Normal DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTNormal = CInt(strSplit(0))
                            ArmorItem.DRNormal = CInt(strSplit(1))
                        Case "Laser DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTLaser = CInt(strSplit(0))
                            ArmorItem.DRLaser = CInt(strSplit(1))
                        Case "Fire DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTFire = CInt(strSplit(0))
                            ArmorItem.DRFire = CInt(strSplit(1))
                        Case "Plasma DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTPlasma = CInt(strSplit(0))
                            ArmorItem.DRPlasma = CInt(strSplit(1))
                        Case "Electrical DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTElectrical = CInt(strSplit(0))
                            ArmorItem.DRElectrical = CInt(strSplit(1))
                        Case "EMP DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTEMP = CInt(strSplit(0))
                            ArmorItem.DREMP = CInt(strSplit(1))
                        Case "Explosion DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            ArmorItem.DTExplode = CInt(strSplit(0))
                            ArmorItem.DRExplode = CInt(strSplit(1))
                            'Drug
                        Case "Modify Stat 0"
                            DrugItem.Stat0 = GetTable_Param(Table_Value(n, m))
                        Case "Modify Stat 1"
                            DrugItem.Stat1 = GetTable_Param(Table_Value(n, m))
                        Case "Modify Stat 2"
                            DrugItem.Stat2 = GetTable_Param(Table_Value(n, m))
                        Case "Instant Amount 0"
                            DrugItem.iAmount0 = CInt(Table_Value(n, m))
                        Case "Instant Amount 1"
                            DrugItem.iAmount1 = CInt(Table_Value(n, m))
                        Case "Instant Amount 2"
                            DrugItem.iAmount2 = CInt(Table_Value(n, m))
                        Case "First Amount 0"
                            DrugItem.fAmount0 = CInt(Table_Value(n, m))
                        Case "First Amount 1"
                            DrugItem.fAmount1 = CInt(Table_Value(n, m))
                        Case "First Amount 2"
                            DrugItem.fAmount2 = CInt(Table_Value(n, m))
                        Case "First Duration Time"
                            DrugItem.Duration1 = CInt(Table_Value(n, m))
                        Case "Second Amount 0"
                            DrugItem.fAmount0 = CInt(Table_Value(n, m))
                        Case "Second Amount 1"
                            DrugItem.fAmount1 = CInt(Table_Value(n, m))
                        Case "Second Amount 2"
                            DrugItem.fAmount2 = CInt(Table_Value(n, m))
                        Case "Second Duration Time"
                            DrugItem.Duration2 = CInt(Table_Value(n, m))
                        Case "Addiction Effect"
                            DrugItem.W_Effect = GetTable_Param(Table_Value(n, m))
                        Case "Addiction Onset Time"
                            DrugItem.W_Onset = CInt(Table_Value(n, m))
                        Case "Addiction Rate"
                            DrugItem.AddictionRate = CInt(Table_Value(n, m))
                            'Misc
                        Case "Power PID"
                            MiscItem.PowerPID = GetTable_Param(Table_Value(n, m))
                        Case "Power Type"
                            If Table_Value(n, m) <> Nothing Then
                                For z = 0 To UBound(CaliberNAME)
                                    If Table_Value(n, m) = CaliberNAME(z) Then
                                        MiscItem.PowerType = z
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "Charges"
                            If Table_Value(n, m) <> Nothing Then MiscItem.Charges = CInt(Table_Value(n, m))
                    End Select
                Next
            Catch
                MsgBox("Error: Param " & Table_Param(m).ToUpper & " PRO Line: " & Table_Value(n, 0), MsgBoxStyle.Critical, "Error table import")
                Continue For
            End Try
            'save profile and goto next profile
            ProFiles.SaveItemProData(pPath, iType, CommonItem, WeaponItem, ArmorItem, AmmoItem, DrugItem, MiscItem)
        Next
        If TableLog_Form.ListBox1.Items.Count > 0 Then TableLog_Form.Show()
        MsgBox("Successfully!", MsgBoxStyle.Information, "Import table")
    End Sub

    Friend Sub Critters_ImportTable(ByVal tableFile As String)
        Dim n As Integer, m As Integer
        Dim ProFile As String
        Dim strSplit() As String

        Dim Table_File() As String
        Try
            Table_File = File.ReadAllLines(tableFile, Encoding.Default)
        Catch ex As Exception
            MsgBox("Can not open this table file!", MsgBoxStyle.Critical, "Open error")
            Exit Sub
        End Try
        Dim Table_Param() As String = Split(Table_File(0), spr)
        Dim Table_Value(UBound(Table_File) - 1, UBound(Table_Param)) As String

        'разделить
        For n = 1 To UBound(Table_File)
            Dim tLine() As String = Split(Table_File(n), spr)
            If tLine(0) <> String.Empty OrElse tLine.Length < Table_Param.Length Then
                If tLine(0) <> String.Empty Then
                    TableLog_Form.ListBox1.Items.Add("Skip Line #" & (n + 1) & " : Used ignore symbol '#' in table line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning Line #" & (n + 1) & " : Error count value param.")
                End If
                Continue For
            End If
            For m = 0 To UBound(Table_Param)
                Table_Value(n - 1, m) = tLine(m)
            Next
        Next

        'Open profile
        If Not Directory.Exists(SaveMOD_Path & PROTO_CRITTERS) Then Directory.CreateDirectory(SaveMOD_Path & PROTO_CRITTERS)
        For n = 0 To UBound(Table_Value)
            ProFile = Table_Value(n, 1)
            If ProFile = Nothing Then Continue For
            Dim filePath = SaveMOD_Path & PROTO_CRITTERS & ProFile
            If Not File.Exists(filePath) Then
                Dim source = DatFiles.CheckFile(PROTO_CRITTERS & ProFile) & PROTO_CRITTERS & ProFile
                FileSystem.CopyFile(source, filePath, FileIO.UIOption.AllDialogs)
            End If
            ProFiles.LoadCritterProData(filePath, CritterPro)

            'Changed values
            Try
                'Common pass 1
                For m = 3 To UBound(Table_Param)
                    Select Case Table_Param(m)
                        Case "Strength"
                            CritterPro.Strength = CInt(Table_Value(n, m))
                        Case "Perception"
                            CritterPro.Perception = CInt(Table_Value(n, m))
                        Case "Endurance"
                            CritterPro.Endurance = CInt(Table_Value(n, m))
                        Case "Charisma"
                            CritterPro.Charisma = CInt(Table_Value(n, m))
                        Case "Intelligence"
                            CritterPro.Intelligence = CInt(Table_Value(n, m))
                        Case "Agility"
                            CritterPro.Agility = CInt(Table_Value(n, m))
                        Case "Luck"
                            CritterPro.Luck = CInt(Table_Value(n, m))
                        Case "Exp Value"
                            CritterPro.ExpVal = CInt(Table_Value(n, m))
                        Case "Damage Type"
                            For z = 0 To UBound(DmgType)
                                If Table_Value(n, m) = DmgType(z) Then
                                    CritterPro.DamageType = z
                                    Exit For
                                End If
                            Next
                            'Armor
                        Case "Normal DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTNormal = CInt(strSplit(0))
                            CritterPro.b_DRNormal = CInt(strSplit(1))
                        Case "Laser DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTLaser = CInt(strSplit(0))
                            CritterPro.b_DRLaser = CInt(strSplit(1))
                        Case "Fire DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTFire = CInt(strSplit(0))
                            CritterPro.b_DRFire = CInt(strSplit(1))
                        Case "Plasma DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTPlasma = CInt(strSplit(0))
                            CritterPro.b_DRPlasma = CInt(strSplit(1))
                        Case "Electrical DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTElectrical = CInt(strSplit(0))
                            CritterPro.b_DRElectrical = CInt(strSplit(1))
                        Case "EMP DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTEMP = CInt(strSplit(0))
                            CritterPro.b_DREMP = CInt(strSplit(1))
                        Case "Explosion DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.b_DTExplode = CInt(strSplit(0))
                            CritterPro.b_DRExplode = CInt(strSplit(1))
                            '
                        Case "Base Normal DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTNormal = CInt(strSplit(0))
                            CritterPro.DRNormal = CInt(strSplit(1))
                        Case "Base Laser DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTLaser = CInt(strSplit(0))
                            CritterPro.DRLaser = CInt(strSplit(1))
                        Case "Base Fire DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTFire = CInt(strSplit(0))
                            CritterPro.DRFire = CInt(strSplit(1))
                        Case "Base Plasma DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTPlasma = CInt(strSplit(0))
                            CritterPro.DRPlasma = CInt(strSplit(1))
                        Case "Base Electrical DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTElectrical = CInt(strSplit(0))
                            CritterPro.DRElectrical = CInt(strSplit(1))
                        Case "Base EMP DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTEMP = CInt(strSplit(0))
                            CritterPro.DREMP = CInt(strSplit(1))
                        Case "Base Explosion DT|DR"
                            strSplit = Table_Value(n, m).Split(splt)
                            CritterPro.DTExplode = CInt(strSplit(0))
                            CritterPro.DRExplode = CInt(strSplit(1))
                    End Select
                Next
                'Calculate pass 2
                For m = 3 To UBound(Table_Param)
                    Select Case Table_Param(m)
                        Case "Action Point"
                            CritterPro.AP = CalcStats.Action_Point(CritterPro.Agility)
                            CritterPro.b_AP = CInt(Table_Value(n, m)) - CritterPro.AP
                        Case "Armor Class"
                            CritterPro.AC = CritterPro.Agility
                            CritterPro.b_AC = CInt(Table_Value(n, m)) - CritterPro.Agility
                        Case "Health Point"
                            CritterPro.HP = CalcStats.Health_Point(CritterPro.Strength, CritterPro.Endurance)
                            CritterPro.b_HP = CInt(Table_Value(n, m)) - CritterPro.HP
                        Case "Healing Rate"
                            CritterPro.Healing = CalcStats.Healing_Rate(CritterPro.Endurance)
                            CritterPro.b_Healing = CInt(Table_Value(n, m)) - CritterPro.Healing
                        Case "Melee Damage"
                            CritterPro.MeleeDmg = CalcStats.Melee_Damage(CritterPro.Strength)
                            CritterPro.b_MeleeDmg = CInt(Table_Value(n, m)) - CritterPro.MeleeDmg
                        Case "Critical Chance"
                            CritterPro.Critical = CritterPro.Luck
                            CritterPro.b_Critical = CInt(Table_Value(n, m)) - CritterPro.Luck
                        Case "Sequence"
                            CritterPro.Sequence = CalcStats.Sequence(CritterPro.Perception)
                            CritterPro.b_Sequence = CInt(Table_Value(n, m)) - CritterPro.Sequence
                        Case "Resistance Radiation"
                            CritterPro.DRRadiation = CalcStats.Radiation(CritterPro.Endurance)
                            CritterPro.b_DRRadiation = CInt(Table_Value(n, m)) - CritterPro.DRRadiation
                        Case "Resistance Poison"
                            CritterPro.DRPoison = CalcStats.Poison(CritterPro.Endurance)
                            CritterPro.b_DRPoison = CInt(Table_Value(n, m)) - CritterPro.DRPoison
                            'Skill
                        Case "Small Guns [Skill]"
                            CritterPro.SmallGuns = CInt(Table_Value(n, m)) - CalcStats.SmallGun_Skill(CritterPro.Agility)
                        Case "Big Guns [Skill]"
                            CritterPro.BigGuns = CInt(Table_Value(n, m)) - CalcStats.BigEnergyGun_Skill(CritterPro.Agility)
                        Case "Energy Weapons [Skill]"
                            CritterPro.EnergyGun = CInt(Table_Value(n, m)) - CalcStats.BigEnergyGun_Skill(CritterPro.Agility)
                        Case "Unarmed [Skill]"
                            CritterPro.Unarmed = CInt(Table_Value(n, m)) - CalcStats.Unarmed_Skill(CritterPro.Agility, CritterPro.Strength)
                        Case "Melee [Skill]"
                            CritterPro.Melee = CInt(Table_Value(n, m)) - CalcStats.Melee_Skill(CritterPro.Agility, CritterPro.Strength)
                        Case "Throwing [Skill]"
                            CritterPro.Throwing = CInt(Table_Value(n, m)) - CalcStats.Throwing_Skill(CritterPro.Agility)
                    End Select
                Next
            Catch
                MsgBox("Error: Param " & Table_Param(m).ToUpper & " PRO Line: " & Table_Value(n, 0), MsgBoxStyle.Critical, "Error Import")
                TableLog_Form.ListBox1.Items.Add("Error Line #" & (n + 1) & " : Error in value param (" & Table_Param(m) & ")")
                Continue For
            End Try
            'Save the profile and goto next profile
            ProFiles.SaveCritterProData(filePath, CritterPro)
        Next
        If TableLog_Form.ListBox1.Items.Count > 0 Then TableLog_Form.Show()
        MsgBox("Successfully!", MsgBoxStyle.Information, "Import table")
    End Sub

    Private Sub CheckAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CheckAllToolStripMenuItem.Click
        CheckedItemsAll(True)
    End Sub

    Private Sub DeselecAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeselecAllToolStripMenuItem.Click
        CheckedItemsAll(False)
    End Sub

    Private Sub CheckedItemsAll(ByVal value As Boolean)
        Dim control As CheckedListBox = GetCheckList()

        For n = 0 To control.Items.Count - 1
            control.SetItemChecked(n, value)
        Next
    End Sub

    Private Sub Create_Button(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        CheckedList = GetCheckList.CheckedItems
        CreateTable()
    End Sub

    Private Function GetCheckList() As CheckedListBox
        Select Case TabControl1.SelectedIndex
            Case TabType.Critter
                Return CheckedListBox6
            Case TabType.Weapon
                Return CheckedListBox1
            Case TabType.Ammo
                Return CheckedListBox2
            Case TabType.Armor
                Return CheckedListBox3
            Case TabType.Drugs
                Return CheckedListBox4
            Case Else 'misc
                Return CheckedListBox5
        End Select
    End Function
End Class