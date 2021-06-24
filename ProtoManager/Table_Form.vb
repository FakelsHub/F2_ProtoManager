Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Imports Prototypes
Imports Enums

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


    Private Sub CreateTable()
        Dim critterPro As CritPro
        Dim commonItem As CmItemPro
        Dim weaponItem As WpItemPro
        Dim armorItem As ArItemPro
        Dim ammoItem As AmItemPro
        Dim drugItem As DgItemPro
        Dim miscItem As McItemPro

        Dim fFile As Integer, iType As Integer = TabControl1.SelectedIndex
        Dim cPath, pathFile As String

        Dim table As List(Of String) = New List(Of String)
        table.Add("Import" & spr & "ProFILE" & spr & "NAME")

        If iType > TabType.Critter Then
            For n = 0 To UBound(Items_LST)
                If Items_LST(n).itemType = Array.IndexOf(ItemTypesName, TabControl1.SelectedTab.Text) Then
                    table.Add(Items_LST(n).proFile)
                End If
            Next

            Dim IsRead As Boolean = False

            Main.GetItemsData()
            GetMsgData("pro_item.msg")

            'Dim tableLine As StringBuilder = New StringBuilder()

            For n = 1 To table.Count - 1
                cPath = DatFiles.CheckFile(PROTO_ITEMS & table(n), False)
                pathFile = String.Concat(cPath, PROTO_ITEMS, table(n))

                Dim cmProDataBuf(Prototypes.ItemCommonLen - 1) As Integer

                fFile = FreeFile()
                FileOpen(fFile, pathFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)

                FileGet(fFile, cmProDataBuf)
                ProFiles.ReverseLoadData(cmProDataBuf, commonItem)
                FileGet(fFile, commonItem.SoundID)

                'tableLine.Clear()

                If cPath.Equals(SaveMOD_Path, StringComparison.OrdinalIgnoreCase) Then
                    table(n) = spr & table(n)
                Else
                    table(n) = "#" & spr & table(n) ' # - ignore mark
                End If

                table(n) &= (spr & Messages.GetNameObject(commonItem.DescID)) ' get name

                For m = 0 To CheckedList.Count - 1
                    If n = 1 Then table(0) &= spr & CheckedList.Item(m).ToString

                    If (CreateTable_Common(table(n), CheckedList.Item(m).ToString, commonItem)) Then Continue For

                    Select Case iType
                        Case TabType.Weapon
                            If Not IsRead Then
                                Dim wnProDataBuf(Prototypes.ItemWeaponLen - 1) As Integer
                                FileGet(fFile, wnProDataBuf)
                                ProFiles.ReverseLoadData(wnProDataBuf, weaponItem)
                                FileGet(fFile, weaponItem.wSoundID)
                                FileClose(fFile)
                                IsRead = True
                            End If
                            CreateTable_Weapon(table(n), CheckedList.Item(m).ToString, weaponItem)
                        Case TabType.Ammo
                            If Not IsRead Then
                                Dim amProDataBuf(Prototypes.ItemAmmoLen - 1) As Integer
                                FileGet(fFile, amProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(amProDataBuf, ammoItem)
                                IsRead = True
                            End If
                            CreateTable_Ammo(table(n), CheckedList.Item(m).ToString, ammoItem)
                        Case TabType.Armor
                            If Not IsRead Then
                                Dim arProDataBuf(Prototypes.ItemArmorLen - 1) As Integer
                                FileGet(fFile, arProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(arProDataBuf, armorItem)
                                IsRead = True
                            End If
                            CreateTable_Armor(table(n), CheckedList.Item(m).ToString, armorItem)
                        Case TabType.Drugs
                            If Not IsRead Then
                                Dim drProDataBuf(Prototypes.ItemDrugsLen - 1) As Integer
                                FileGet(fFile, drProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(drProDataBuf, drugItem)
                                IsRead = True
                            End If
                            CreateTable_Drugs(table(n), CheckedList.Item(m).ToString, drugItem)
                        Case Else ' misc
                            If Not IsRead Then
                                Dim msProDataBuf(Prototypes.ItemMiscLen - 1) As Integer
                                FileGet(fFile, msProDataBuf)
                                FileClose(fFile)
                                ProFiles.ReverseLoadData(msProDataBuf, miscItem)
                                IsRead = True
                            End If
                            CreateTable_Misc(table(n), CheckedList.Item(m).ToString, miscItem)
                    End Select
                Next
                IsRead = False
                'table(n) &= tableLine.ToString
            Next
        Else
            ' Critter table
            If Critter_LST Is Nothing Then Main.CreateCritterList()

            Progress_Form.ShowProgressBar(CInt(UBound(Critter_LST) / 2))

            GetMsgData("pro_crit.msg")
            For n = 1 To UBound(Critter_LST) + 1

                Dim proFile As String = Critter_LST(n - 1).proFile
                cPath = DatFiles.CheckFile(PROTO_CRITTERS & proFile, False)
                pathFile = String.Concat(cPath, PROTO_CRITTERS & proFile)

                If FileSystem.GetFileInfo(pathFile).Length < 416 Then
                    table.Add("#" & spr & proFile & spr & "<BadFormat>")
                    'Log
                    Main.PrintLog("Bad Format: " & pathFile)
                    Application.DoEvents()
                    Continue For
                End If

                ProFiles.LoadCritterProData(pathFile, critterPro)

                If cPath.Equals(SaveMOD_Path, StringComparison.OrdinalIgnoreCase) Then
                    table.Add(spr & proFile)
                Else
                    table.Add("#" & spr & proFile) ' # - ignore mark
                End If

                table(n) &= spr & Messages.GetNameObject(critterPro.DescID)
                For m = 0 To CheckedList.Count - 1
                    'создаем строку с параметрами
                    If n = 1 Then table(0) &= spr & CheckedList.Item(m).ToString
                    CreateTable_Critter(table(n), CheckedList.Item(m).ToString, critterPro)
                Next
                If ((n Mod 2) <> 0) Then Progress_Form.ProgressBar1.Value += 1
            Next
        End If

        SaveTable(TabControl1.SelectedTab.Text, table)
        Progress_Form.Close()
    End Sub

    Private Sub SaveTable(ByVal fileName As String, table As List(Of String))
        SaveFileDialog1.FileName = fileName
        If SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
        fileName = SaveFileDialog1.FileName

SaveRetry:
        Try
            File.WriteAllLines(fileName, table, ASCIIEncoding.Default)
        Catch
            If MsgBox("Error save table file!", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then GoTo SaveRetry
            SaveTable(fileName, table)
        End Try

        If MsgBox("Open saved table file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Process.Start(SaveFileDialog1.FileName)
    End Sub

    Private Function CreateTable_Common(ByRef tableLine As String, param As String, ByRef item As CmItemPro) As Boolean
        Select Case param
            Case "Cost"
                tableLine &= spr & item.Cost
            Case "Weight"
                tableLine &= spr & item.Weight
            Case "Size"
                tableLine &= spr & item.Size
            Case "Shoot Thru [Flag]"
                tableLine &= spr & CBool(item.Flags And Flags.ShootThru)
            Case "Light Thru [Flag]"
                tableLine &= spr & CBool(item.Flags And Flags.LightThru)
            Case Else
                Return False
        End Select
        Return True
    End Function

    Private Sub CreateTable_Weapon(ByRef tableLine As String, param As String, ByRef item As WpItemPro)
        Select Case param
            Case "Min Strength"
                tableLine &= spr & item.MinST
            Case "Damage Type"
                tableLine &= spr & DmgType(item.DmgType)
            Case "Min Damage"
                tableLine &= spr & item.MinDmg
            Case "Max Damage"
                tableLine &= spr & item.MaxDmg
            Case "Range Primary Attack"
                tableLine &= spr & item.MaxRangeP
            Case "Range Secondary Attack"
                tableLine &= spr & item.MaxRangeS
            Case "AP Cost Primary Attack"
                tableLine &= spr & item.MPCostP
            Case "AP Cost Secondary Attack"
                tableLine &= spr & item.MPCostS
            Case "Max Ammo"
                tableLine &= spr & item.MaxAmmo
            Case "Rounds Brust"
                tableLine &= spr & item.Rounds
            Case "Caliber"
                If item.Caliber <> &HFFFFFFFF Then
                    tableLine &= spr & CaliberNAME(item.Caliber)
                Else
                    tableLine &= spr
                End If
            Case "Ammo PID"
                If item.AmmoPID <> &HFFFFFFFF Then
                    tableLine &= spr & Items_LST(item.AmmoPID - 1).itemName & " [" & item.AmmoPID & "]"
                Else
                    tableLine &= spr
                End If
            Case "Critical Fail"
                tableLine &= spr & item.CritFail
            Case "Perk"
                If item.Perk <> &HFFFFFFFF Then
                    tableLine &= spr & Perk_NAME(item.Perk) & " [" & item.Perk & "]"
                Else
                    tableLine &= spr
                End If
        End Select
    End Sub

    Private Sub CreateTable_Ammo(ByRef tableLine As String, param As String, ByRef item As AmItemPro)
        Select Case param
            Case "Dam Div"
                tableLine &= spr & item.DamDiv
            Case "Dam Mult"
                tableLine &= spr & item.DamMult
            Case "AC Adjust"
                tableLine &= spr & item.ACAdjust
            Case "DR Adjust"
                tableLine &= spr & item.DRAdjust
            Case "Quantity"
                tableLine &= spr & item.Quantity
            Case "Caliber"
                If item.Caliber <> -1 Then
                    tableLine &= spr & CaliberNAME(item.Caliber)
                Else
                    tableLine &= spr
                End If
        End Select
    End Sub

    Private Sub CreateTable_Armor(ByRef tableLine As String, param As String, ByRef item As ArItemPro)
        Select Case param
            Case "Armor Class"
                tableLine &= spr & item.AC
            Case "Normal DT|DR"
                tableLine &= spr & item.DTNormal & "|" & item.DRNormal
            Case "Laser DT|DR"
                tableLine &= spr & item.DTLaser & "|" & item.DRLaser
            Case "Fire DT|DR"
                tableLine &= spr & item.DTFire & "|" & item.DRFire
            Case "Plasma DT|DR"
                tableLine &= spr & item.DTPlasma & "|" & item.DRPlasma
            Case "Electrical DT|DR"
                tableLine &= spr & item.DTElectrical & "|" & item.DRElectrical
            Case "EMP DT|DR"
                tableLine &= spr & item.DTEMP & "|" & item.DREMP
            Case "Explosion DT|DR"
                tableLine &= spr & item.DTExplode & "|" & item.DRExplode
            Case "Perk"
                If item.Perk <> -1 Then
                    tableLine &= spr & Perk_NAME(item.Perk) & " [" & item.Perk & "]"
                Else
                    tableLine &= spr
                End If
        End Select
    End Sub

    Private Sub CreateTable_Drugs(ByRef tableLine As String, param As String, ByRef item As DgItemPro)
        Select Case param
            Case "Modify Stat 0"
                If item.Stat0 <> &HFFFFFFFF Then
                    tableLine &= spr & DrugEffect(2 + (item.Stat0)).ToString & " [" & item.Stat0 & "]"
                Else
                    tableLine &= spr
                End If
            Case "Modify Stat 1"
                If item.Stat1 <> &HFFFFFFFF Then
                    tableLine &= spr & DrugEffect(2 + (item.Stat1)).ToString & " [" & item.Stat1 & "]"
                Else
                    tableLine &= spr
                End If
            Case "Modify Stat 2"
                If item.Stat2 <> &HFFFFFFFF Then
                    tableLine &= spr & DrugEffect(2 + (item.Stat2)).ToString & " [" & item.Stat2 & "]"
                Else
                    tableLine &= spr
                End If
            Case "Instant Amount 0"
                tableLine &= spr & item.iAmount0
            Case "Instant Amount 1"
                tableLine &= spr & item.iAmount1
            Case "Instant Amount 2"
                tableLine &= spr & item.iAmount2
            Case "First Amount 0"
                tableLine &= spr & item.fAmount0
            Case "First Amount 1"
                tableLine &= spr & item.fAmount1
            Case "First Amount 2"
                tableLine &= spr & item.fAmount2
            Case "First Duration Time"
                tableLine &= spr & item.Duration1
            Case "Second Amount 0"
                tableLine &= spr & item.fAmount0
            Case "Second Amount 1"
                tableLine &= spr & item.fAmount1
            Case "Second Amount 2"
                tableLine &= spr & item.fAmount2
            Case "Second Duration Time"
                tableLine &= spr & item.Duration2
            Case "Addiction Effect"
                If item.W_Effect <> -1 Then
                    tableLine &= spr & Perk_NAME(item.W_Effect) & " [" & item.W_Effect & "]"
                Else
                    tableLine &= spr
                End If
            Case "Addiction Onset Time"
                tableLine &= spr & item.W_Onset
            Case "Addiction Rate"
                tableLine &= spr & item.AddictionRate
        End Select
    End Sub

    Private Sub CreateTable_Misc(ByRef tableLine As String, param As String, ByRef item As McItemPro)
        Select Case param
            Case "Power PID"
                If item.PowerPID <> &HFFFFFFFF Then
                    tableLine &= spr & Items_LST(item.PowerPID - 1).itemName & " [" & item.PowerPID & "]"
                Else
                    tableLine &= spr
                End If
            Case "Power Type"
                If item.PowerType < UBound(CaliberNAME) Then
                    tableLine &= spr & CaliberNAME(item.PowerType)
                Else
                    tableLine &= spr
                End If
            Case "Charges"
                tableLine &= spr & item.Charges
        End Select
    End Sub

    Private Sub CreateTable_Critter(ByRef tableLine As String, param As String, ByRef critter As CritPro)
        Select Case param
            Case "Strength"
                tableLine &= spr & critter.Strength
            Case "Perception"
                tableLine &= spr & critter.Perception
            Case "Endurance"
                tableLine &= spr & critter.Endurance
            Case "Charisma"
                tableLine &= spr & critter.Charisma
            Case "Intelligence"
                tableLine &= spr & critter.Intelligence
            Case "Agility"
                tableLine &= spr & critter.Agility
            Case "Luck"
                tableLine &= spr & critter.Luck
                '
            Case "Health Point"
                tableLine &= spr & (critter.HP + critter.b_HP)
            Case "Action Point"
                tableLine &= spr & (critter.AP + critter.b_AP)
            Case "Armor Class"
                tableLine &= spr & (critter.AC + critter.b_AC)
            Case "Melee Damage"
                tableLine &= spr & (critter.MeleeDmg + critter.b_MeleeDmg)
            Case "Damage Type"
                tableLine &= spr & DmgType(critter.DamageType)
            Case "Critical Chance"
                tableLine &= spr & (critter.Critical + critter.b_Critical)
            Case "Sequence"
                tableLine &= spr & (critter.Sequence + critter.b_Sequence)
            Case "Healing Rate"
                tableLine &= spr & (critter.Healing + critter.b_Healing)
            Case "Exp Value"
                tableLine &= spr & critter.ExpVal
                '
            Case "Small Guns [Skill]"
                tableLine &= spr & CStr(CalcStats.SmallGun_Skill(critter.Agility) + critter.SmallGuns)
            Case "Big Guns [Skill]"
                tableLine &= spr & CStr(CalcStats.BigEnergyGun_Skill(critter.Agility) + critter.BigGuns)
            Case "Energy Weapons [Skill]"
                tableLine &= spr & CStr(CalcStats.BigEnergyGun_Skill(critter.Agility) + critter.EnergyGun)
            Case "Unarmed [Skill]"
                tableLine &= spr & CStr(CalcStats.Unarmed_Skill(critter.Agility, critter.Strength) + critter.Unarmed)
            Case "Melee [Skill]"
                tableLine &= spr & CStr(CalcStats.Melee_Skill(critter.Agility, critter.Strength) + critter.Melee)
            Case "Throwing [Skill]"
                tableLine &= spr & CStr(CalcStats.Throwing_Skill(critter.Agility) + critter.Throwing)
                '
            Case "Resistance Radiation"
                tableLine &= spr & (critter.DRRadiation + critter.b_DRRadiation)
            Case "Resistance Poison"
                tableLine &= spr & (critter.DRPoison + critter.b_DRPoison)
                '
            Case "Normal DT|DR"
                tableLine &= spr & critter.b_DTNormal & "|" & critter.b_DRNormal
            Case "Laser DT|DR"
                tableLine &= spr & critter.b_DTLaser & "|" & critter.b_DRLaser
            Case "Fire DT|DR"
                tableLine &= spr & critter.b_DTFire & "|" & critter.b_DRFire
            Case "Plasma DT|DR"
                tableLine &= spr & critter.b_DTPlasma & "|" & critter.b_DRPlasma
            Case "Electrical DT|DR"
                tableLine &= spr & critter.b_DTElectrical & "|" & critter.b_DRElectrical
            Case "EMP DT|DR"
                tableLine &= spr & critter.b_DTEMP & "|" & critter.b_DREMP
            Case "Explosion DT|DR"
                tableLine &= spr & critter.b_DTExplode & "|" & critter.b_DRExplode
                '
            Case "Base Normal DT|DR"
                tableLine &= spr & critter.DTNormal & "|" & critter.DRNormal
            Case "Base Laser DT|DR"
                tableLine &= spr & critter.DTLaser & "|" & critter.DRLaser
            Case "Base Fire DT|DR"
                tableLine &= spr & critter.DTFire & "|" & critter.DRFire
            Case "Base Plasma DT|DR"
                tableLine &= spr & critter.DTPlasma & "|" & critter.DRPlasma
            Case "Base Electrical DT|DR"
                tableLine &= spr & critter.DTElectrical & "|" & critter.DRElectrical
            Case "Base EMP DT|DR"
                tableLine &= spr & critter.DTEMP & "|" & critter.DREMP
            Case "Base Explosion DT|DR"
                tableLine &= spr & critter.DTExplode & "|" & critter.DRExplode
        End Select
    End Sub

    Friend Sub Items_ImportTable(ByVal tableFile As String)
        Dim table() As String
        Try
            table = File.ReadAllLines(tableFile, Encoding.Default)
        Catch ex As Exception
            MsgBox("Can not open this table file!", MsgBoxStyle.Critical, "Open error")
            Exit Sub
        End Try

        Dim tableParam() As String = Split(table(0), spr)
        Dim tableValue(UBound(table) - 1, UBound(tableParam)) As String

        Dim n, m As Integer
        'разделить
        For n = 1 To UBound(table)
            Dim tLine() As String = Split(table(n), spr)
            If tLine(0) <> Nothing OrElse tLine.Length < tableParam.Length Then
                If tLine(0) <> Nothing Then
                    TableLog_Form.ListBox1.Items.Add("Skip Line: Table Line " & (n + 1) & " - Used '#' symbol in begin line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning: Table Line " & (n + 1) & " - Error count value parametr.")
                End If
                Continue For
            End If
            For m = 0 To UBound(tableParam)
                tableValue(n - 1, m) = tLine(m)
            Next
        Next

        If Not Directory.Exists(SaveMOD_Path & PROTO_ITEMS) Then Directory.CreateDirectory(SaveMOD_Path & PROTO_ITEMS)

        Dim item As ItemPrototype

        'открыть профайл
        For n = 0 To UBound(tableValue)
            Dim ProFile As String = tableValue(n, 1)
            If ProFile = Nothing Then Continue For

            Dim pPath = DatFiles.CheckFile(PROTO_ITEMS & ProFile)

            Dim iType As ItemType

            Select Case FileSystem.GetFileInfo(pPath).Length
                Case PrototypeSize.Misc 'Misc
                    iType = ItemType.Misc
                    item = New MiscItemObj(pPath)
                    CType(item, MiscItemObj).Load()
                Case PrototypeSize.Weapon 'Оружие
                    iType = ItemType.Weapon
                    item = New WeaponItemObj(pPath)
                    CType(item, WeaponItemObj).Load()
                Case PrototypeSize.Drug 'Наркотик
                    iType = ItemType.Drugs
                    item = New DrugsItemObj(pPath)
                    CType(item, DrugsItemObj).Load()
                Case PrototypeSize.Ammo 'Патрон
                    iType = ItemType.Ammo
                    item = New AmmoItemObj(pPath)
                    CType(item, AmmoItemObj).Load()
                Case PrototypeSize.Armor 'Броня
                    iType = ItemType.Armor
                    item = New ArmorItemObj(pPath)
                    CType(item, ArmorItemObj).Load()
                Case Else
                    TableLog_Form.ListBox1.Items.Add("Error: Pro file '" & ProFile & "' item type does not match the file size.")
                    Continue For
            End Select

            Try
                'изменить значения
                For m = 3 To UBound(tableParam)
                    Select Case tableParam(m)
                        'Common
                        Case "Cost"
                            item.Cost = CInt(tableValue(n, m))
                        Case "Weight"
                            item.Weight = CInt(tableValue(n, m))
                        Case "Size"
                            item.Size = CInt(tableValue(n, m))
                        Case "Shoot Thru [Flag]"
                            If tableValue(n, m) = "True" Then
                                item.Flags = item.Flags Or Flags.ShootThru
                            Else
                                item.Flags = item.Flags And (Not (Flags.ShootThru))
                            End If
                        Case "Light Thru [Flag]"
                            If tableValue(n, m) = "True" Then
                                item.Flags = item.Flags Or Flags.LightThru
                            Else
                                item.Flags = item.Flags And (Not (Flags.LightThru))
                            End If
                        Case Else
                            Select Case iType
                                Case ItemType.Weapon
                                    SetWeaponParams(tableParam(m), tableValue(n, m), CType(item, WeaponItemObj))
                                Case ItemType.Ammo
                                    SetAmmoParams(tableParam(m), tableValue(n, m), CType(item, AmmoItemObj))
                                Case ItemType.Armor
                                    SetArmorParams(tableParam(m), tableValue(n, m), CType(item, ArmorItemObj))
                                Case ItemType.Drugs
                                    SetDrugsParams(tableParam(m), tableValue(n, m), CType(item, DrugsItemObj))
                                Case ItemType.Misc
                                    SetMiscParams(tableParam(m), tableValue(n, m), CType(item, MiscItemObj))
                            End Select
                    End Select
                Next
            Catch
                MsgBox("Error: Param " & tableParam(m).ToUpper & " PRO Line: " & tableValue(n, 0), MsgBoxStyle.Critical, "Error table import")
                Continue For
            End Try

            'save profile and goto next profile
            ProFiles.SaveItemProData(pPath, iType, item)
        Next
        If TableLog_Form.ListBox1.Items.Count > 0 Then TableLog_Form.Show()
        MsgBox("Successfully!", MsgBoxStyle.Information, "Import table")
    End Sub

    Private Function GetTable_Param(ByRef tParam As String) As Integer
        If tParam <> Nothing Then
            Dim y As Integer = InStr(tParam, "[", CompareMethod.Binary)
            Return Convert.ToInt32(tParam.Substring(y, tParam.Length - (y + 1)))
        End If
        Return -1
    End Function

    Private Sub SetWeaponParams(ByVal tParam As String, ByVal tValue As String, item As WeaponItemObj)
        Select Case tParam
            Case "Min Strength"
                item.MinST = CInt(tValue)
            Case "Damage Type"
                For z = 0 To UBound(DmgType)
                    If tValue = DmgType(z) Then
                        item.DmgType = CType(z, WeaponItemObj.DamageType)
                        Exit For
                    End If
                Next
            Case "Min Damage"
                item.MinDmg = CInt(tValue)
            Case "Max Damage"
                item.MaxDmg = CInt(tValue)
            Case "Range Primary Attack"
                item.MaxRangeP = CInt(tValue)
            Case "Range Secondary Attack"
                item.MaxRangeS = CInt(tValue)
            Case "AP Cost Primary Attack"
                item.MPCostP = CInt(tValue)
            Case "AP Cost Secondary Attack"
                item.MPCostS = CInt(tValue)
            Case "Max Ammo"
                item.MaxAmmo = CInt(tValue)
            Case "Rounds Brust"
                item.Rounds = CInt(tValue)
            Case "Caliber"
                If tValue <> Nothing Then
                    For z = 0 To UBound(CaliberNAME)
                        If tValue = CaliberNAME(z) Then
                            item.Caliber = z
                            Exit For
                        End If
                    Next
                End If
            Case "Ammo PID"
                item.AmmoPID = GetTable_Param(tValue)
            Case "Critical Fail"
                item.CritFail = CInt(tValue)
            Case "Perk"
                item.Perk = GetTable_Param(tValue)
        End Select
    End Sub

    Private Sub SetAmmoParams(ByVal tParam As String, ByVal tValue As String, item As AmmoItemObj)
        Select Case tParam
            Case "Dam Div"
                item.DamDiv = CInt(tValue)
            Case "Dam Mult"
                item.DamMult = CInt(tValue)
            Case "AC Adjust"
                item.ACAdjust = CInt(tValue)
            Case "DR Adjust"
                item.DRAdjust = CInt(tValue)
            Case "Quantity"
                item.Quantity = CInt(tValue)
            Case "Caliber"
                If tValue <> Nothing Then
                    For z = 0 To UBound(CaliberNAME)
                        If tValue = CaliberNAME(z) Then
                            item.Caliber = z
                            Exit For
                        End If
                    Next
                End If
        End Select
    End Sub

    Private Sub SetArmorParams(ByVal tParam As String, ByVal tValue As String, item As ArmorItemObj)
        Dim strSplit() As String
        Select Case tParam
            Case "Armor Class"
                item.AC = CInt(tValue)
            Case "Normal DT|DR"
                strSplit = tValue.Split(splt)
                item.DTNormal = CInt(strSplit(0))
                item.DRNormal = CInt(strSplit(1))
            Case "Laser DT|DR"
                strSplit = tValue.Split(splt)
                item.DTLaser = CInt(strSplit(0))
                item.DRLaser = CInt(strSplit(1))
            Case "Fire DT|DR"
                strSplit = tValue.Split(splt)
                item.DTFire = CInt(strSplit(0))
                item.DRFire = CInt(strSplit(1))
            Case "Plasma DT|DR"
                strSplit = tValue.Split(splt)
                item.DTPlasma = CInt(strSplit(0))
                item.DRPlasma = CInt(strSplit(1))
            Case "Electrical DT|DR"
                strSplit = tValue.Split(splt)
                item.DTElectrical = CInt(strSplit(0))
                item.DRElectrical = CInt(strSplit(1))
            Case "EMP DT|DR"
                strSplit = tValue.Split(splt)
                item.DTEMP = CInt(strSplit(0))
                item.DREMP = CInt(strSplit(1))
            Case "Explosion DT|DR"
                strSplit = tValue.Split(splt)
                item.DTExplode = CInt(strSplit(0))
                item.DRExplode = CInt(strSplit(1))
            Case "Perk"
                item.Perk = GetTable_Param(tValue)
        End Select
    End Sub

    Private Sub SetDrugsParams(ByVal tParam As String, ByVal tValue As String, item As DrugsItemObj)
        Select Case tParam
            Case "Modify Stat 0"
                item.Stat0 = GetTable_Param(tValue)
            Case "Modify Stat 1"
                item.Stat1 = GetTable_Param(tValue)
            Case "Modify Stat 2"
                item.Stat2 = GetTable_Param(tValue)
            Case "Instant Amount 0"
                item.iAmount0 = CInt(tValue)
            Case "Instant Amount 1"
                item.iAmount1 = CInt(tValue)
            Case "Instant Amount 2"
                item.iAmount2 = CInt(tValue)
            Case "First Amount 0"
                item.fAmount0 = CInt(tValue)
            Case "First Amount 1"
                item.fAmount1 = CInt(tValue)
            Case "First Amount 2"
                item.fAmount2 = CInt(tValue)
            Case "First Duration Time"
                item.Duration1 = CInt(tValue)
            Case "Second Amount 0"
                item.fAmount0 = CInt(tValue)
            Case "Second Amount 1"
                item.fAmount1 = CInt(tValue)
            Case "Second Amount 2"
                item.fAmount2 = CInt(tValue)
            Case "Second Duration Time"
                item.Duration2 = CInt(tValue)
            Case "Addiction Effect"
                item.W_Effect = GetTable_Param(tValue)
            Case "Addiction Onset Time"
                item.W_Onset = CInt(tValue)
            Case "Addiction Rate"
                item.AddictionRate = CInt(tValue)
        End Select
    End Sub

    Private Sub SetMiscParams(ByVal tParam As String, ByVal tValue As String, item As MiscItemObj)
        Select Case tParam
            Case "Power PID"
                item.PowerPID = GetTable_Param(tValue)
            Case "Power Type"
                If tValue <> Nothing Then
                    For z = 0 To UBound(CaliberNAME)
                        If tValue = CaliberNAME(z) Then
                            item.PowerType = z
                            Exit For
                        End If
                    Next
                End If
            Case "Charges"
                If tValue <> Nothing Then item.Charges = CInt(tValue)
        End Select
    End Sub

    Friend Sub Critters_ImportTable(ByVal tableFile As String)
        Dim critter As CritPro

        Dim n As Integer, m As Integer
        Dim ProFile As String
        Dim strSplit() As String

        Dim table() As String
        Try
            table = File.ReadAllLines(tableFile, Encoding.Default)
        Catch ex As Exception
            MsgBox("Can not open this table file!", MsgBoxStyle.Critical, "Open error")
            Exit Sub
        End Try
        Dim tableParam() As String = Split(table(0), spr)
        Dim tableValue(UBound(table) - 1, UBound(tableParam)) As String

        'разделить
        For n = 1 To UBound(table)
            Dim tLine() As String = Split(table(n), spr)
            If tLine(0) <> String.Empty OrElse tLine.Length < tableParam.Length Then
                If tLine(0) <> String.Empty Then
                    TableLog_Form.ListBox1.Items.Add("Skip Line #" & (n + 1) & " : Used ignore symbol '#' in table line.")
                Else
                    TableLog_Form.ListBox1.Items.Add("Warning Line #" & (n + 1) & " : Error count value param.")
                End If
                Continue For
            End If
            For m = 0 To UBound(tableParam)
                tableValue(n - 1, m) = tLine(m)
            Next
        Next

        If Not Directory.Exists(SaveMOD_Path & PROTO_CRITTERS) Then Directory.CreateDirectory(SaveMOD_Path & PROTO_CRITTERS)

        'Open profile
        For n = 0 To UBound(tableValue)
            ProFile = tableValue(n, 1)
            If ProFile = Nothing Then Continue For

            Dim filePath = DatFiles.CheckFile(PROTO_CRITTERS & ProFile)
            ProFiles.LoadCritterProData(filePath, critter)

            'Changed values
            Try
                'Common pass 1
                For m = 3 To UBound(tableParam)
                    Select Case tableParam(m)
                        Case "Strength"
                            critter.Strength = CInt(tableValue(n, m))
                        Case "Perception"
                            critter.Perception = CInt(tableValue(n, m))
                        Case "Endurance"
                            critter.Endurance = CInt(tableValue(n, m))
                        Case "Charisma"
                            critter.Charisma = CInt(tableValue(n, m))
                        Case "Intelligence"
                            critter.Intelligence = CInt(tableValue(n, m))
                        Case "Agility"
                            critter.Agility = CInt(tableValue(n, m))
                        Case "Luck"
                            critter.Luck = CInt(tableValue(n, m))
                        Case "Exp Value"
                            critter.ExpVal = CInt(tableValue(n, m))
                        Case "Damage Type"
                            For z = 0 To UBound(DmgType)
                                If tableValue(n, m) = DmgType(z) Then
                                    critter.DamageType = z
                                    Exit For
                                End If
                            Next
                            'Armor
                        Case "Normal DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTNormal = CInt(strSplit(0))
                            critter.b_DRNormal = CInt(strSplit(1))
                        Case "Laser DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTLaser = CInt(strSplit(0))
                            critter.b_DRLaser = CInt(strSplit(1))
                        Case "Fire DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTFire = CInt(strSplit(0))
                            critter.b_DRFire = CInt(strSplit(1))
                        Case "Plasma DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTPlasma = CInt(strSplit(0))
                            critter.b_DRPlasma = CInt(strSplit(1))
                        Case "Electrical DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTElectrical = CInt(strSplit(0))
                            critter.b_DRElectrical = CInt(strSplit(1))
                        Case "EMP DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTEMP = CInt(strSplit(0))
                            critter.b_DREMP = CInt(strSplit(1))
                        Case "Explosion DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.b_DTExplode = CInt(strSplit(0))
                            critter.b_DRExplode = CInt(strSplit(1))
                            '
                        Case "Base Normal DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTNormal = CInt(strSplit(0))
                            critter.DRNormal = CInt(strSplit(1))
                        Case "Base Laser DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTLaser = CInt(strSplit(0))
                            critter.DRLaser = CInt(strSplit(1))
                        Case "Base Fire DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTFire = CInt(strSplit(0))
                            critter.DRFire = CInt(strSplit(1))
                        Case "Base Plasma DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTPlasma = CInt(strSplit(0))
                            critter.DRPlasma = CInt(strSplit(1))
                        Case "Base Electrical DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTElectrical = CInt(strSplit(0))
                            critter.DRElectrical = CInt(strSplit(1))
                        Case "Base EMP DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTEMP = CInt(strSplit(0))
                            critter.DREMP = CInt(strSplit(1))
                        Case "Base Explosion DT|DR"
                            strSplit = tableValue(n, m).Split(splt)
                            critter.DTExplode = CInt(strSplit(0))
                            critter.DRExplode = CInt(strSplit(1))
                    End Select
                Next
                'Calculate pass 2
                For m = 3 To UBound(tableParam)
                    Select Case tableParam(m)
                        Case "Action Point"
                            critter.AP = CalcStats.Action_Point(critter.Agility)
                            critter.b_AP = CInt(tableValue(n, m)) - critter.AP
                        Case "Armor Class"
                            critter.AC = critter.Agility
                            critter.b_AC = CInt(tableValue(n, m)) - critter.Agility
                        Case "Health Point"
                            critter.HP = CalcStats.Health_Point(critter.Strength, critter.Endurance)
                            critter.b_HP = CInt(tableValue(n, m)) - critter.HP
                        Case "Healing Rate"
                            critter.Healing = CalcStats.Healing_Rate(critter.Endurance)
                            critter.b_Healing = CInt(tableValue(n, m)) - critter.Healing
                        Case "Melee Damage"
                            critter.MeleeDmg = CalcStats.Melee_Damage(critter.Strength)
                            critter.b_MeleeDmg = CInt(tableValue(n, m)) - critter.MeleeDmg
                        Case "Critical Chance"
                            critter.Critical = critter.Luck
                            critter.b_Critical = CInt(tableValue(n, m)) - critter.Luck
                        Case "Sequence"
                            critter.Sequence = CalcStats.Sequence(critter.Perception)
                            critter.b_Sequence = CInt(tableValue(n, m)) - critter.Sequence
                        Case "Resistance Radiation"
                            critter.DRRadiation = CalcStats.Radiation(critter.Endurance)
                            critter.b_DRRadiation = CInt(tableValue(n, m)) - critter.DRRadiation
                        Case "Resistance Poison"
                            critter.DRPoison = CalcStats.Poison(critter.Endurance)
                            critter.b_DRPoison = CInt(tableValue(n, m)) - critter.DRPoison
                            'Skill
                        Case "Small Guns [Skill]"
                            critter.SmallGuns = CInt(tableValue(n, m)) - CalcStats.SmallGun_Skill(critter.Agility)
                        Case "Big Guns [Skill]"
                            critter.BigGuns = CInt(tableValue(n, m)) - CalcStats.BigEnergyGun_Skill(critter.Agility)
                        Case "Energy Weapons [Skill]"
                            critter.EnergyGun = CInt(tableValue(n, m)) - CalcStats.BigEnergyGun_Skill(critter.Agility)
                        Case "Unarmed [Skill]"
                            critter.Unarmed = CInt(tableValue(n, m)) - CalcStats.Unarmed_Skill(critter.Agility, critter.Strength)
                        Case "Melee [Skill]"
                            critter.Melee = CInt(tableValue(n, m)) - CalcStats.Melee_Skill(critter.Agility, critter.Strength)
                        Case "Throwing [Skill]"
                            critter.Throwing = CInt(tableValue(n, m)) - CalcStats.Throwing_Skill(critter.Agility)
                    End Select
                Next
            Catch
                MsgBox("Error: Param " & tableParam(m).ToUpper & " PRO Line: " & tableValue(n, 0), MsgBoxStyle.Critical, "Error Import")
                TableLog_Form.ListBox1.Items.Add("Error Line #" & (n + 1) & " : Error in value param (" & tableParam(m) & ")")
                Continue For
            End Try
            'Save the profile and goto next profile
            ProFiles.SaveCritterProData(filePath, critter)
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