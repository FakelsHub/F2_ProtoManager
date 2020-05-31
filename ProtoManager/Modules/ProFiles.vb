Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Imports Prototypes

Module ProFiles

    Friend Enum Status
        NotExist
        IsNormal
        IsModFolder
        IsBadFile
    End Enum

    ''' <summary>
    ''' Возвращает имя Frm файла для инвентаря(ivent), или имя FID предмета, если файл для инвентаря не определен.  
    ''' </summary>
    Friend Function GetItemInvenFID(ByVal nPro As Integer, ByRef Inventory As Boolean) As String
        Dim FID As Integer = -1
        Dim iFID As Integer = -1

        Dim cPath As String = DatFiles.CheckFile(PROTO_ITEMS & Items_LST(nPro).proFile)
        Try
            Using readFile As New BinaryReader(File.Open(cPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                readFile.BaseStream.Seek(Prototypes.offsetFrmID, SeekOrigin.Begin)
                FID = ReverseBytes(readFile.ReadInt32())
                readFile.BaseStream.Seek(Prototypes.offsetInvenFID, SeekOrigin.Begin)
                iFID = ReverseBytes(readFile.ReadInt32())
            End Using
        Catch ex As Exception
            Return Nothing
        End Try

        Dim lstName As String
        If iFID > -1 Then
            iFID = iFID And (Not &H7000000)
            lstName = Iven_FRM.ElementAtOrDefault(iFID)
        Else
            If FID = -1 Then Return Nothing
            lstName = Items_FRM.ElementAtOrDefault(FID)
            Inventory = False
        End If

        If lstName Is Nothing Then
            Main.PrintLog("Invalid FID number of the prototype file: " & PROTO_ITEMS & Items_LST(nPro).proFile)
            Return lstName
        End If
        Return lstName.ToLower
    End Function

    ''' <summary>
    ''' Создает pro-файл по указаному имени и пути.
    ''' </summary>
    Friend Sub CreateProFile(ByVal path As String, ByVal pName As String)
        path = SaveMOD_Path & path
        Dim nProFile As String = path & pName

        If File.Exists(nProFile) Then
            File.SetAttributes(nProFile, FileAttributes.Normal)
            File.Delete(nProFile)
        End If
        If Not (Directory.Exists(path)) Then Directory.CreateDirectory(path)
        File.Move("template", nProFile)
        If proRO Then File.SetAttributes(nProFile, FileAttributes.ReadOnly Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)

        'Log 
        Main.PrintLog("Create Pro: " & nProFile)
    End Sub

    ''' <summary>
    ''' Возвращает номер Description ID из про-файла предмета, и его тип.
    ''' </summary>
    Friend Function GetProItemsNameID(ByRef ProFile As String, ByVal n As Integer) As Integer
        Dim NameID As Integer
        Dim TypeID As Integer = -1

        Dim cPath As String = DatFiles.CheckFile(PROTO_ITEMS & ProFile)

        Try
            Using rFile As New BinaryReader(File.Open(cPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                rFile.BaseStream.Seek(Prototypes.offsetDescID, SeekOrigin.Begin)
                NameID = rFile.ReadInt32()
                rFile.BaseStream.Seek(Prototypes.offsetISubType, SeekOrigin.Begin)
                TypeID = ReverseBytes(rFile.ReadInt32())
            End Using
        Catch ex As EndOfStreamException
            NameID = 0
            TypeID = ItemType.Unknown
            MsgBox("The file is in an incorrect format or damaged." & vbLf & cPath)
        Catch ex As Exception
            TypeID = ItemType.Unknown
        End Try

        ' Определяем тип предмета
        If TypeID >= 0 AndAlso TypeID < ItemType.Unknown Then
            Items_LST(n).itemType = TypeID
        Else
            Items_LST(n).itemType = ItemType.Unknown
        End If

        Return ReverseBytes(NameID)
    End Function

    ''' <summary>
    ''' Проверяет прото файл на соответствие размера и установленного атрибута только-чтения
    ''' </summary>
    ''' <param name="proFile"></param>
    ''' <param name="size"></param>
    ''' <param name="fileAttr"></param>
    ''' <returns>Возвращает результат проверки</returns>
    Friend Function ProtoCheckFile(ByVal proFile As String, ByVal size As Integer, ByRef fileAttr As String) As Status
        Dim cPath As String
        If size <> 416 Then '415
            cPath = DatFiles.CheckFile(PROTO_ITEMS & proFile, unpack:=False)
        Else
            cPath = DatFiles.CheckFile(PROTO_CRITTERS & proFile, unpack:=False)
            If CalcStats.GetFormula = CalcStats.FormulaType.Fallout1 Then size -= 4
        End If
        If cPath = Nothing Then Return Status.NotExist

        Dim pro As New FileInfo(cPath)
        'If pro.Exists = False Then Return Status.NotExist

        If pro.Length <> size Then ' check valid size
            fileAttr = "BAD!"
            Return Status.IsBadFile
        ElseIf pro.DirectoryName.StartsWith(SaveMOD_Path) Then
            If (pro.IsReadOnly) Then fileAttr = "R/O"
            Return Status.IsModFolder
        End If
        Return Status.IsNormal
    End Function

    ''' <summary>
    ''' Возвращает номер FID из про-файла криттера.
    ''' </summary>
    Friend Function GetFID(ByVal nPro As Integer) As Integer
        Dim FID As Integer = -1

        Dim cPath As String = DatFiles.CheckFile(PROTO_CRITTERS & Critter_LST(nPro).proFile)
        Try
            Using readFile As New BinaryReader(File.Open(cPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                readFile.BaseStream.Seek(Prototypes.offsetFrmID, SeekOrigin.Begin)
                FID = ReverseBytes(readFile.ReadInt32())
            End Using
        Catch ex As Exception
            Return Nothing
        End Try

        Return If(FID = -1, 0, FID)
    End Function

    ''' <summary>
    ''' Возвращает имя FID из про-файла криттера.
    ''' </summary>
    Friend Function GetCritterFID(ByVal nPro As Integer) As String
        Dim hp, bhp As Integer
        Dim FID As Integer = -1

        Dim cPath As String = DatFiles.CheckFile(PROTO_CRITTERS & Critter_LST(nPro).proFile)
        Try
            Using readFile As New BinaryReader(File.Open(cPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                readFile.BaseStream.Seek(Prototypes.offsetFrmID, SeekOrigin.Begin)
                FID = ReverseBytes(readFile.ReadInt32())
                readFile.BaseStream.Seek(Prototypes.offsetHP, SeekOrigin.Begin)
                hp = ReverseBytes(readFile.ReadInt32())
                readFile.BaseStream.Seek(Prototypes.offsetbHP, SeekOrigin.Begin)
                bhp = ReverseBytes(readFile.ReadInt32())
            End Using
        Catch ex As Exception
            Return Nothing
        End Try

        Critter_LST(nPro).crtHP = hp + bhp

        If FID = -1 Then Return Nothing
        FID -= &H1000000I

        Return Critters_FRM(FID).ToLower
    End Function

    ''' <summary>
    ''' Возвращает номер Description ID из про-файла криттера.
    ''' </summary>
    Friend Function GetProCritNameID(ByRef ProFile As String) As Integer
        Dim NameID As Integer
        Dim fFile As Byte = FreeFile()
        Dim cPath = DatFiles.CheckFile(PROTO_CRITTERS & ProFile)

        Try
            FileOpen(fFile, cPath, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            FileGet(fFile, NameID, 5)
        Catch
            Return 0
        Finally
            FileClose(fFile)
        End Try

        Return ReverseBytes(NameID)
    End Function

    ''' <summary>
    ''' Удаляет пустые строки в конце массива и лишние пробелы в строке.
    ''' </summary>
    Friend Function ClearEmptyLines(ByVal lst As String()) As String()
        Dim count As Integer = UBound(lst)

        For n As Integer = 0 To count
            lst(n) = lst(n).Trim
        Next
        For n As Integer = count To 0 Step -1
            If lst(n).Length > 0 Then Exit For
            ReDim Preserve lst(n - 1)
        Next

        Return lst
    End Function

    ''' <summary>
    ''' Сохраняет структуру криттера в pro-файл.
    ''' </summary>
    Friend Sub SaveCritterProData(ByVal proFile As String, ByRef CritterStruct As CritPro)
        If File.Exists(proFile) Then
            File.SetAttributes(proFile, FileAttributes.Normal Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)
        End If

        Dim sBuff As Integer() = ReverseSaveData(CritterStruct, Prototypes.CritterLen)
        If CritterStruct.DamageType > 6 Then
            Array.Resize(sBuff, Prototypes.CritterLen - 1)
            File.Delete(proFile) ' удаляем файл для перезаписи его размера.
        End If

        Dim fFile As Byte = FreeFile()
        FileOpen(fFile, proFile, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, sBuff)
        FileClose(fFile)

        If proRO Then File.SetAttributes(proFile, FileAttributes.ReadOnly Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)
    End Sub

    ''' <summary>
    ''' Получает данные из pro-файла криттера в структуре.
    ''' </summary>
    Friend Function LoadCritterProData(ByVal PathProFile As String, ByRef CritterStruct As CritPro) As Boolean
        Dim cProData(Prototypes.CritterLen - 1) As Integer  ' read f2 buffer
        Dim f1ProData(Prototypes.CritterLen - 2) As Integer ' read f1 buffer

        Dim fFile As Integer = FreeFile()
        Try
            FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            Dim file As New FileInfo(PathProFile)
            If file.Length = 412 Then
                FileGet(fFile, f1ProData)
                f1ProData.CopyTo(cProData, 0)
                cProData(Prototypes.CritterLen - 1) = &H7000000 'this index 7
                ProFiles.ReverseLoadData(cProData, CritterStruct)
            ElseIf file.Length = 416 Then
                FileGet(fFile, cProData)
                ProFiles.ReverseLoadData(cProData, CritterStruct)
            Else
                Throw New System.Exception
            End If
        Catch
            Return True 'for error
        Finally
            FileClose(fFile)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Помещает данные из pro-файла криттера в массив.
    ''' </summary>
    Friend Function LoadCritterProData(ByVal PathProFile As String, ByRef CrttrProData As Integer()) As Boolean
        Dim fFile As Byte = FreeFile()
        Dim f1ProData(Prototypes.CritterLen - 2) As Integer ' read f1 buffer

        PathProFile = DatFiles.CheckFile(PROTO_CRITTERS & PathProFile)

        Try
            FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            If FileSystem.GetFileInfo(PathProFile).Length = 412 Then
                FileGet(fFile, f1ProData)
                f1ProData.CopyTo(CrttrProData, 0)
                CrttrProData(Prototypes.CritterLen - 1) = -1
            Else
                FileGet(fFile, CrttrProData)
            End If

            For n = 0 To CrttrProData.Length - 1
                CrttrProData(n) = ProFiles.ReverseBytes(CrttrProData(n))
            Next
        Catch ex As Exception
            Return True
        Finally
            FileClose(fFile)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Сохраняет структуру предмета в pro-файл.
    ''' </summary>
    Friend Sub SaveItemProData(ByVal PathProFile As String, ByVal iType As Integer,
                              ByRef CommonItem As CmItemPro, ByRef WeaponItem As WpItemPro, ByRef ArmorItem As ArItemPro,
                              ByRef AmmoItem As AmItemPro, ByRef DrugItem As DgItemPro, ByRef MiscItem As McItemPro,
                              Optional ByRef ContanerItem As CnItemPro = Nothing, Optional ByRef KeyItem As kItemPro = Nothing)
        If File.Exists(PathProFile) Then
            File.SetAttributes(PathProFile, FileAttributes.Normal Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)
            File.Delete(PathProFile) ' удаляем файл для перезаписи его размера.
        End If

        Dim fFile As Integer = FreeFile()
        FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, ProFiles.ReverseSaveData(CommonItem, Prototypes.ItemCommonLen))
        FilePut(fFile, CommonItem.SoundID)
        Select Case iType
            Case ItemType.Weapon
                FilePut(fFile, ProFiles.ReverseSaveData(WeaponItem, Prototypes.ItemWeaponLen))
                FilePut(fFile, WeaponItem.wSoundID)
            Case ItemType.Armor
                FilePut(fFile, ProFiles.ReverseSaveData(ArmorItem, Prototypes.ItemArmorLen))
            Case ItemType.Drugs
                FilePut(fFile, ProFiles.ReverseSaveData(DrugItem, Prototypes.ItemDrugsLen))
            Case ItemType.Ammo
                FilePut(fFile, ProFiles.ReverseSaveData(AmmoItem, Prototypes.ItemAmmoLen))
            Case ItemType.Misc
                FilePut(fFile, ProFiles.ReverseSaveData(MiscItem, Prototypes.ItemMiscLen))
            Case ItemType.Container
                FilePut(fFile, ProFiles.ReverseSaveData(ContanerItem, Prototypes.ItemContLen))
            Case ItemType.Key
                FilePut(fFile, ProFiles.ReverseSaveData(KeyItem, Prototypes.ItemKeyLen))
        End Select
        FileClose(fFile)

        If proRO Then File.SetAttributes(PathProFile, FileAttributes.ReadOnly Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)
    End Sub

    ''' <summary>
    ''' Помещает данные из pro-файла предмета в структуру.
    ''' </summary>
    Friend Sub LoadItemProData(ByVal PathProFile As String, ByVal iType As Integer,
                              ByRef CommonItem As CmItemPro, ByRef WeaponItem As WpItemPro, ByRef ArmorItem As ArItemPro,
                              ByRef AmmoItem As AmItemPro, ByRef DrugItem As DgItemPro, ByRef MiscItem As McItemPro,
                              Optional ByRef ContanerItem As CnItemPro = Nothing, Optional ByRef KeyItem As kItemPro = Nothing)
        Dim cmProDataBuf(Prototypes.ItemCommonLen - 1) As Integer
        Dim fFile As Integer = FreeFile()

        FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
        FileGet(fFile, cmProDataBuf)
        ProFiles.ReverseLoadData(cmProDataBuf, CommonItem)
        FileGet(fFile, CommonItem.SoundID)

        Select Case iType
            Case ItemType.Weapon
                Dim wnProDataBuf(Prototypes.ItemWeaponLen - 1) As Integer
                FileGet(fFile, wnProDataBuf)
                ProFiles.ReverseLoadData(wnProDataBuf, WeaponItem)
                FileGet(fFile, WeaponItem.wSoundID)
            Case ItemType.Armor
                Dim arProDataBuf(Prototypes.ItemArmorLen - 1) As Integer
                FileGet(fFile, arProDataBuf)
                ProFiles.ReverseLoadData(arProDataBuf, ArmorItem)
            Case ItemType.Ammo
                Dim amProDataBuf(Prototypes.ItemAmmoLen - 1) As Integer
                FileGet(fFile, amProDataBuf)
                ProFiles.ReverseLoadData(amProDataBuf, AmmoItem)
            Case ItemType.Container
                FileGet(fFile, ContanerItem)
                ContanerItem.MaxSize = ProFiles.ReverseBytes(ContanerItem.MaxSize)
                ContanerItem.OpenFlags = ProFiles.ReverseBytes(ContanerItem.OpenFlags)
            Case ItemType.Drugs
                Dim drProDataBuf(Prototypes.ItemDrugsLen - 1) As Integer
                FileGet(fFile, drProDataBuf)
                ProFiles.ReverseLoadData(drProDataBuf, DrugItem)
            Case ItemType.Misc
                Dim msProDataBuf(Prototypes.ItemMiscLen - 1) As Integer
                FileGet(fFile, msProDataBuf)
                ProFiles.ReverseLoadData(msProDataBuf, MiscItem)
            Case ItemType.Key
                FileGet(fFile, KeyItem)
                KeyItem.Unknown = ProFiles.ReverseBytes(KeyItem.Unknown)
        End Select

        FileClose(fFile)
    End Sub

    Friend Sub ReverseLoadData(Of T As Structure)(ByRef buffer() As Integer, ByRef struct As T)
        For n = 0 To buffer.Length - 1
            buffer(n) = ReverseBytes(buffer(n))
        Next
        struct = CType(ConvertBytesToStruct(buffer, struct.GetType), T)
    End Sub

    Friend Function ReverseSaveData(ByVal struct As Object, ByVal isize As Integer) As Integer()
        Dim bSize As Integer = Marshal.SizeOf(struct)
        Dim bytes(bSize - 1) As Byte
        Dim buffer(isize - 1) As Integer

        ConvertStructToBytes(bytes, bSize, struct)
        Array.Reverse(bytes)

        For n As Integer = 0 To buffer.Length - 1
            bSize -= 4
            buffer(n) = BitConverter.ToInt32(bytes, bSize)
        Next

        Return buffer
    End Function

    ''' <summary>
    ''' Инвертирует значение в BigEndian и обратно.
    ''' </summary>
    Friend Function ReverseBytes(ByVal value As Integer) As Integer
        If value = 0 OrElse value = &HFFFFFFFF Then Return value

        Return (value << 24) Or
               (value And &HFF00) << 8 Or
               (value And &HFF0000) >> 8 Or
               (value >> 24) And &HFF
    End Function

    ''' <summary>
    ''' Преобразовывает структуру в массив.
    ''' </summary>
    Private Function ConvertStructToBytes(ByRef bytes() As Byte, ByVal bSize As Integer, ByVal struct As Object) As Byte()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(bSize)
        Marshal.StructureToPtr(struct, ptr, False)
        Marshal.Copy(ptr, bytes, 0, bSize)
        Marshal.FreeHGlobal(ptr)
        Return bytes
    End Function

    ''' <summary>
    ''' Преобразовывает массив в структуру.
    ''' </summary>
    Private Function ConvertBytesToStruct(ByVal Buff() As Integer, ByVal strcType As Type) As Object
        Dim mGC As GCHandle = GCHandle.Alloc(Buff, GCHandleType.Pinned)
        Dim obj As Object = Marshal.PtrToStructure(mGC.AddrOfPinnedObject, strcType)
        mGC.Free()
        Return obj
    End Function

End Module
