Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Imports Prototypes

Module ProFiles

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
            iFID -= &H7000000
            lstName = Iven_FRM(iFID)
        Else
            If FID = -1 Then Return Nothing
            lstName = Items_FRM(FID)
            Inventory = False
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
    Friend Function GetProItemsNameID(ByRef ProFile As String, ByRef n As Integer) As Integer
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

    Friend Sub ReverseLoadData(Of T As Structure)(ByRef buffer() As Integer, ByRef Struct As T)
        For n = 0 To buffer.Length - 1
            buffer(n) = ReverseBytes(buffer(n))
        Next
        Struct = fnBytesToStruct(buffer, Struct.GetType)
    End Sub

    Friend Function ReverseSaveData(ByVal Struct As Object, ByVal isize As Integer) As Integer()
        Dim bsize As Integer = Marshal.SizeOf(Struct)
        Dim bytes(bsize - 1) As Byte
        Dim buffer(isize - 1) As Integer

        fnStructToBytes(bytes, bsize, Struct)
        For n = 0 To bytes.Length - 1 Step 4
            If (n / 4) >= buffer.Length Then Exit For
            Dim value As Integer = BitConverter.ToInt32(bytes, n)
            If value = 0 OrElse value = -1 Then
                buffer(n / 4) = value
                Continue For
            End If
            buffer(n / 4) = ((bytes(n) And &HFF) << 24) + ((bytes(n + 1) And &HFF) << 16) + ((bytes(n + 2) And &HFF) << 8) + (bytes(n + 3) And &HFF)
        Next

        Return buffer
    End Function

    ''' <summary>
    ''' Инвертирует значение в BigEndian и обратно.
    ''' </summary>
    Friend Function ReverseBytes(ByVal Value As Integer) As Integer
        If Value = 0 OrElse Value = -1 Then Return Value
        Dim bytes() As Byte = BitConverter.GetBytes(Value)
        Array.Reverse(bytes)
        Return BitConverter.ToInt32(bytes, 0)
        'Return (Value And &HFF) << 24 Or (Value And &HFF00) << 8 Or (Value And &HFF0000) >> 8 Or (Value And &HFF000000) >> 24
    End Function

    ''' <summary>
    ''' Преобразовывает структуру в массив.
    ''' </summary>
    Private Function fnStructToBytes(ByRef bytes() As Byte, ByVal bsize As Integer, ByVal Struct As Object) As Byte()
        Dim Ptr As IntPtr = Marshal.AllocHGlobal(bsize)
        Marshal.StructureToPtr(Struct, Ptr, False)
        Marshal.Copy(Ptr, bytes, 0, bsize)
        Marshal.FreeHGlobal(Ptr)
        Return bytes
    End Function

    ''' <summary>
    ''' Преобразовывает массив в структуру.
    ''' </summary>
    Private Function fnBytesToStruct(ByVal Buff() As Integer, ByVal StrcType As Type) As Object
        Dim MyGC As GCHandle = GCHandle.Alloc(Buff, GCHandleType.Pinned)
        Dim Obj As Object = Marshal.PtrToStructure(MyGC.AddrOfPinnedObject, StrcType)
        MyGC.Free()
        Return Obj
    End Function

End Module
