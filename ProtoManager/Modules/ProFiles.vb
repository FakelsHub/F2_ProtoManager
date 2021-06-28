Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Imports Prototypes
Imports Enums

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
    Friend Function GetProItemsDataIDs(ByRef ProFile As String, ByVal n As Integer) As Integer
        Dim NameID, pID As Integer
        Dim type As ItemType = ItemType.Unknown

        Dim cPath As String = DatFiles.CheckFile(PROTO_ITEMS & ProFile)

        Try
            Using brFile As New BinaryReader(File.Open(cPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                pID = brFile.ReadInt32()
                NameID = brFile.ReadInt32()
                brFile.BaseStream.Seek(Prototypes.offsetISubType, SeekOrigin.Begin)
                type = CType(ReverseBytes(brFile.ReadInt32()), ItemType)
            End Using
        Catch ex As EndOfStreamException
            NameID = 0
            type = ItemType.Unknown
            MsgBox("The file is in an incorrect format or damaged." & vbLf & cPath)
        Catch ex As Exception
            type = ItemType.Unknown
        End Try

        If type >= 0 AndAlso type < ItemType.Unknown Then
            Items_LST(n).itemType = type
            Items_LST(n).PID = ReverseBytes(pID)
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

        Critter_LST(nPro).FID = FID
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
        Critter_LST(nPro).FID = FID

        If FID = -1 Then Return Nothing
        FID -= &H1000000I

        Return Critters_FRM(FID).ToLower
    End Function

    ''' <summary>
    ''' Возвращает номер Description ID из про-файла криттера.
    ''' </summary>
    Friend Function GetProCritterDataIDs(ByRef crtList As CrittersLst) As Integer
        Dim nameID, pID, fID As Integer
        Dim cPath = DatFiles.CheckFile(PROTO_CRITTERS & crtList.proFile)

        Dim fFile As Integer = FreeFile()
        Try
            FileOpen(fFile, cPath, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            FileGet(fFile, pID)
            FileGet(fFile, nameID)
            FileGet(fFile, fID)
        Catch
            Return 0
        Finally
            FileClose(fFile)
            crtList.PID = ReverseBytes(pID)
            crtList.FID = ReverseBytes(fID)
        End Try

        Return ReverseBytes(nameID)
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

        Dim fFile As Integer = FreeFile()
        FileOpen(fFile, proFile, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, sBuff)
        FileClose(fFile)

        If proRO Then File.SetAttributes(proFile, FileAttributes.ReadOnly Or FileAttributes.Archive Or FileAttributes.NotContentIndexed)
    End Sub

    ''' <summary>
    ''' Получает данные из pro-файла криттера в структуре.
    ''' </summary>
    Friend Function LoadCritterProData(ByVal PathProFile As String, ByRef CritterStruct As CritPro) As Boolean
        Dim critterProData(Prototypes.CritterLen - 1) As Integer  ' read f2 buffer

        Dim fFile As Integer = FreeFile()
        Try
            FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            Dim file As New FileInfo(PathProFile)

            If file.Length = 412 Then
                Dim proData(Prototypes.CritterLen - 2) As Integer ' read f1 buffer
                FileGet(fFile, proData)
                proData.CopyTo(critterProData, 0)
                critterProData(Prototypes.CritterLen - 1) = &H7000000 ' set index 7

            ElseIf file.Length = 416 Then
                FileGet(fFile, critterProData)
            Else
                Throw New System.Exception
            End If
        Catch
            Return True 'for error
        Finally
            ProFiles.ReverseLoadData(critterProData, CritterStruct)
            FileClose(fFile)
        End Try

        Return False
    End Function

    ''' <summary>
    ''' Помещает данные из pro-файла криттера в массив.
    ''' </summary>
    Friend Function LoadCritterProData(ByVal PathProFile As String, ByRef crtProData As Integer()) As Boolean
        Dim fFile As Integer = FreeFile()

        PathProFile = DatFiles.CheckFile(PROTO_CRITTERS & PathProFile)

        Try
            FileOpen(fFile, PathProFile, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
            If FileSystem.GetFileInfo(PathProFile).Length = 412 Then
                Dim f1ProData(Prototypes.CritterLen - 2) As Integer ' read f1 buffer
                FileGet(fFile, f1ProData)
                f1ProData.CopyTo(crtProData, 0)
                crtProData(Prototypes.CritterLen - 1) = -1
            Else
                FileGet(fFile, crtProData)
            End If

            For n = 0 To crtProData.Length - 1
                crtProData(n) = ProFiles.ReverseBytes(crtProData(n))
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
    Friend Sub SaveItemProData(ByVal pathProFile As String, ByVal item As IPrototype)
        If File.Exists(pathProFile) Then
            File.SetAttributes(pathProFile, FileAttributes.Normal Or FileAttributes.NotContentIndexed)
        End If

        item.Save(pathProFile)

        If proRO Then File.SetAttributes(pathProFile, FileAttributes.ReadOnly Or FileAttributes.NotContentIndexed)
    End Sub

    Friend Sub ReverseLoadData(Of T As Structure)(ByRef buffer() As Integer, ByRef struct As T)
        For n = 0 To buffer.Length - 1
            buffer(n) = ReverseBytes(buffer(n))
        Next
        struct = CType(ConvertBytesToStruct(buffer, struct.GetType), T)
    End Sub

    Friend Sub ReverseLoadData(Of T As Structure)(ByRef buffer() As Byte, ByRef struct As T)
        ReverseBytes(buffer, buffer.Length)

        Dim mGC As GCHandle = GCHandle.Alloc(buffer, GCHandleType.Pinned)
        struct = CType(Marshal.PtrToStructure(mGC.AddrOfPinnedObject, struct.GetType), T)
        mGC.Free()
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

    Friend Function SaveDataReverse(Of T As Structure)(ByVal struct As T) As Byte()
        Dim bSize As Integer = Marshal.SizeOf(struct)
        Dim buffer(bSize - 1) As Byte
        ConvertStructToBytes(buffer, bSize, struct)
        ReverseBytes(buffer, bSize And Not (&H3))
        Return buffer
    End Function

    ''' <summary>
    ''' Инвертирует значение в BigEndian и обратно.
    ''' </summary>
    Friend Function ReverseBytes(ByVal value As Integer) As Integer
        If value = 0 OrElse value = -1 Then Return value

        Return (value << 24) Or
               (value And &HFF00) << 8 Or
               (value And &HFF0000) >> 8 Or
               (value >> 24) And &HFF
    End Function

    Private Sub ReverseBytes(ByRef bytes() As Byte, ByVal length As Integer)
        While (length > 0)
            length -= 4
            Array.Reverse(bytes, length, 4)
        End While

        'Dim n = 0
        'Do
        '    Dim i = n + 3       ' i = 3
        '    Dim v As Byte = bytes(i)
        '    bytes(i) = bytes(n) ' [3] <- [0]
        '    bytes(n) = v        ' [0] <- [3]
        '    i = n + 1           ' i = 1
        '    n += 2              ' n = 2
        '    v = bytes(n)
        '    bytes(n) = bytes(i) ' [2] <- [1]
        '    bytes(i) = v        ' [1] <- [2]
        '    n += 2              ' n = 4
        'Loop While (n < count)
    End Sub

    ''' <summary>
    ''' Преобразовывает структуру в массив.
    ''' </summary>
    Private Sub ConvertStructToBytes(ByRef bytes() As Byte, ByVal bSize As Integer, ByVal struct As Object)
        Dim ptr As IntPtr = Marshal.AllocHGlobal(bSize)
        Marshal.StructureToPtr(struct, ptr, False)
        Marshal.Copy(ptr, bytes, 0, bSize)
        Marshal.FreeHGlobal(ptr)
    End Sub

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
