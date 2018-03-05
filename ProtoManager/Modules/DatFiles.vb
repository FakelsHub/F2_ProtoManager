Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Module DatFiles

    'Declare zlib functions "Compress" and "Uncompress" for compressing Byte Arrays
    '    <DllImport("zlib.DLL", EntryPoint:="compress")> _
    '    Private Function CompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    ' Leave function empty - DLLImport attribute forwards calls to CompressByteArray to compress in zlib.dLL
    '    End Function
    '    <DllImport("zlib.DLL", EntryPoint:="uncompress")> _
    '    Private Function UncompressByteArray(ByVal dest As Byte(), ByRef destLen As Integer, ByVal src As Byte(), ByVal srcLen As Integer) As Integer
    ' Leave function empty - DLLImport attribute forwards calls to UnCompressByteArray to Uncompress in zlib.dLL
    '   End Function

    Friend Const MasterDAT As String = "\master.dat"
    Friend Const CritterDAT As String = "\critter.dat"
    Friend Const DIR_DATA As String = "\data" 'for test "\MyTestData"

    Friend Const ART_CRITTERS As String = "\art\critters\"
    Friend Const ART_INVEN As String = "\art\inven\"
    Friend Const ART_ITEMS As String = "\art\items\"

    Friend Const PROTO_CRITTERS As String = "\proto\critters\"
    Friend Const PROTO_ITEMS As String = "\proto\items\"

    ' to List path file
    Friend Const itemsLstPath As String = "\proto\items\items.lst"
    Friend Const crittersLstPath As String = "\proto\critters\critters.lst"
    Friend Const scriptsLstPath As String = "\scripts\scripts.lst"
    Friend Const miscLstPath As String = "\proto\misc\misc.lst"

    Friend Const artCrittersLstPath As String = "\art\critters\critters.lst"
    Friend Const artItemsLstPath As String = "\art\items\items.lst"
    Friend Const artInvenLstPath As String = "\art\inven\inven.lst"

    'Friend Const proCritMsgPath As String = "\Text\English\Game\pro_crit.msg"
    'Friend Const proItemMsgPath As String = "\Text\English\Game\pro_item.msg"

    ''' <summary>
    ''' Проверить наличее файла и возвратить путь к нему, если такого файла не найдено то извлечь его из Dat архива.
    ''' full = false, возвращает короткий путь к файлу.
    ''' </summary>
    Friend Function CheckFile(ByVal pFile As String, Optional ByVal full As Boolean = True,
                              Optional ByVal сDat As Boolean = False, Optional ByVal unpack As Boolean = True) As String
        Dim fPath As String = String.Concat(SaveMOD_Path, pFile)
        If File.Exists(fPath) Then
            Return If(full, fPath, SaveMOD_Path)
        End If
        fPath = String.Concat(GameDATA_Path & pFile)
        If File.Exists(fPath) Then
            Return If(full, fPath, GameDATA_Path)
        End If

        Dim fmDat As String = String.Concat(Game_Path, MasterDAT, pFile)
        Dim fcDat As String = String.Concat(Game_Path, CritterDAT, pFile)
        If File.Exists(fmDat) OrElse File.Exists(fcDat) Then
            If сDat Then
                Return If(full, fcDat, String.Concat(Game_Path, CritterDAT)) 'папка
            Else
                Return If(full, fmDat, String.Concat(Game_Path, MasterDAT)) 'папка
            End If
        Else
            fPath = String.Concat(Cache_Patch, pFile)
            If unpack AndAlso Not (File.Exists(fPath)) Then UnPackFile(pFile, сDat)
            Return If(full, fPath, Cache_Patch)
        End If
    End Function

    Friend Function CheckDirFile(ByVal pFile As String, ByVal cache As Boolean, Optional ByVal сDat As Boolean = False) As Boolean
        'Check save folder
        Dim fPath As String = String.Concat(SaveMOD_Path, pFile)
        If File.Exists(fPath) Then Return True

        'Check data folder
        fPath = String.Concat(GameDATA_Path & pFile)
        If File.Exists(fPath) Then Return True

        'Check .dat folder
        Dim fmDat As String = String.Concat(Game_Path, MasterDAT, pFile)
        Dim fcDat As String = String.Concat(Game_Path, CritterDAT, pFile)
        If File.Exists(fmDat) OrElse File.Exists(fcDat) Then
            If сDat Then
                Return True
            Else
                Return True
            End If
        ElseIf cache Then
            'Check cache folder
            fPath = String.Concat(Cache_Patch, pFile)
            If File.Exists(fPath) Then Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Проверить содержится ли указанный файл в кэш папке,
    ''' если его нет то извлечь его.
    ''' </summary>
    Friend Function UnDatFile(ByVal pFile As String, ByVal size As Integer) As Boolean
        Dim cPathFile As String = Cache_Patch & pFile

        If File.Exists(cPathFile) AndAlso FileSystem.GetFileInfo(cPathFile).Length = size Then Return True
        UnPackFile(pFile)
        If File.Exists(cPathFile) Then Return True 'successfully unpack

        Return False
    End Function

    'Извлечь требуемый файл в кэш папку
    Private Sub UnPackFile(ByVal pFile As String, Optional ByVal сDat As Boolean = False)
        Main.PrintLog("Extracting file: " & pFile)
        Dim fileDAT As String
        If сDat Then
            fileDAT = CritterDAT
        Else
            fileDAT = MasterDAT
        End If
        Shell(WorkAppDIR & "\dat2.exe x -d cache """ & Game_Path & fileDAT & """ " & pFile.Remove(0, 1), AppWinStyle.Hide, True)
    End Sub

    ''' <summary>
    ''' Получить frm файл криттера в gif формате
    ''' </summary>
    Friend Sub CritterFrmGif(ByRef FrmFile As String)
        Dim checkFile As String = ART_CRITTERS & FrmFile & "aa.frm"
        Dim cPath As String = SaveMOD_Path & checkFile

        ExtractConvertFRM(cPath, checkFile, FrmFile & "aa", CritterDAT)

        Dim gifFile As String = Path.ChangeExtension(checkFile, ".gif") 'ART_CRITTERS & FrmFile & "aa.gif"

        On Error Resume Next
        If File.Exists(WorkAppDIR & gifFile) Then
            FileSystem.MoveFile(WorkAppDIR & gifFile, Cache_Patch & gifFile)
        Else
            FileSystem.MoveFile(WorkAppDIR & ART_CRITTERS & FrmFile & "aa_sw.gif", Cache_Patch & gifFile)
        End If
        Directory.Delete(WorkAppDIR & "\art", True)
    End Sub

    ''' <summary>
    ''' Получить frm файл предмета в gif формате
    ''' </summary>
    Friend Sub ItemFrmGif(ByVal iPath As String, ByVal FrmFile As String)
        Dim checkFile As String = "\art\" & iPath & FrmFile & ".frm"
        Dim cPath As String = SaveMOD_Path & checkFile

        ExtractConvertFRM(cPath, checkFile, FrmFile, MasterDAT)

        Dim gifFile As String = Path.ChangeExtension(checkFile, ".gif")

        On Error Resume Next
        If File.Exists(WorkAppDIR & gifFile) Then
            FileSystem.MoveFile(WorkAppDIR & gifFile, Cache_Patch & gifFile)
        Else
            FileSystem.MoveFile(WorkAppDIR & "\art\" & iPath & FrmFile & "_ne.gif", Cache_Patch & gifFile)
        End If
        Directory.Delete(WorkAppDIR & "\art", True)
    End Sub

    ''' <summary>
    ''' Преобразовать frm файл в gif формат
    ''' </summary>
    Private Sub ExtractConvertFRM(ByVal cPath As String, ByVal checkFile As String, ByVal FrmFile As String, ByVal nameDAT As String)
        If Not (File.Exists(cPath)) Then
            cPath = GameDATA_Path & checkFile
            If Not (File.Exists(cPath)) Then
                cPath = Game_Path & nameDAT & checkFile
                If Not (File.Exists(cPath)) Then
                    'Извлечь требуемый файл
                    Shell(WorkAppDIR & "\frm2gif.exe -d -f """ & Game_Path & nameDAT & """ -p color.pal " & FrmFile & ".frm", AppWinStyle.Hide, True, 2000)
                    Exit Sub
                End If
            End If
        End If

        FileSystem.CopyFile(cPath, WorkAppDIR & checkFile)
        Shell(WorkAppDIR & "\frm2gif.exe -p color.pal ." & checkFile, AppWinStyle.Hide, True)
    End Sub

End Module
