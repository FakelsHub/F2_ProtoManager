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
    Private Const CritterDAT As String = "\critter.dat"
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
    ''' Проверить наличее файла и возвратить путь к нему, 
    ''' если такого файла не найдено то извлечь его из Dat архива.
    ''' </summary>
    Friend Function CheckFile(ByRef pFile As String, Optional ByRef сDat As Boolean = False, Optional ByRef unpack As Boolean = True) As String
        If File.Exists(SaveMOD_Path & pFile) Then
            Return SaveMOD_Path
        ElseIf File.Exists(GameDATA_Path & pFile) Then
            Return GameDATA_Path
        ElseIf File.Exists(Game_Path & MasterDAT & pFile) OrElse File.Exists(Game_Path & CritterDAT & pFile) Then
            If сDat Then
                Return Game_Path & CritterDAT 'папка
            Else
                Return Game_Path & MasterDAT  'папка
            End If
        Else
            If Not (File.Exists(Cache_Patch & pFile)) And unpack Then UnPackFile(pFile, сDat)
        Return Cache_Patch
        End If
    End Function

    ''' <summary>
    ''' Проверить содержится ли указанный файл в кэш папке,
    ''' если его нет то извлечь его.
    ''' </summary>
    Friend Function UnDatFile(ByRef pFile As String, ByVal size As Integer) As Boolean
        If File.Exists(Cache_Patch & pFile) AndAlso FileSystem.GetFileInfo(Cache_Patch & pFile).Length = size Then Return True
        UnPackFile(pFile)
        If File.Exists(Cache_Patch & pFile) Then Return True
        Return False
    End Function

    'Извлечь требуемый файл в кэш папку
    Private Sub UnPackFile(ByVal pFile As String, Optional ByVal сDat As Boolean = False)
        Main_Form.TextBox1.Text = "Extracting file: " & pFile & vbCrLf & Main_Form.TextBox1.Text
        Dim fileDAT As String
        If сDat Then
            fileDAT = CritterDAT
        Else
            fileDAT = MasterDAT
        End If
        Shell(WorkAppDIR & "\dat2.exe x -d cache """ & Game_Path & fileDAT & """ " & pFile.Remove(0, 1), AppWinStyle.Hide, True)
    End Sub

    ''' <summary>
    ''' Преобразовать frm файл криттера в gif формат
    ''' </summary>
    Friend Sub CritterFrmGif(ByRef FrmFile As String)
        Dim checkFile As String = ART_CRITTERS & FrmFile & "aa.frm"
        Dim cPath As String = SaveMOD_Path & checkFile

        If Not (File.Exists(cPath)) Then
            cPath = GameDATA_Path & checkFile
            If Not (File.Exists(cPath)) Then
                cPath = Game_Path & CritterDAT & checkFile
                If Not (File.Exists(cPath)) Then
                    'Извлечь требуемый файл
                    Shell(WorkAppDIR & "\frm2gif.exe -d -f """ & Game_Path & CritterDAT & """ -p color.pal " & FrmFile & "aa.frm", AppWinStyle.Hide, True)
                    GoTo SKIP
                End If
            End If
        End If
        FileSystem.CopyFile(cPath, WorkAppDIR & checkFile)
        Shell(WorkAppDIR & "\frm2gif.exe -p color.pal ." & checkFile, AppWinStyle.Hide, True)
SKIP:
        Dim gifFile As String = Path.ChangeExtension(checkFile, ".gif") 'ART_CRITTERS & FrmFile & "aa.gif"
        If File.Exists(WorkAppDIR & gifFile) Then
            FileSystem.MoveFile(WorkAppDIR & gifFile, Cache_Patch & gifFile)
        Else
            FileSystem.MoveFile(WorkAppDIR & ART_CRITTERS & FrmFile & "aa_sw.gif", Cache_Patch & gifFile)
        End If
        Directory.Delete(WorkAppDIR & "\art", True)
    End Sub

    ''' <summary>
    ''' Преобразовать frm файл предмета в gif формат
    ''' </summary>
    Friend Sub ItemFrmGif(ByVal iPath As String, ByVal FrmFile As String)
        Dim checkFile As String = "\art\" & iPath & FrmFile & ".frm"
        Dim cPath As String = SaveMOD_Path & checkFile

        If Not (File.Exists(cPath)) Then
            cPath = GameDATA_Path & checkFile
            If Not (File.Exists(cPath)) Then
                cPath = Game_Path & MasterDAT & checkFile
                If Not (File.Exists(cPath)) Then
                    'Извлечь требуемый файл
                    Shell(WorkAppDIR & "\frm2gif.exe -d -f """ & Game_Path & MasterDAT & """ -p color.pal " & FrmFile & ".frm", AppWinStyle.Hide, True)
                    GoTo SKIP
                End If
            End If
        End If
        FileSystem.CopyFile(cPath, WorkAppDIR & checkFile)
        Shell(WorkAppDIR & "\frm2gif.exe -p color.pal ." & checkFile, AppWinStyle.Hide, True)
SKIP:
        Dim gifFile As String = Path.ChangeExtension(checkFile, ".gif")
        If File.Exists(WorkAppDIR & gifFile) Then
            FileSystem.MoveFile(WorkAppDIR & gifFile, Cache_Patch & gifFile)
        Else
            FileSystem.MoveFile(WorkAppDIR & "\art\" & iPath & FrmFile & "_ne.gif", Cache_Patch & gifFile)
        End If
        Directory.Delete(WorkAppDIR & "\art", True)
    End Sub

End Module
