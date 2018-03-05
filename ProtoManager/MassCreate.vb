Imports System.IO
Imports Prototypes

Public Class MassCreate

    Private cmmn As CommonPro
    Private wall As WallsPro
    Private tile As TilesPro

    Private Path As String
    Private proLst As List(Of String)

    Friend Sub New()
        InitializeComponent()
        FBDialog.SelectedPath = Game_Path
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        Me.ShowDialog()
    End Sub

    Private Sub StartCreate(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim data As Integer()
        Dim PID As Integer, FIDnum As Integer = NumericUpDown3.Value
        Dim MaxCount As Integer = NumericUpDown1.Value + NumericUpDown2.Value

        If FBDialog.ShowDialog = DialogResult.OK Then
            Path = FBDialog.SelectedPath + "\"
        Else
            Exit Sub
        End If

        proLst = New List(Of String)

        If ComboBox1.SelectedIndex = 0 Then
            proLst.Add("TILES")
            tile.Unknown = -1
            tile.MaterialID = ComboBox2.SelectedIndex
            data = ProFiles.ReverseSaveData(tile, 4)
            PID = &H4000000
        Else
            proLst.Add("WALLS")
            wall.FalgsExt = &H2000
            wall.ScriptID = -1
            wall.MaterialID = ComboBox2.SelectedIndex
            data = ProFiles.ReverseSaveData(wall, 6)
            PID = &H3000000
        End If

        Progress_Form.ShowProgressBar(NumericUpDown2.Value)

        For ProNum As Integer = NumericUpDown1.Value To MaxCount
            ProDataSave(ProNum, PID, FIDnum, data)
            FIDnum += 1
            Progress_Form.ProgressBar1.Value = ProNum - NumericUpDown1.Value
        Next
        Progress_Form.Close()

        If PID = &H4000000 AndAlso MessageBox.Show("Create 'End of Group' proto tiles?", "", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            ProDataSave(MaxCount + 1, PID, 21, data)
        End If

        File.WriteAllLines(Path & "Creates.lst", proLst)

        Process.Start("explorer", Path)
    End Sub

    Private Sub ProDataSave(ByVal ProNum As Integer, ByVal PID As Integer, ByVal FIDnum As Integer, ByRef Data As Integer())
        cmmn.ProtoID = (PID + ProNum)
        cmmn.DescID = (ProNum * 100)
        cmmn.FrmID = ((PID - 1) + FIDnum)

        Dim file As String = Format(ProNum, "00000000"".pro""")
        Dim fFile As Integer = FreeFile()
        FileOpen(fFile, Path & file, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, ProFiles.ReverseSaveData(cmmn, 3))
        FilePut(fFile, Data)
        FileClose(fFile)

        proLst.Add(file)
    End Sub

End Class