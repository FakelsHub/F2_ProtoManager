﻿Imports System.IO
Imports Prototypes

Public Class MassCreate

    Private common As CommonProto
    Private wall As WallsProto
    Private tile As TilesProto

    Private savePath As String

    Friend Sub New()
        InitializeComponent()

        FBDialog.SelectedPath = Game_Path
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0

        Me.ShowDialog()
    End Sub

    Private Sub StartCreate(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        If FBDialog.ShowDialog = DialogResult.OK Then
            savePath = FBDialog.SelectedPath + "\"
        Else
            Exit Sub
        End If

        Dim PID As Integer, FID As Integer = CInt(NumericUpDown3.Value)
        Dim MaxCount As Integer = CInt(NumericUpDown1.Value + NumericUpDown2.Value)

        Dim data As Integer()
        Dim proLst = New List(Of String)

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

        Progress_Form.ShowProgressBar(CInt(NumericUpDown2.Value))

        For ProNum As Integer = CInt(NumericUpDown1.Value) To MaxCount
            ProDataSave(ProNum, PID, FID, data, proLst)
            FID += 1
            Progress_Form.ProgressBar1.Value = ProNum - CInt(NumericUpDown1.Value)
        Next
        Progress_Form.Close()

        If PID = &H4000000 AndAlso MessageBox.Show("Create 'End of Group' proto tiles?", "", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            ProDataSave(MaxCount + 1, PID, 21, data, proLst)
        End If

        File.WriteAllLines(savePath & "Creates.lst", proLst)

        Process.Start("explorer", savePath)
    End Sub

    Private Sub ProDataSave(ByVal proNum As Integer, ByVal pid As Integer, ByVal fid As Integer, ByRef data As Integer(), proLst As List(Of String))
        common.ProtoID = (pid + proNum)
        common.DescID = (proNum * 100)
        common.FrmID = ((pid - 1) + fid)

        Dim file As String = Format(proNum, "00000000"".pro""")

        Dim fFile As Integer = FreeFile()
        FileOpen(fFile, savePath & file, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        FilePut(fFile, ProFiles.ReverseSaveData(common, 3))
        FilePut(fFile, data)
        FileClose(fFile)

        proLst.Add(file)
    End Sub

End Class