Friend Class Progress_Form

    Friend Sub ShowProgressBar(ByVal maxValue As Integer)
        'Progress_Form.MdiParent = Main_Form Main_Form.SplitContainer1.Panel1.Handle.ToInt32
        SetParent(Me.Handle.ToInt32, Main_Form.Handle.ToInt32)
        Me.SetDesktopLocation(CInt(Main_Form.Width / 4), CInt(Main_Form.Height / 2.25))

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = maxValue

        Me.Show()
        Application.DoEvents()
    End Sub

End Class