Public Class Hex_Form
    Dim HexView As New System.ComponentModel.Design.ByteViewer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParent(Me.Handle.ToInt32, Main_Form.SplitContainer1.Panel1.Handle.ToInt32)
        HexView.Left = 10
        HexView.Top = 22
        HexView.Height = 490
        HexView.Width = 425
        HexView.CellBorderStyle = TableLayoutPanelCellBorderStyle.None
        Controls.Add(HexView)
        GroupBox1.SendToBack()
    End Sub

    Friend Sub LoadHex(ByVal path As String, ByVal file As String)
        Me.Text = "HEX View: " & file
        HexView.SetFile(path & file) 'или .SetBytes
        If Me.Visible = False Then Me.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Process.Start(HEX_Path, Me.Tag)
    End Sub
End Class