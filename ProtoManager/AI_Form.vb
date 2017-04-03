Imports System.Drawing

Public Class AI_Form

    Private Const AIfile As String = "data\AI.txt"
    Private AIpath As String        'текущий путь до AI.txt

    Private AI_Packet(,) As String   'Список всех имен

    Private fReady As Boolean

    Private Const _width As Integer = 287
    Private Const _capform As String = "] AI Packet Editor"

    Friend Sub Initialize(ByVal AIPacket As Integer)
        ToolStripComboBox1.Items.AddRange(Items_NAME)
        AIpath = Check_File("data\aigenmsg.txt") & "\data\aigenmsg.txt"
        ComboBox8.Items.AddRange((From t In AI.GetAll_AIPacket(AIpath)).ToArray)
        ClearItem(ComboBox8)
        AIpath = Check_File("data\aibdymsg.txt") & "\data\aibdymsg.txt"
        ComboBox4.Items.AddRange((From t In AI.GetAll_AIPacket(AIpath)).ToArray)
        ClearItem(ComboBox4)
        '
        AIpath = Check_File(AIfile) & "\" & AIfile
        AI_Packet = AI.GetAll_AIPacket(AIpath)
        ComboBox0.Items.AddRange((From t In AI_Packet Take (AI_Packet.GetLength(1))).ToArray)
        '
        If AIPacket <> -1 Then
            For i As Int32 = 0 To AI_Packet.GetLength(1) - 1
                If AI.GetIniParam(AI_Packet(0, i), "packet_num", AIpath) = AIPacket Then
                    ComboBox0.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            ComboBox0.SelectedIndex = 0
        End If
        '
        Me.Text = "[" & ComboBox0.Text & _capform
        Me.Width -= _width
        Me.Show()
    End Sub

    ' Рекурсивный перебор контролов класса формы
    Private Sub FormControl(ByRef сntr As Control, ByRef Section As String)
        Dim KeyValue As String
        For Each _control As Control In сntr.Controls
            If TypeOf _control Is NumericUpDown Then
                _control.Text = AI.GetIniParam(Section, _control.Tag, AIpath)
            ElseIf (TypeOf _control Is ComboBox) And _control.Tag <> Nothing Then
                KeyValue = AI.GetIniStringParam(Section, _control.Tag, AIpath)
                If KeyValue = "<Unknown>" Then _control.BackColor = Color.Linen Else _control.BackColor = Color.White
                _control.Text = KeyValue
            ElseIf TypeOf _control Is GroupBox Then
                FormControl(_control, Section)
            End If
        Next
    End Sub

    Private Sub ClearItem(ByVal cntrl As ComboBox)
        Dim c As Integer = cntrl.Items.Count - 1
        For i = c To (c / 2) Step -1
            cntrl.Items.RemoveAt(i)
        Next
    End Sub

    Private Sub Taunts_Show(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If sender.tag = 0 Then
            Me.Width += _width
            sender.tag = 1
            sender.Image = My.Resources.LeftArrow
        Else
            Me.Width -= _width
            sender.tag = 0
            sender.Image = My.Resources.RightArrow
        End If
    End Sub

    Private Sub Select_AI_Packet(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox0.SelectedIndexChanged
        fReady = False
        Dim Section As String = ComboBox0.Text
        Me.Text = "[" & Section & _capform
        SetControlValue(Section)
        fReady = True
    End Sub

    Private Sub SetControlValue(ByRef Section As String)
        FormControl(Me, Section)
        'chem_primary_desire
        ListView1.Items.Clear()
        Dim drug_lst() As String = Split(AI.GetIniStringParam(Section, "chem_primary_desire", AIpath), ",")
        If drug_lst(0) <> "<Unknown>" And drug_lst(0) <> "-1" Then
            For i = 0 To drug_lst.GetLength(0) - 1
                ListView1.Items.Add(drug_lst(i))
                ListView1.Items(i).SubItems.Add(Items_NAME(drug_lst(i) - 1))
            Next
        Else
            ListView1.Items.Add("-1")
            ListView1.Items(0).SubItems.Add(drug_lst(0)) '"<Unknown>"
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim AITxtfrm As New AI_TextForm
        AITxtfrm.Owner = Me
        AITxtfrm.Text &= ComboBox0.Text
        AITxtfrm.Initialize(CInt(AI_Packet(1, ComboBox0.SelectedIndex)), CInt(AI_Packet(1, ComboBox0.SelectedIndex + 1)), AIpath)
        AITxtfrm.Show()
        Button1.Enabled = True
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        AddDrugsToolStripMenuItem.Enabled = True
        ContextMenuStrip1.Focus()
    End Sub

    Private Sub AddDrugsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddDrugsToolStripMenuItem.Click
        ListView1.Items.Add(ToolStripComboBox1.SelectedIndex + 1)
        ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(Items_NAME(ToolStripComboBox1.SelectedIndex))
    End Sub

    Private Sub ListView1_AfterLabelEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles ListView1.AfterLabelEdit
        If e.Label = Nothing Or e.CancelEdit Then Exit Sub
        Try
            ListView1.Items(e.Item).SubItems(1).Text = Items_NAME(CInt(e.Label) - 1)
        Catch
            ListView1.Items(e.Item).SubItems(1).Text = "<Error PID>"
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        On Error Resume Next
        ListView1.Items.RemoveAt(ListView1.FocusedItem.Index)
    End Sub

    Private Sub MoveUpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveUpToolStripMenuItem.Click
        Try
            Dim sel_index As Integer = ListView1.FocusedItem.Index
            Dim l_sel As ListViewItem = ListView1.Items(sel_index).Clone
            Dim l_up As ListViewItem = ListView1.Items(sel_index - 1).Clone
            ListView1.Items(sel_index - 1).Text = l_sel.Text
            ListView1.Items(sel_index - 1).SubItems(1).Text = l_sel.SubItems(1).Text
            ListView1.Items(sel_index).Text = l_up.Text
            ListView1.Items(sel_index).SubItems(1).Text = l_up.SubItems(1).Text
            ListView1.Items(sel_index - 1).Selected = True
        Catch
        End Try
    End Sub

    Private Sub MoveDownToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveDownToolStripMenuItem.Click
        Try
            Dim sel_index As Integer = ListView1.FocusedItem.Index
            Dim l_sel As ListViewItem = ListView1.Items(sel_index).Clone
            Dim l_down As ListViewItem = ListView1.Items(sel_index + 1).Clone
            ListView1.Items(sel_index + 1).Text = l_sel.Text
            ListView1.Items(sel_index + 1).SubItems(1).Text = l_sel.SubItems(1).Text
            ListView1.Items(sel_index).Text = l_down.Text
            ListView1.Items(sel_index).SubItems(1).Text = l_down.SubItems(1).Text
            ListView1.Items(sel_index + 1).Selected = True
        Catch
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Select_AI_Packet(Nothing, Nothing)
    End Sub

End Class