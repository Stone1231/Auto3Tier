Imports System.Text
Public Class Form8

    Private ReadOnly Property PageSize() As Integer
        Get
            If CheckBox5.Checked = True Then
                If IsNumeric(TextBox1.Text.Trim) = True Then
                    Return TextBox1.Text
                Else
                    Return 10
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    Public Sub LoadData()
        ComboBox1.DataSource = GetTableDv()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CodeType = "cs"

        'Dim intTId As Integer = TId
        'TId = ComboBox1.SelectedItem("TId")
        'Table_Initial()
        'GetUI()
        'GetCodeFile()
        'mdFile.classDT.Clear()
        'TId = intTId
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim intTId As Integer = TId
        TId = ComboBox1.SelectedItem("TId")
        Table_Initial()

        ComboBox2.DataSource = FunDv()

        mdFile.classDT.Clear()
        TId = intTId
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        CheckBox2.Visible = Not CheckBox1.Checked
        CheckBox3.Visible = Not CheckBox1.Checked
        CheckBox6.Visible = Not CheckBox1.Checked
        CheckBox7.Visible = Not CheckBox1.Checked
        Label8.Visible = Not CheckBox1.Checked

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        BindGV1()
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.ColumnIndex = 4 And e.RowIndex > -1 Then
            If IsNothing(DataGridView1.Rows(e.RowIndex).Cells("defVal").Value) Then
                DataGridView1.Rows(e.RowIndex).Cells("defVal").Value = ""
            End If
        End If
    End Sub

    Private Function FunDv() As DataTable
        Dim dt As DataTable = GetFunDv(TId, "List")
        Dim dr As DataRow = dt.NewRow

        dr("FId") = 0
        dr("TId") = TId
        dr("Name") = "GetAll"
        dr("SqlText") = ""
        dr("ReturnType") = "List"
        dr("Summary") = ""

        dt.Rows.Add(dr)

        Return dt
    End Function

    Private Sub BindGV1()
        DataGridView1.DataSource = GetFunParamDv(ComboBox2.SelectedItem("FId"))
        SetGV1()
    End Sub

    Private Sub SetGV1()
        Dim i As Integer
        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells("FromType").Value = "TextBox"
            DataGridView1.Rows(i).Cells("defVal").Value = GetUIDefValue(DataGridView1.Rows(i).Cells("ParamType").Value)
            DataGridView1.Rows(i).Cells("FromName").Value = DataGridView1.Rows(i).Cells("ParamName").Value
        Next
    End Sub

End Class