Imports System.Text
Public Class Form2
    Private mFId As Integer
    Private mCFId As Integer
    Private Property FId() As Integer
        Get
            Return mFId
        End Get
        Set(ByVal value As Integer)
            mFId = value
        End Set
    End Property
    Private Property CFId() As Integer
        Get
            Return mCFId
        End Get
        Set(ByVal value As Integer)
            mCFId = value
        End Set
    End Property

    Public Sub LoadData()

        FId = -1
        CFId = -1

        Dim dv As DataView = GetTableDv.DefaultView
        Dim i As Integer

        For i = 0 To dv.Count - 1
            If TableSourceInfo(dv(i)("Name")).DefaultView.Count = 0 Then
                'DelTable(dv(i)("TId"))
                mdDel.DeleteTb(dv(i)("TId"))
            End If
        Next

        Dim dvSource As DataView = TablesSource.DefaultView
        For i = 0 To dvSource.Count - 1
            If GetTableInfoByName(dvSource(i)("TABLE_NAME")).DefaultView.Count = 0 Then
                AddTable(dvSource(i)("TABLE_NAME"), GetObjectName(dvSource(i)("TABLE_NAME")), "")
            End If
        Next

        DataGridView1.DataSource = TableDv()

        If DataGridView1.Rows.Count > 0 Then
            TId = DataGridView1.Rows(0).Cells(0).Value

            SetColData(TableName)

            BindColData()

            DataGridView5.DataSource = FunDv()
            DataGridView12.DataSource = CrossFunDv()

            FunParamDv = GetFunParamDv(FId)
            DataGridView4.DataSource = FunParamDv

            CrossParamDv = GetCrossParamDv(CFId)
            DataGridView13.DataSource = CrossParamDv

            DataGridView1.Rows(0).DefaultCellStyle.BackColor = Color.GreenYellow
        End If

    End Sub

    Private Sub Form2_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Form0.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        Form3.Visible = False
        Form0.Visible = True
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        If ChenkFunParam() = True Then

            If IsNothing(FunParamDv) = True Then
                FunParamDv = GetFunParamDv(-1)
                DataGridView4.DataSource = FunParamDv
            End If

            Dim i As Integer
            For i = 0 To FunParamDv.Rows.Count - 1
                If FunParamDv.Rows(i)("Name") = Trim(TextBox4.Text) Then
                    MessageBox.Show("Function 參數名稱重覆")
                    Exit Sub
                End If
                FunParamDv.Rows(i)("FId") = i
            Next

            Dim newParam As DataRow
            newParam = FunParamDv.NewRow
            newParam("MId") = -1
            newParam("FId") = FunParamDv.Rows.Count
            newParam("Name") = Trim(TextBox4.Text)
            newParam("Type") = ComboBox2.SelectedItem.ToString
            newParam("Summary") = Trim(TextBox21.Text)
            FunParamDv.Rows.Add(newParam)

            CleanFunParamInfo()

        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If CheckFun() = True Then
            If GetFunInfoByName(Trim(TextBox1.Text)).DefaultView.Count > 0 Then
                MessageBox.Show("Function Name重覆")
            Else
                If CheckFunSchema() = True Then
                    AddFun(Trim(TextBox1.Text), TextBox3.Text, ComboBox1.SelectedItem.ToString, GetValue(TextBox2.Text))
                    FId = GetFId()
                    SaveFunParam()
                    DataGridView5.DataSource = FunDv()
                End If
            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If FId = -1 Then
            MessageBox.Show("請選擇一個Function")
        Else
            If CheckFun() = True Then
                If GetFunOtherByName(FId, Trim(TextBox1.Text)).DefaultView.Count > 0 Then
                    MessageBox.Show("Function Name重覆")
                Else
                    If CheckFunSchema() = True Then
                        UpDateFun(FId, Trim(TextBox1.Text), TextBox3.Text, ComboBox1.SelectedItem.ToString, GetValue(TextBox2.Text))
                        SaveFunParam()
                        DataGridView5.DataSource = FunDv()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If CheckCrossFun() = True Then
            If GetCrossFunInfoByName(TextBox19.Text).DefaultView.Count > 0 Then
                MessageBox.Show("Function Name重覆")
            Else
                If CheckCrossSchema() = True Then
                    AddCrossFun(Trim(TextBox19.Text), TextBox17.Text, GetValue(TextBox18.Text))
                    CFId = GetCFId()
                    SaveCrossParam()
                    DataGridView12.DataSource = CrossFunDv()
                End If
            End If
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        If CFId = -1 Then
            MessageBox.Show("請選擇一個Function")
        Else
            If CheckCrossFun() = True Then
                If GetCrossOtherByName(CFId, Trim(TextBox19.Text)).DefaultView.Count > 0 Then
                    MessageBox.Show("Function Name重覆")
                Else
                    If CheckCrossSchema() = True Then
                        UpDateCrossFun(CFId, Trim(TextBox19.Text), TextBox17.Text, GetValue(TextBox18.Text))
                        SaveCrossParam()
                        DataGridView12.DataSource = CrossFunDv()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click

        If ChenkCrossParam() = True Then

            If IsNothing(CrossParamDv) = True Then
                CrossParamDv = GetCrossParamDv(-1)
                DataGridView13.DataSource = CrossParamDv
            End If

            Dim i As Integer
            For i = 0 To CrossParamDv.Rows.Count - 1
                If CrossParamDv.Rows(i)("Name") = Trim(TextBox20.Text) Then
                    MessageBox.Show("Function 參數名稱重覆")
                    Exit Sub
                End If
                CrossParamDv.Rows(i)("cFId") = i
            Next

            Dim newParam As DataRow
            newParam = CrossParamDv.NewRow
            newParam("cMId") = -1
            newParam("cFId") = CrossParamDv.Rows.Count
            newParam("Name") = Trim(TextBox20.Text)
            newParam("Type") = ComboBox9.SelectedItem.ToString
            newParam("Summary") = Trim(TextBox22.Text)
            CrossParamDv.Rows.Add(newParam)

            CleanCrossParamInfo()

        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form3.Visible = True
        Form3.LoadData()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        CleanFunInfo()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        CleanCrossInfo()
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Form7.Visible = True
        Form7.LoadData()
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        RetSetColData(TableName, TId)
        BindColData()
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        'select
        Dim sb As New StringBuilder

        sb.Append("SELECT * FROM " & TableName() & " " & GetLine())

        'where
        Dim dt As DataTable = TablesSchema(TableName())
        Dim i As Integer

        If IsNothing(FunParamDv) = True Then
            FunParamDv = GetFunParamDv(-1)
            DataGridView4.DataSource = FunParamDv
        End If

        Dim isAdd As Boolean

        sb.Append("WHERE ")

        For i = 0 To ListBox2.SelectedItems.Count - 1
            If i > 0 Then
                sb.Append(" AND " & GetLine())
            Else
                sb.Append(GetLine())
            End If
            sb.Append(ListBox2.SelectedItems(i)("name"))
            sb.Append(" = @")
            sb.Append(ColumnProperty(ListBox2.SelectedItems(i)("name")) & " ")

            isAdd = True

            For j As Integer = 0 To FunParamDv.Rows.Count - 1
                If FunParamDv.Rows(j)("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name")) Then
                    isAdd = False
                End If
            Next

            If isAdd Then
                Dim newParam As DataRow
                newParam = FunParamDv.NewRow
                newParam("MId") = -1
                newParam("FId") = FunParamDv.Rows.Count
                newParam("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name"))
                newParam("Type") = dt.Columns(ListBox2.SelectedItems(i)("name")).DataType.Name
                newParam("Summary") = ColumnSummary(dt.Columns(ListBox2.SelectedItems(i)("name")).ColumnName)
                FunParamDv.Rows.Add(newParam)
            End If
        Next

        TextBox3.Text = sb.ToString

        Me.ComboBox1.SelectedItem = "List"
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        'delete
        Dim sb As New StringBuilder

        sb.Append("DELETE " & TableName() & " " & GetLine())

        'where
        Dim dt As DataTable = TablesSchema(TableName())
        Dim i As Integer

        If IsNothing(FunParamDv) = True Then
            FunParamDv = GetFunParamDv(-1)
            DataGridView4.DataSource = FunParamDv
        End If

        Dim isAdd As Boolean

        sb.Append("WHERE ")

        For i = 0 To ListBox2.SelectedItems.Count - 1
            If i > 0 Then
                sb.Append(" AND " & GetLine())
            Else
                sb.Append(GetLine())
            End If
            sb.Append(ListBox2.SelectedItems(i)("name"))
            sb.Append(" = @")
            sb.Append(ColumnProperty(ListBox2.SelectedItems(i)("name")) & " ")

            isAdd = True

            For j As Integer = 0 To FunParamDv.Rows.Count - 1
                If FunParamDv.Rows(j)("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name")) Then
                    isAdd = False
                End If
            Next

            If isAdd Then
                Dim newParam As DataRow
                newParam = FunParamDv.NewRow
                newParam("MId") = -1
                newParam("FId") = FunParamDv.Rows.Count
                newParam("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name"))
                newParam("Type") = dt.Columns(ListBox2.SelectedItems(i)("name")).DataType.Name
                newParam("Summary") = ColumnSummary(dt.Columns(ListBox2.SelectedItems(i)("name")).ColumnName)
                FunParamDv.Rows.Add(newParam)
            End If
        Next


        TextBox3.Text = sb.ToString

        Me.ComboBox1.SelectedItem = "Boolean"
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        'update
        Dim dt As DataTable = TablesSchema(TableName())

        Dim i As Integer

        Dim sb As New StringBuilder

        Dim isAdd As Boolean

        sb.Append("UPDATE " & TableName() & " SET " & GetLine())

        For i = 0 To ListBox2.SelectedItems.Count - 1

            If i > 0 Then
                sb.Append("," & GetLine())
            End If

            sb.Append(ListBox2.SelectedItems(i)("name"))
            sb.Append(" = ")
            sb.Append("@" & ColumnProperty(ListBox2.SelectedItems(i)("name")))

            isAdd = True

            For j As Integer = 0 To FunParamDv.Rows.Count - 1
                If FunParamDv.Rows(j)("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name")) Then
                    isAdd = False
                End If
            Next

            If isAdd Then
                Dim newParam As DataRow
                newParam = FunParamDv.NewRow
                newParam("MId") = -1
                newParam("FId") = FunParamDv.Rows.Count
                newParam("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name"))
                newParam("Type") = dt.Columns(ListBox2.SelectedItems(i)("name")).DataType.Name
                newParam("Summary") = ColumnSummary(dt.Columns(ListBox2.SelectedItems(i)("name")).ColumnName)
                FunParamDv.Rows.Add(newParam)
            End If

        Next

        TextBox3.Text = sb.ToString

        Me.ComboBox1.SelectedItem = "Boolean"
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        'where
        Dim dt As DataTable = TablesSchema(TableName())
        Dim i As Integer

        If IsNothing(FunParamDv) = True Then
            FunParamDv = GetFunParamDv(-1)
            DataGridView4.DataSource = FunParamDv
        End If

        Dim isAdd As Boolean

        Dim sb As New StringBuilder

        sb.Append("WHERE ")

        For i = 0 To ListBox2.SelectedItems.Count - 1
            If i > 0 Then
                sb.Append(" AND " & GetLine())
            Else
                sb.Append(GetLine())
            End If
            sb.Append(ListBox2.SelectedItems(i)("name"))
            sb.Append(" = @")
            sb.Append(ColumnProperty(ListBox2.SelectedItems(i)("name")) & " ")

            isAdd = True

            For j As Integer = 0 To FunParamDv.Rows.Count - 1
                If FunParamDv.Rows(j)("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name")) Then
                    isAdd = False
                End If
            Next

            If isAdd Then
                Dim newParam As DataRow
                newParam = FunParamDv.NewRow
                newParam("MId") = -1
                newParam("FId") = FunParamDv.Rows.Count
                newParam("Name") = ColumnProperty(ListBox2.SelectedItems(i)("name"))
                newParam("Type") = dt.Columns(ListBox2.SelectedItems(i)("name")).DataType.Name
                newParam("Summary") = ColumnSummary(dt.Columns(ListBox2.SelectedItems(i)("name")).ColumnName)
                FunParamDv.Rows.Add(newParam)
            End If

        Next

        If TextBox3.Text <> "" Then
            sb.Insert(0, GetLine(), 1)
            sb.Insert(0, TextBox3.Text, 1)
        End If

        TextBox3.Text = sb.ToString
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Form8.Visible = True
        Form8.LoadData()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.ColumnIndex = 4 And e.RowIndex > -1 Then

            TId = DataGridView1.Rows(e.RowIndex).Cells(0).Value

            CleanFunInfo()

            SetColData(TableName, TId)

            BindColData()

            DataGridView5.DataSource = FunDv()

            Dim i As Integer
            For i = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            Next
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.GreenYellow
            'DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.GreenYellow
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex > -1 Then
            UpDateTable(GetValue(DataGridView1.Rows(e.RowIndex).Cells(0).Value), Trim(GetValue(DataGridView1.Rows(e.RowIndex).Cells(2).Value)), GetValue(DataGridView1.Rows(e.RowIndex).Cells(3).Value))
        End If
    End Sub

    Private Sub DataGridView2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        If e.RowIndex > -1 Then

            Dim dv As DataView = GetColumnInfo(DataGridView2.Rows(e.RowIndex).Cells(0).Value).DefaultView

            Dim row As DataGridViewRow = DataGridView2.Rows(e.RowIndex)

            UpDateColumn(row.Cells(0).Value,
                         Trim(GetValue(row.Cells(2).Value)),
                         GetValue(row.Cells(3).Value),
                         GetValue(row.Cells("InputType").Value),
                         row.Cells("Min").Value,
                         row.Cells("Max").Value,
                         row.Cells("Required").Value,
                         GetValue(row.Cells("Remark").Value))
            row.Cells(3).ToolTipText = GetValue(row.Cells(3).Value)
            row.Cells("Remark").ToolTipText = GetValue(row.Cells("Remark").Value)
        End If
    End Sub

    Private Sub DataGridView4_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick

        If e.RowIndex > -1 Then
            If e.ColumnIndex = (DataGridView4.ColumnCount - 1) Then
                FunParamDv.Rows.RemoveAt(DataGridView4.Rows(e.RowIndex).Cells(0).Value)
                Dim i As Integer
                For i = 0 To FunParamDv.Rows.Count - 1
                    FunParamDv.Rows(i)("FId") = i
                Next
            End If
        End If

    End Sub

    Private Sub DataGridView5_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        If e.RowIndex > -1 Then
            Select Case e.ColumnIndex
                Case 1
                    FId = DataGridView5.Rows(e.RowIndex).Cells(0).Value
                    ReadFunInfo()
                Case 2
                    If MessageBox.Show("確定刪除", "", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                        'DelFun(DataGridView5.Rows(e.RowIndex).Cells(0).Value)
                        mdDel.DeleteFun(DataGridView5.Rows(e.RowIndex).Cells(0).Value)
                        CleanFunInfo()
                        DataGridView5.DataSource = FunDv()
                    End If
            End Select
        End If
    End Sub

    Private Sub DataGridView12_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView12.CellClick
        If e.RowIndex > -1 Then
            Select Case e.ColumnIndex
                Case 1
                    CFId = DataGridView12.Rows(e.RowIndex).Cells(0).Value
                    ReadCrossInfo()
                Case 2
                    If MessageBox.Show("確定刪除", "", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                        'DelCrossFun(DataGridView12.Rows(e.RowIndex).Cells(0).Value)
                        mdDel.DeleteCross(DataGridView12.Rows(e.RowIndex).Cells(0).Value)
                        CleanCrossInfo()
                        DataGridView12.DataSource = CrossFunDv()
                    End If
            End Select
        End If
    End Sub

    Private Sub DataGridView13_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView13.CellClick

        If e.RowIndex > -1 Then
            If e.ColumnIndex = (DataGridView13.ColumnCount - 1) Then
                CrossParamDv.Rows.RemoveAt(DataGridView13.Rows(e.RowIndex).Cells(0).Value)
                Dim i As Integer
                For i = 0 To CrossParamDv.Rows.Count - 1
                    CrossParamDv.Rows(i)("cFId") = i
                Next
            End If
        End If

    End Sub

    Dim FunParamDv As DataTable
    Private Sub ReadFunInfo()

        Dim dt As DataTable = GetFunInfo(FId)

        TextBox1.Text = dt.DefaultView(0)("Name")
        TextBox2.Text = dt.DefaultView(0)("Summary")
        TextBox3.Text = dt.DefaultView(0)("SqlText")
        ComboBox1.SelectedItem = dt.DefaultView(0)("ReturnType")

        FunParamDv = GetFunParamDv(FId)
        Dim i As Integer
        For i = 0 To FunParamDv.Rows.Count - 1
            FunParamDv.Rows(i)("FId") = i
        Next
        DataGridView4.DataSource = FunParamDv

    End Sub

    Dim CrossParamDv As DataTable
    Private Sub ReadCrossInfo()

        Dim dt As DataTable = GetCrossFunInfo(CFId)

        TextBox19.Text = dt.DefaultView(0)("Name")
        TextBox18.Text = dt.DefaultView(0)("Summary")
        TextBox17.Text = dt.DefaultView(0)("SqlText")

        CrossParamDv = GetCrossParamDv(CFId)
        Dim i As Integer
        For i = 0 To CrossParamDv.Rows.Count - 1
            CrossParamDv.Rows(i)("cFId") = i
        Next

        DataGridView13.DataSource = CrossParamDv

    End Sub

    Private Sub CleanFunInfo()

        FId = -1

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.SelectedIndex = -1

        If DataGridView4.Rows.Count > 0 Then
            FunParamDv = GetFunParamDv(-1)
            DataGridView4.DataSource = FunParamDv
        End If

        CleanFunParamInfo()

    End Sub

    Private Sub CleanFunParamInfo()

        TextBox4.Text = ""
        TextBox21.Text = ""
        ComboBox2.SelectedIndex = -1

    End Sub

    Private Sub SaveFunParam()

        DelFunParamByFId(FId)

        If FunParamDv.DefaultView.Count > 0 Then
            Dim i As Integer
            For i = 0 To FunParamDv.DefaultView.Count - 1
                AddFunParam(FId, FunParamDv.DefaultView(i)("Name"), FunParamDv.DefaultView(i)("Type"), FunParamDv.DefaultView(i)("Summary"))
            Next
        End If

    End Sub

    Private Sub CleanCrossInfo()

        CFId = -1

        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""

        If DataGridView13.Rows.Count > 0 Then
            CrossParamDv = GetCrossParamDv(-1)
            DataGridView13.DataSource = CrossParamDv
        End If

        CleanCrossParamInfo()

    End Sub

    Private Sub CleanCrossParamInfo()

        TextBox20.Text = ""
        TextBox22.Text = ""
        ComboBox9.SelectedIndex = -1

    End Sub

    Private Sub SaveCrossParam()

        DelCrossParamByCFId(CFId)

        If CrossParamDv.DefaultView.Count > 0 Then
            Dim i As Integer
            For i = 0 To CrossParamDv.DefaultView.Count - 1
                AddCrossParam(CFId, CrossParamDv.DefaultView(i)("Name"), CrossParamDv.DefaultView(i)("Type"), CrossParamDv.DefaultView(i)("Summary"))
            Next
        End If

    End Sub

    Private Function CheckFun() As Boolean

        Dim bool As Boolean = True
        Dim str As String = ""

        If Trim(TextBox1.Text) = "" Then
            bool = False
            str = "請輸入Function名稱"
        Else
            Select Case Trim(TextBox1.Text).ToLower
                Case "GetInfo".ToLower
                    bool = False
                    str = "GetInfo為資料表基本函式之ㄧ 請修改Function名稱"
                Case "AddNew".ToLower
                    bool = False
                    str = "AddNew為資料表基本函式之ㄧ 請修改Function名稱"
                Case "Update".ToLower
                    bool = False
                    str = "Update為資料表基本函式之ㄧ 請修改Function名稱"
                Case "Del".ToLower
                    bool = False
                    str = "Del為資料表基本函式之ㄧ 請修改Function名稱"
                Case "GetAll".ToLower
                    bool = False
                    str = "GetAll為資料表基本函式之ㄧ 請修改Function名稱"
            End Select
        End If

        If ComboBox1.SelectedIndex < 0 Then
            bool = False
            str &= Chr(13) & "請選擇回傳類型"
        End If

        If Trim(TextBox3.Text) = "" Then
            bool = False
            str &= Chr(13) & "請輸入SQL語法"
        End If

        If bool = False Then
            MessageBox.Show(str)
        End If

        Return bool

    End Function

    Private Function ChenkFunParam() As Boolean

        Dim bool As Boolean = True
        Dim str As String = ""

        If Trim(TextBox4.Text) = "" Then
            bool = False
            str = "請輸入參數名稱"
        End If

        If ComboBox2.SelectedIndex < 0 Then
            bool = False
            str &= Chr(13) & "請選擇參數型別"
        End If

        If bool = False Then
            MessageBox.Show(str)
        End If

        Return bool

    End Function

    Public Function CheckFunSchema() As Boolean

        Dim strSQL As String = ""
        Dim strReturn = ComboBox1.SelectedItem.ToString
        Dim strErr As String = ""
        Dim bool As Boolean = True
        Dim i As Integer

        Dim arySQL As Array = TextBox3.Text.Replace(Chr(10), GetLine()).Replace(Chr(13), GetLine()).Split(GetLine())

        For i = 0 To arySQL.Length - 1
            If arySQL(i).ToString.Trim.StartsWith("--") = False Then
                strSQL &= arySQL(i)
                strSQL &= " "
            End If
        Next

        If strSQL.Contains("insert ") = True Or _
        strSQL.Contains("INSERT ") = True Or _
        strSQL.Contains("Insert ") = True Then
            bool = False
            strErr = "SQL不可輸入Insert語法" & Chr(13)
        ElseIf strSQL.Trim.StartsWith("update ") = True Or _
        strSQL.Trim.StartsWith("UPDATE ") = True Or _
        strSQL.Trim.StartsWith("Update ") = True Or _
        strSQL.Trim.StartsWith("delete ") = True Or _
        strSQL.Trim.StartsWith("DELETE ") = True Or _
        strSQL.Trim.StartsWith("Delete ") = True Then
            If strReturn <> "Boolean" Then
                bool = False
                strErr = "傳回值請選擇Boolean" & Chr(13)
            Else
                Dim strTB As String = TableName()
                If strSQL.Contains(" " & strTB & " ") = False And strSQL.Contains("[" & strTB & "]") = False Then
                    bool = False
                    strErr = "只能對資料表" & TableName() & "做修改刪除" & Chr(13)
                End If
            End If
        Else
            If strReturn = "Boolean" Then
                bool = False
                strErr = "傳回值不適合選擇Boolean" & Chr(13)
            End If
        End If

        If bool = True And FunParamDv.DefaultView.Count > 0 Then
            For i = 0 To FunParamDv.DefaultView.Count - 1
                If strSQL.Contains("@" & FunParamDv.DefaultView(i)("Name").ToString) = True Then
                    strSQL = ReplaceSQL(strSQL, FunParamDv.DefaultView(i)("Name").ToString, FunParamDv.DefaultView(i)("Type"))
                Else
                    bool = False
                    strErr = "多出SQL不需要的參數" & Chr(13)
                End If
            Next
        End If

        If bool = True And strSQL.Contains("@") = True Then
            bool = False
            strErr = "缺少SQL需要的參數" & Chr(13)
        End If
        Dim dtTable As DataTable
        If bool = True Then
            Try
                dtTable = GetSchema(strSQL)
            Catch ex As Data.Common.DbException 'SqlClient.SqlException
                bool = False
                strErr = ex.Message
                CloseOutDbCom()
            End Try

            If bool = True And strReturn <> "Boolean" Then
                Dim dtTable2 As DataTable = TablesSchema(TableName())

                If dtTable.Columns.Count = dtTable2.Columns.Count Then
                    For i = 0 To dtTable.Columns.Count - 1
                        If dtTable2.Columns.Contains(dtTable.Columns(i).ColumnName) = False Then
                            bool = False
                            strErr = "不符合資料表" & TableName() & "格式" & Chr(13)
                        End If
                    Next
                Else
                    bool = False
                    strErr = "不符合資料表" & TableName() & "格式" & Chr(13)
                End If
            End If
        End If

        If bool = False Then
            MessageBox.Show(strErr)
        End If

        Return bool

    End Function

    Private Function CheckCrossFun() As Boolean

        Dim bool As Boolean = True
        Dim str As String = ""

        If Trim(TextBox19.Text) = "" Then
            bool = False
            str = Chr(13) & "請輸入Function名稱"
        End If

        If Trim(TextBox17.Text) = "" Then
            bool = False
            str &= Chr(13) & "請輸入SQL語法"
        End If

        If bool = False Then
            MessageBox.Show(str)
        End If

        Return bool

    End Function

    Private Function ChenkCrossParam() As Boolean

        Dim bool As Boolean = True
        Dim str As String = ""

        If Trim(TextBox20.Text) = "" Then
            bool = False
            str = "請輸入參數名稱"
        End If

        If ComboBox9.SelectedIndex < 0 Then
            bool = False
            str &= Chr(13) & "請選擇參數型別"
        End If

        If bool = False Then
            MessageBox.Show(str)
        End If

        Return bool

    End Function

    Private Function CheckCrossSchema() As Boolean

        Dim strSQL As String = ""
        Dim strErr As String = ""
        Dim bool As Boolean = True
        Dim i As Integer

        Dim arySQL As Array = TextBox17.Text.Replace(Chr(10), GetLine()).Replace(Chr(13), GetLine()).Split(GetLine())

        For i = 0 To arySQL.Length - 1
            If arySQL(i).ToString.Trim.StartsWith("--") = False Then
                strSQL &= arySQL(i)
                strSQL &= " "
            End If
        Next

        If strSQL.Contains("insert ") = True Or _
        strSQL.Contains("INSERT ") = True Or _
        strSQL.Contains("Insert ") = True Or _
        strSQL.Trim.StartsWith("update ") = True Or _
        strSQL.Trim.StartsWith("UPDATE ") = True Or _
        strSQL.Trim.StartsWith("Update ") = True Or _
        strSQL.Trim.StartsWith("delete ") = True Or _
        strSQL.Trim.StartsWith("DELETE ") = True Or _
        strSQL.Trim.StartsWith("Delete ") = True Then
            bool = False
            strErr = "SQL不可輸入Insert,update,delete語法" & Chr(13)
        End If

        If bool = True And CrossParamDv.DefaultView.Count > 0 Then
            For i = 0 To CrossParamDv.DefaultView.Count - 1
                If strSQL.Contains("@" & CrossParamDv.DefaultView(i)("Name").ToString) = True Then
                    strSQL = ReplaceSQL(strSQL, CrossParamDv.DefaultView(i)("Name").ToString, CrossParamDv.DefaultView(i)("Type"))
                Else
                    bool = False
                    strErr = "存在SQL不需要的參數" & Chr(13)
                End If
            Next
        End If

        If bool = True And strSQL.Contains("@") = True Then
            bool = False
            strErr = "缺少SQL需要的參數" & Chr(13)
        End If

        If bool = True Then
            Try
                Dim dtTable As DataTable = GetSchema(strSQL)
            Catch ex As Data.Common.DbException
                bool = False
                strErr = ex.Message
                CloseOutDbCom()
            End Try
        End If

        If bool = False Then
            MessageBox.Show(strErr)
        End If

        Return bool

    End Function

    Private Function TableDv() As DataView

        Dim Dv As DataView
        Dv = GetTableDv().DefaultView
        'Dim i As Integer
        'For i = 0 To Dv.Count - 1
        '    If Trim(Dv(i)("Summary")) <> "" Then
        '        Dv(i)("Name") = Dv(i)("Name") & "(" & Dv(i)("Summary") & ")"
        '    End If
        'Next

        Return Dv

    End Function

    Private Function FunDv() As DataView

        Dim Dv As DataView
        Dv = GetFunDv().DefaultView
        Dim i As Integer
        For i = 0 To Dv.Count - 1
            If Trim(Dv(i)("Summary")) <> "" Then
                Dv(i)("Name") = Dv(i)("Name") & "(" & Dv(i)("Summary") & ")"
            End If
        Next

        Return Dv

    End Function

    Private Function CrossFunDv() As DataView

        Dim Dv As DataView
        Dv = GetCrossFunDv().DefaultView
        Dim i As Integer
        For i = 0 To Dv.Count - 1
            If Trim(Dv(i)("Summary")) <> "" Then
                Dv(i)("Name") = Dv(i)("Name") & "(" & Dv(i)("Summary") & ")"
            End If
        Next

        Return Dv

    End Function

    Public Sub BindColData()

        Dim dv As DataView = GetColumnDv().DefaultView

        DataGridView2.DataSource = dv

        Dim dt As DataTable = TablesSchema(TableName)
        Dim i As Integer
        For i = 0 To DataGridView2.Rows.Count - 1

            Dim colName As String = GetValue(DataGridView2.Rows(i).Cells("GV2Name").Value)
            Dim j As Integer
            For j = 0 To dt.PrimaryKey.GetUpperBound(0)
                If colName = dt.PrimaryKey(j).ColumnName Then
                    DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.GreenYellow
                End If
            Next
            If dt.Columns(colName).AutoIncrement = True Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.GreenYellow
            End If
        Next

        ListBox2.DataSource = dv
    End Sub

    Private Sub RetSetColData(ByVal TableName As String, Optional ByVal intTId As Integer = -1)
        Dim i As Integer
        Dim dv As DataView = GetColumnDv(intTId).DefaultView

        For i = 0 To dv.Count - 1
            DeleteColumn(dv(i)("ColId"))
        Next

        Dim dvSource As DataView = ColumnsSource(TableName).DefaultView
        Dim colCommons As String
        Dim Input As String
        Dim Min As Object
        Dim Max As Object
        Dim Required As Boolean
        Dim Remark As String

        Dim colName As String
        Select Case DbType
            Case "SQLITE"
                colName = "name"
            Case Else
                colName = "COLUMN_NAME"
        End Select

        For i = 0 To dvSource.Count - 1
            dv.RowFilter = "Name ='" & dvSource(i)(colName) & "'"
            If dv.Count > 0 Then
                colCommons = dv(0)("Summary")
                Input = dv(0)("InputType")
                Min = dv(0)("MinVal")
                Max = dv(0)("MaxVal")
                Required = dv(0)("Required")
                Remark = GetValue(dv(0)("Remark"))
            Else
                colCommons = ""
                Input = "Text"
                Min = DBNull.Value
                Max = DBNull.Value
                Required = False
                Remark = ""
            End If

            AddColumn(dvSource(i)(colName),
                      GetObjectName(dvSource(i)(colName)),
                      colCommons,
                      Input,
                      Min,
                      Max,
                      Required,
                      Remark)
        Next
    End Sub

End Class