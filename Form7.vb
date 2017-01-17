Imports System.Text
Public Class Form7
    Private ReadOnly Property GridView() As String
        Get
            If TextBox4.Visible = True Then
                If TextBox4.Text.Trim() <> "" Then
                    Return TextBox4.Text.Trim()
                Else
                    Return "GridView1"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property btnAdd() As String
        Get
            If TextBox5.Visible = True Then
                If TextBox5.Text.Trim() <> "" Then
                    Return TextBox5.Text.Trim()
                Else
                    Return "btnAdd"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property btnSave() As String
        Get
            If TextBox6.Visible = True Then
                If TextBox6.Text.Trim() <> "" Then
                    Return TextBox6.Text.Trim()
                Else
                    Return "btnSave"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property btnBack() As String
        Get
            If TextBox7.Visible = True Then
                If TextBox7.Text.Trim() <> "" Then
                    Return TextBox7.Text.Trim()
                Else
                    Return "btnBack"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property btnQuery() As String
        Get
            If TextBox12.Visible = True Then
                If TextBox12.Text.Trim() <> "" Then
                    Return TextBox12.Text.Trim()
                Else
                    Return "btnQuery"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property ObjectDataSource() As String
        Get
            If TextBox8.Visible = True Then
                If TextBox8.Text.Trim() <> "" Then
                    Return TextBox8.Text.Trim()
                Else
                    Return "ObjectDataSource1"
                End If
            Else
                Return ""
            End If
        End Get
    End Property

    Private ReadOnly Property ReadData() As String
        Get
            If TextBox9.Visible = True Then
                If TextBox9.Text.Trim() <> "" Then
                    Return TextBox9.Text.Trim()
                Else
                    Return "ReadData"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property CleanData() As String
        Get
            If TextBox10.Visible = True Then
                If TextBox10.Text.Trim() <> "" Then
                    Return TextBox10.Text.Trim()
                Else
                    Return "CleanData"
                End If
            Else
                Return ""
            End If
        End Get
    End Property
    Private ReadOnly Property BindGV() As String
        Get
            If TextBox11.Visible = True Then
                If TextBox11.Text.Trim() <> "" Then
                    Return TextBox11.Text.Trim()
                Else
                    Return "BindGV"
                End If
            Else
                Return ""
            End If
        End Get
    End Property

    Private ReadOnly Property haveDEL() As Boolean
        Get
            If CheckBox2.Visible = False Then
                Return False
            Else
                Return CheckBox2.Checked
            End If
        End Get
    End Property
    Private ReadOnly Property haveReadOnly() As Boolean
        Get
            If CheckBox3.Visible = False Then
                Return False
            Else
                Return CheckBox3.Checked
            End If
        End Get
    End Property
    Private ReadOnly Property haveListOnly() As Boolean
        Get
            If CheckBox6.Visible = False Then
                Return False
            Else
                Return CheckBox6.Checked
            End If
        End Get
    End Property
    Private ReadOnly Property haveAddOnly() As Boolean
        Get
            If CheckBox1.Visible = False Then
                Return False
            Else
                Return CheckBox1.Checked
            End If
        End Get
    End Property
    Private ReadOnly Property haveMultiView() As Boolean
        Get
            If CheckBox7.Visible = False Then
                Return False
            Else
                Return CheckBox7.Checked
            End If
        End Get
    End Property
    Private ReadOnly Property haveObjectDataSource() As Boolean
        Get
            If CheckBox4.Visible = False Then
                Return False
            Else
                Return CheckBox4.Checked
            End If
        End Get
    End Property
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

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim intTId As Integer = TId
        TId = ComboBox1.SelectedItem("TId")
        Table_Initial()

        ComboBox2.DataSource = FunDv()
        BindGV2()

        mdFile.classDT.Clear()
        TId = intTId
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        BindGV1()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            CodeType = RadioButton1.Tag
        Else
            CodeType = RadioButton2.Tag
        End If
        Dim intTId As Integer = TId
        TId = ComboBox1.SelectedItem("TId")
        Table_Initial()
        GetUI()
        GetCodeFile()
        mdFile.classDT.Clear()
        TId = intTId
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        CheckBox2.Visible = Not CheckBox1.Checked
        CheckBox3.Visible = Not CheckBox1.Checked
        CheckBox6.Visible = Not CheckBox1.Checked
        CheckBox7.Visible = Not CheckBox1.Checked
        Label8.Visible = Not CheckBox1.Checked

        Label12.Visible = Not CheckBox1.Checked
        Label15.Visible = Not CheckBox1.Checked
        Label17.Visible = Not CheckBox1.Checked
        TextBox4.Visible = Not CheckBox1.Checked

        'TextBox5.Visible = Not CheckBox1.Checked
        'Label9.Visible = Not CheckBox1.Checked
        SetBtnAdd(Not CheckBox1.Checked)

        TextBox8.Visible = Not CheckBox1.Checked
        TextBox9.Visible = Not CheckBox1.Checked
        TextBox11.Visible = Not CheckBox1.Checked
        Panel1.Visible = Not CheckBox1.Checked

        DataGridView2.Columns("chkList").Visible = Not haveAddOnly
        Label13.Visible = Not haveAddOnly
        TextBox12.Visible = Not haveAddOnly
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged, CheckBox3.VisibleChanged

        Label10.Visible = Not CheckBox3.Checked
        Label16.Visible = Not CheckBox3.Checked

        SetBtnSave()
        SetBtnAdd()
        SetBtnBack()

        DataGridView2.Columns("EditType").Visible = Not haveReadOnly
        SetGV2()
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Me.TextBox8.Visible = CheckBox4.Checked
        Me.Label12.Visible = CheckBox4.Checked
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        CheckBox1.Visible = Not CheckBox6.Checked
        CheckBox3.Visible = Not CheckBox6.Checked
        CheckBox7.Visible = Not CheckBox6.Checked

        Label15.Visible = Not CheckBox6.Checked
        Label16.Visible = Not CheckBox6.Checked

        SetBtnSave()
        SetBtnAdd(Not CheckBox6.Checked)
        SetBtnBack(Not CheckBox6.Checked)

        TextBox9.Visible = Not CheckBox6.Checked
        TextBox10.Visible = Not CheckBox6.Checked

        DataGridView2.Columns("chkDOC").Visible = Not haveListOnly
        If haveReadOnly = False Then
            DataGridView2.Columns("EditType").Visible = Not haveListOnly
        End If
        DataGridView2.Columns("EditName").Visible = Not haveListOnly
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If haveReadOnly = False Then
            Label9.Visible = CheckBox7.Checked
            TextBox5.Visible = CheckBox7.Checked
        End If
        SetBtnBack()
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.ColumnIndex = 4 And e.RowIndex > -1 Then
            If IsNothing(DataGridView1.Rows(e.RowIndex).Cells("defVal").Value) Then
                DataGridView1.Rows(e.RowIndex).Cells("defVal").Value = ""
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        If e.ColumnIndex = 5 And e.RowIndex > -1 Then            
            DataGridView2.Rows(e.RowIndex).Cells("EditName").Value = GetControlName(DataGridView2.Rows(e.RowIndex).Cells("EditType").Value, ColumnProperty(DataGridView2.Rows(e.RowIndex).Cells("FieldName").Value), False)
        End If
    End Sub

    Private Sub DataGridView2_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView2.Sorted
        SetGV2()
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

    Private Sub BindGV2()
        SetColData(ComboBox1.SelectedItem("Name"), TId)
        DataGridView2.DataSource = GetColumnDv()
        SetGV2()
    End Sub

    Private Sub SetGV2()
        Dim i, j As Integer
        For i = 0 To DataGridView2.Rows.Count - 1
            DataGridView2.Rows(i).Cells("chkDOC").Value = True

            If haveReadOnly = False Then
                DataGridView2.Rows(i).Cells("EditType").Value = "TextBox"
                DataGridView2.Rows(i).Cells("EditName").Value = "txt" & ColumnProperty(DataGridView2.Rows(i).Cells("FieldName").Value)
            Else
                DataGridView2.Rows(i).Cells("EditName").Value = "literal" & ColumnProperty(DataGridView2.Rows(i).Cells("FieldName").Value)
            End If

            Dim mSetSummary As String = DataGridView2.Rows(i).Cells("Summary").Value
            If mSetSummary.Trim <> "" Then
                DataGridView2.Rows(i).Cells("FieldSummary").Value = DataGridView2.Rows(i).Cells("FieldName").Value & "(" & mSetSummary.Trim & ")"
            Else
                DataGridView2.Rows(i).Cells("FieldSummary").Value = DataGridView2.Rows(i).Cells("FieldName").Value
            End If

            For j = 0 To mdFile.classDT.PrimaryKey.GetUpperBound(0)
                If DataGridView2.Rows(i).Cells("FieldName").Value = mdFile.classDT.PrimaryKey(j).ColumnName Then
                    DataGridView2.Rows(i).Cells("FieldSummary").Style.BackColor = Color.SkyBlue
                    DataGridView2.Rows(i).Cells("chkDOC").ReadOnly = Not haveReadOnly
                    DataGridView2.Rows(i).Cells("EditType").ReadOnly = Not haveReadOnly
                End If
            Next
            If mdFile.classDT.Columns(DataGridView2.Rows(i).Cells("FieldName").Value).AutoIncrement = True Then
                DataGridView2.Rows(i).Cells("FieldSummary").Style.BackColor = Color.GreenYellow
                DataGridView2.Rows(i).Cells("chkDOC").Value = haveReadOnly
                DataGridView2.Rows(i).Cells("chkDOC").ReadOnly = Not haveReadOnly
            End If
        Next
    End Sub

    Private Sub SetBtnBack(Optional ByVal chk As Boolean = True)
        If chk = True Then
            If haveMultiView = False And haveReadOnly = True Then
                TextBox7.Visible = False
                Label11.Visible = False
            Else
                Me.TextBox7.Visible = True
                Me.Label11.Visible = True
            End If
        Else
            Me.TextBox7.Visible = False
            Me.Label11.Visible = False
        End If
    End Sub

    Private Sub SetBtnAdd(Optional ByVal chk As Boolean = True)
        If chk = True Then
            If haveMultiView = True And haveReadOnly = False Then
                TextBox5.Visible = True
                Label9.Visible = True
            Else
                Me.TextBox5.Visible = False
                Me.Label9.Visible = False
            End If
        Else
            Me.TextBox5.Visible = False
            Me.Label9.Visible = False
        End If
    End Sub

    Private Sub SetBtnSave()
        If haveAddOnly = True Then
            TextBox6.Visible = True
            Label10.Visible = True
        Else
            If haveListOnly = True Then
                TextBox6.Visible = False
                Label10.Visible = False
            Else
                If haveReadOnly = True Then
                    TextBox6.Visible = False
                    Label10.Visible = False
                Else
                    TextBox6.Visible = True
                    Label10.Visible = True
                End If
            End If
        End If
    End Sub

    Private Sub GetUI()
        Dim strSET As New StringBuilder
        strSET.Append("    <div>" & GetLine())
        If haveMultiView = True Then
            strSET.Append("        <asp:MultiView ID=""MultiView1"" runat=""server"">" & GetLine())
            strSET.Append("            <asp:View ID=""View1"" runat=""server"">" & GetLine())
        End If

        Dim haveQuery As Boolean = False
        If haveAddOnly = False Then
            For Each Item As DataGridViewRow In DataGridView1.Rows
                Select Case Item.Cells("FromType").Value
                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                        haveQuery = True
                End Select
            Next
        End If
        If haveQuery = True Then
            strSET.Append("                <table>" & GetLine())
            strSET.Append("                    <tr>" & GetLine())
            For Each Item As DataGridViewRow In DataGridView1.Rows
                Select Case Item.Cells("FromType").Value
                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                        strSET.Append(String.Format("                        <td>{0}<asp:{1} ID=""{2}"" runat=""server""></asp:{1}></td>", Item.Cells("ParamSummary").Value, Item.Cells("FromType").Value, GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)) & GetLine())
                End Select
            Next
            strSET.Append(String.Format("                        <td><asp:Button ID=""{0}"" runat=""server"" Text=""琩高"" ", btnQuery))
            If CodeType = "cs" Then
                strSET.Append(String.Format("OnClick=""{0}_Click"" ", btnQuery))
            End If
            strSET.Append("/></td>" & GetLine())
            strSET.Append("                    </tr>" & GetLine())
            strSET.Append("                </table>" & GetLine())
        End If

        If btnAdd <> "" Then
            strSET.Append(String.Format("                <asp:Button ID=""{0}"" runat=""server"" Text=""穝糤"" ", btnAdd))
            If CodeType = "cs" Then
                strSET.Append(String.Format("OnClick=""{0}_Click"" ", btnAdd))
            End If
            strSET.Append("/>" & GetLine())
        End If

        If GridView <> "" Then
            strSET.Append(String.Format("                <asp:GridView ID=""{0}"" runat=""server"" AutoGenerateColumns=""False"" DataKeyNames=""{1}"" ", GridView, mdFile.strKEY))
            If PageSize > 0 Then
                strSET.Append(String.Format("AllowPaging=""true"" PageSize=""{0}"" ", PageSize.ToString()))
            End If
            If ObjectDataSource <> "" Then
                strSET.Append(String.Format("DataSourceID=""{0}"" ", ObjectDataSource))
            End If
            If CodeType = "cs" Then
                If PageSize > 0 And ObjectDataSource = "" Then
                    strSET.Append(String.Format("OnPageIndexChanging=""{0}_PageIndexChanging"" ", GridView))
                End If
                If (haveListOnly = False Or Me.haveDEL = True) Then
                    strSET.Append(String.Format("OnRowCommand=""{0}_RowCommand""", GridView))
                End If
            End If
            strSET.Append(">" & GetLine())

            strSET.Append("                    <Columns>" & GetLine())

            For Each Item As DataGridViewRow In DataGridView2.Rows
                If Item.Cells("chkList").Value = True Then
                    strSET.Append(String.Format("                        <asp:BoundField DataField=""{0}"" HeaderText=""{1}"" ", ColumnProperty(Item.Cells("FieldName").Value), ColumnSummary(Item.Cells("FieldName").Value)))
                    If mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name = "DateTime" Then
                        strSET.Append("DataFormatString=""{0:d}"" ")
                    End If
                    strSET.Append("/>" & GetLine())
                End If
            Next

            If haveListOnly = False Then
                strSET.Append("                        <asp:ButtonField ButtonType=""Button"" CommandName=""read"" Text=""弄"" />" & GetLine())
            End If

            If haveDEL = True Then
                strSET.Append("                        <asp:TemplateField>" & GetLine())
                strSET.Append("                            <ItemTemplate>" & GetLine())
                strSET.Append(String.Format("                                <asp:Button ID=""btnDel"" runat=""server"" Text=""埃"" CommandName=""del"" CommandArgument='<%#{0}.Rows.Count%>' OnClientClick=""return confirm('絋﹚埃?')"" />", GridView) & GetLine())
                strSET.Append("                            </ItemTemplate>" & GetLine())
                strSET.Append("                        </asp:TemplateField>" & GetLine())
            End If
            strSET.Append("                    </Columns>" & GetLine())
            strSET.Append("                </asp:GridView>" & GetLine())
        End If

        If ObjectDataSource <> "" Then
            strSET.Append(String.Format("                <asp:ObjectDataSource ID=""{0}"" runat=""server"" SelectMethod=""{1}"" TypeName=""Business.{2}Biz"">", ObjectDataSource, ComboBox2.SelectedItem("Name"), mdFile.strClass) & GetLine())
            If Me.DataGridView1.Rows.Count > 0 Then
                strSET.Append("                    <SelectParameters>" & GetLine())
                For Each Item As DataGridViewRow In DataGridView1.Rows
                    Dim DefaultValue As String
                    If Item.Cells("defVal").Value.ToString.Trim = """""" Then
                        DefaultValue = " "
                    Else
                        DefaultValue = Item.Cells("defVal").Value.ToString.Trim
                    End If

                    Select Case Item.Cells("FromType").Value
                        Case "Request"
                            strSET.Append(String.Format("                        <asp:QueryStringParameter DefaultValue=""{0}"" Name=""{1}"" QueryStringField=""{2}"" Type=""{3}"" />", DefaultValue, Item.Cells("ParamName").Value, Item.Cells("FromName").Value, Item.Cells("ParamType").Value) & GetLine())
                        Case "Session"
                            strSET.Append(String.Format("                        <asp:SessionParameter DefaultValue=""{0}"" Name=""{1}"" SessionField=""{2}"" Type=""{3}"" />", DefaultValue, Item.Cells("ParamName").Value, Item.Cells("FromName").Value, Item.Cells("ParamType").Value) & GetLine())
                        Case Else
                            strSET.Append(String.Format("                        <asp:Parameter DefaultValue=""{0}"" Name=""{1}"" Type=""{2}"" />", DefaultValue, Item.Cells("ParamName").Value, Item.Cells("ParamType").Value) & GetLine())
                    End Select
                Next
                strSET.Append("                    </SelectParameters>" & GetLine())
            End If
            strSET.Append("                </asp:ObjectDataSource>" & GetLine())
        End If

        If haveMultiView = True Then
            strSET.Append("            </asp:View>" & GetLine())
            strSET.Append("            <asp:View ID=""View2"" runat=""server"">" & GetLine())
        End If

        If haveListOnly = False Then
            strSET.Append("            <table>" & GetLine())
            For Each Item As DataGridViewRow In DataGridView2.Rows
                If Item.Cells("chkDOC").Value = True Then
                    Dim strEditType As String
                    If btnSave <> "" Then
                        strEditType = Item.Cells("EditType").Value
                    Else
                        strEditType = "Literal"
                    End If
                    strSET.Append(String.Format("                <tr><th align=""right"">{0}</th><td><asp:{1} ID=""{2}"" runat=""server""></asp:{1}></td></tr>", ColumnSummary(Item.Cells("FieldName").Value), strEditType, Item.Cells("EditName").Value) & GetLine())
                End If
            Next
            strSET.Append("            </table>" & GetLine())
        End If

        If btnSave <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("                <asp:Button ID=""{0}"" runat=""server"" Text=""穝糤"" OnClick=""{0}_Click"" />", btnSave) & GetLine())
                Case "vb"
                    strSET.Append(String.Format("                <asp:Button ID=""{0}"" runat=""server"" Text=""穝糤"" />", btnSave) & GetLine())
            End Select
        End If

        If btnBack <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("                <asp:Button ID=""{0}"" runat=""server"" Text="""" OnClick=""{0}_Click"" />", btnBack) & GetLine())
                Case "vb"
                    strSET.Append(String.Format("                <asp:Button ID=""{0}"" runat=""server"" Text="""" />", btnBack) & GetLine())
            End Select
        End If

        If haveMultiView = True Then
            strSET.Append("            </asp:View>" & GetLine())
            strSET.Append("        </asp:MultiView>" & GetLine())
        End If

        strSET.Append("    </div>")

        TextBox2.Text = strSET.ToString
    End Sub

    Private Sub GetCodeFile()
        Dim strSET As New StringBuilder
        Dim i As Integer

        'ゅンkey
        If ReadData <> "" Then
            If mdFile.ArrayKEY.Length > 0 Then
                strSET.Append(GetRegion("ゅン把计"))
                For i = 0 To mdFile.ArrayKEY.Length - 1
                    strSET.Append(GetViewState(ColumnProperty(mdFile.ArrayKEY(i)), mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name))
                Next
                strSET.Append(EndRegion())
            End If
        End If

        Dim haveBindGV As Boolean = False
        If BindGV <> "" Then
            If ObjectDataSource <> "" Then
                For Each Item As DataGridViewRow In DataGridView1.Rows
                    Select Case Item.Cells("FromType").Value
                        Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox", "other"
                            haveBindGV = True
                    End Select
                Next
            Else
                haveBindGV = True
            End If
        End If

        '把计
        If haveBindGV = True Then
            Dim defVal As String
            Dim ListParams As New StringBuilder
            For Each Item As DataGridViewRow In DataGridView1.Rows
                If Item.Cells("defVal").Value.ToString.Trim = """""" Then
                    defVal = ""
                Else
                    defVal = Item.Cells("defVal").Value.ToString.Trim
                End If
                Select Case Item.Cells("FromType").Value
                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                        If ObjectDataSource = "" Then
                            ListParams.Append(GetViewState(Item.Cells("FromName").Value, Item.Cells("ParamType").Value, defVal))
                        Else
                            If defVal = "" Then
                                defVal = """ """ 'obds﹃
                            End If
                            ListParams.Append(GetReadOnlyProperty(Item.Cells("FromName").Value, Item.Cells("ParamType").Value, Item.Cells("FromType").Value, defVal))
                        End If
                    Case "Session"
                        If ObjectDataSource = "" Then
                            ListParams.Append(GetSession(Item.Cells("FromName").Value, Item.Cells("ParamType").Value, defVal))
                        End If
                    Case "Request"
                        If ObjectDataSource = "" Then
                            ListParams.Append(GetRequest(Item.Cells("FromName").Value, Item.Cells("ParamType").Value, defVal))
                        End If
                End Select
            Next
            If ListParams.ToString() <> "" Then
                strSET.Append(GetRegion("把计"))
                strSET.Append(ListParams.ToString())
                strSET.Append(EndRegion())
            End If
        End If

        Select Case CodeType
            Case "cs"
                strSET.Append("    protected void Page_Load(object sender, EventArgs e)" & GetLine())
                strSET.Append("    {" & GetLine())
                strSET.Append("        if (!Page.IsPostBack)" & GetLine())
                strSET.Append("        {" & GetLine())
                If haveMultiView = True Then
                    strSET.Append("            MultiView1.ActiveViewIndex = 0;" & GetLine())
                End If
                If haveBindGV = True And ObjectDataSource = "" Then
                    strSET.Append(String.Format("            {0}();", BindGV) & GetLine())
                End If
                strSET.Append("        }" & GetLine())
                strSET.Append("    }" & GetLine())
            Case "vb"
                strSET.Append("    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load" & GetLine())
                strSET.Append("        If Not Page.IsPostBack Then" & GetLine())
                If haveMultiView = True Then
                    strSET.Append("            MultiView1.ActiveViewIndex = 0" & GetLine())
                End If
                If haveBindGV = True And ObjectDataSource = "" Then
                    strSET.Append(String.Format("            {0}()", BindGV) & GetLine())
                End If
                strSET.Append("        End If" & GetLine())
                strSET.Append("    End Sub" & GetLine())
        End Select

        If GridView <> "" Then
            If ObjectDataSource = "" And PageSize > 0 Then
                Select Case CodeType
                    Case "cs"
                        strSET.Append(String.Format("    protected void {0}_PageIndexChanging(object sender, GridViewPageEventArgs e)", GridView) & GetLine())
                        strSET.Append("    {" & GetLine())
                        strSET.Append(String.Format("        {0}.PageIndex = e.NewPageIndex;", GridView) & GetLine())
                        strSET.Append(String.Format("        {0}();", BindGV) & GetLine())
                        strSET.Append("    }" & GetLine())
                    Case "vb"
                        strSET.Append(String.Format("    Protected Sub {0}_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles {0}.PageIndexChanging", GridView) & GetLine())
                        strSET.Append(String.Format("        {0}.PageIndex = e.NewPageIndex", GridView) & GetLine())
                        strSET.Append(String.Format("        {0}()", BindGV) & GetLine())
                        strSET.Append("    End Sub" & GetLine())
                End Select
            End If

            If GridView <> "" And (haveListOnly = False Or Me.haveDEL = True) Then
                Select Case CodeType
                    Case "cs"
                        strSET.Append(String.Format("    protected void {0}_RowCommand(object sender, GridViewCommandEventArgs e)", GridView) & GetLine())
                        strSET.Append("    {" & GetLine())
                        strSET.Append("        switch(e.CommandName)" & GetLine())
                        strSET.Append("        {" & GetLine())
                        If haveListOnly = False Then
                            strSET.Append("            case ""read"":" & GetLine())

                            If Me.haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 1;" & GetLine())
                            End If

                            If mdFile.ArrayKEY.Length > 1 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    strSET.Append(String.Format("                {0} = {1};", ColumnProperty(mdFile.ArrayKEY(i)), ParseType(String.Format("{0}.DataKeys[int.Parse(e.CommandArgument.ToString())].Values[{1}]", GridView, i), mdFile.classDT.Columns(mdFile.ArrayKEY(i)).DataType.Name)) & GetLine())
                                Next
                            ElseIf mdFile.ArrayKEY.Length = 1 Then
                                strSET.Append(String.Format("                {0} = {1};", ColumnProperty(mdFile.ArrayKEY(0)), ParseType(String.Format("{0}.DataKeys[int.Parse(e.CommandArgument.ToString())].Value", GridView), mdFile.classDT.Columns(mdFile.ArrayKEY(0)).DataType.Name)) & GetLine())
                            End If

                            strSET.Append(String.Format("                {0}();", ReadData) & GetLine())
                            strSET.Append("                break;" & GetLine())
                        End If
                        If haveDEL = True Then
                            strSET.Append("            case ""del"":" & GetLine())
                            strSET.Append(String.Format("                if ({0}Biz.Del(", mdFile.strClass))
                            If mdFile.ArrayKEY.Length > 1 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append(", ")
                                    End If
                                    strSET.Append(ParseType(String.Format("{0}.DataKeys[int.Parse(e.CommandArgument.ToString())].Values[{1}]", GridView, i), mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name))
                                Next
                            ElseIf mdFile.ArrayKEY.Length = 1 Then
                                strSET.Append(ParseType(String.Format("{0}.DataKeys[int.Parse(e.CommandArgument.ToString())].Value", GridView), mdFile.classDT.Columns(mdFile.ArrayKEY(0).ToString()).DataType.Name))
                            End If
                            strSET.Append(") == true)" & GetLine())
                            strSET.Append("                {" & GetLine())
                            strSET.Append("                    //埃Θ " & GetLine())
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                    {0}();", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                    {0}.DataBind();", GridView) & GetLine())
                            End If
                            strSET.Append("                }" & GetLine())
                            strSET.Append("                break;" & GetLine())
                        End If
                        strSET.Append("        }" & GetLine())
                        strSET.Append("    }" & GetLine())
                    Case "vb"
                        strSET.Append(String.Format("    Protected Sub {0}_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles {0}.RowCommand", GridView) & GetLine())
                        strSET.Append("        Select Case e.CommandName" & GetLine())
                        If haveListOnly = False Then
                            strSET.Append("            Case ""read""" & GetLine())

                            If Me.haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 1" & GetLine())
                            End If

                            If mdFile.ArrayKEY.Length > 1 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    strSET.Append(String.Format("                {0} = {1}", ColumnProperty(mdFile.ArrayKEY(i)), String.Format("{0}.DataKeys(e.CommandArgument).Values({1})", GridView, i)) & GetLine())
                                Next
                            ElseIf mdFile.ArrayKEY.Length = 1 Then
                                strSET.Append(String.Format("                {0} = {1}", ColumnProperty(mdFile.ArrayKEY(0)), String.Format("{0}.DataKeys(e.CommandArgument).Value", GridView)) & GetLine())
                            End If

                            strSET.Append(String.Format("                {0}()", ReadData) & GetLine())
                        End If
                        If haveDEL = True Then
                            strSET.Append("            Case ""del""" & GetLine())
                            strSET.Append(String.Format("                If {0}Biz.Del(", strClass))
                            If mdFile.ArrayKEY.Length > 1 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append(", ")
                                    End If
                                    strSET.Append(String.Format("{0}.DataKeys(e.CommandArgument).Values({1})", GridView, i))
                                Next
                            ElseIf mdFile.ArrayKEY.Length = 1 Then
                                strSET.Append(String.Format("{0}.DataKeys(e.CommandArgument).Value", GridView))
                            End If
                            strSET.Append(") = True Then" & GetLine())
                            strSET.Append("                    '埃Θ" & GetLine())
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                    {0}()", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                    {0}.DataBind()", GridView) & GetLine())
                            End If
                            strSET.Append("                End If" & GetLine())
                        End If
                        strSET.Append("        End Select" & GetLine())
                        strSET.Append("    End Sub" & GetLine())
                End Select
            End If
        End If

        Dim haveQuery As Boolean = False
        If haveAddOnly = False Then
            For Each Item As DataGridViewRow In DataGridView1.Rows
                Select Case Item.Cells("FromType").Value
                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                        haveQuery = True
                End Select
            Next
        End If
        If haveQuery = True Then
            Dim defVal As String
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    protected void {0}_Click(object sender, EventArgs e)", btnQuery) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If ObjectDataSource = "" Then
                        For Each Item As DataGridViewRow In DataGridView1.Rows
                            If Item.Cells("defVal").Value = "" Then
                                defVal = GetDefValue(Item.Cells("ParamType").Value)
                            Else
                                defVal = Item.Cells("defVal").Value
                            End If
                            Select Case Item.Cells("FromType").Value
                                Case "DropDownList", "ListBox", "RadioButtonList"
                                    strSET.Append(String.Format("        if ({0}.SelectedValue == """")", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)) & GetLine())
                                    strSET.Append("        {" & GetLine())
                                    strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, defVal) & GetLine())
                                    strSET.Append("        }" & GetLine())
                                    strSET.Append("        else" & GetLine())
                                    strSET.Append("        {" & GetLine())
                                    strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, ParseType(String.Format("{0}.SelectedValue", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)), Item.Cells("ParamType").Value, "")) & GetLine())
                                    strSET.Append("        }" & GetLine())
                                Case "TextBox"
                                    strSET.Append(String.Format("        if ({0}.Text.Trim() == """")", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)) & GetLine())
                                    strSET.Append("        {" & GetLine())
                                    strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, defVal) & GetLine())
                                    strSET.Append("        }" & GetLine())
                                    strSET.Append("        else" & GetLine())
                                    strSET.Append("        {" & GetLine())

                                    Dim dataType As String = Item.Cells("ParamType").Value
                                    If dataType = "String" Then
                                        strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, String.Format("{0}.Text", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value))) & GetLine())
                                    Else
                                        strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, ParseType(String.Format("{0}.Text", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)), dataType, "")) & GetLine())
                                    End If

                                    strSET.Append("        }" & GetLine())
                                Case "CheckBox"
                                    Dim dataType As String = Item.Cells("ParamType").Value
                                    If dataType = "Boolean" Then
                                        strSET.Append(String.Format("            {0} = {1};", Item.Cells("FromName").Value, String.Format("{0}.Checked", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value))) & GetLine())
                                    Else
                                        strSET.Append(String.Format("        {0} = {1};", Item.Cells("FromName").Value, ParseType(String.Format("{0}.Checked", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)), Item.Cells("ParamType").Value)) & GetLine())
                                    End If
                            End Select
                        Next
                    End If
                    strSET.Append(String.Format("        {0}();", BindGV) & GetLine())
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Protected Sub {0}_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles {0}.Click", btnQuery) & GetLine())
                    If ObjectDataSource = "" Then
                        For Each Item As DataGridViewRow In DataGridView1.Rows
                            If Item.Cells("defVal").Value = "" Then
                                defVal = GetDefValue(Item.Cells("ParamType").Value)
                            Else
                                defVal = Item.Cells("defVal").Value
                            End If
                            Select Case Item.Cells("FromType").Value
                                Case "DropDownList", "ListBox", "RadioButtonList"
                                    strSET.Append(String.Format("        If {0}.SelectedValue = """" Then", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)) & GetLine())
                                    strSET.Append(String.Format("            {0} = {1}", Item.Cells("FromName").Value, defVal) & GetLine())
                                    strSET.Append("                Else" & GetLine())
                                    strSET.Append(String.Format("            {0} = {1}", Item.Cells("FromName").Value, String.Format("{0}.SelectedValue", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value))) & GetLine() & GetLine())
                                    strSET.Append("        End If" & GetLine())
                                Case "TextBox"
                                    strSET.Append(String.Format("        If {0}.Text.Trim = """" Then", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value)) & GetLine())
                                    strSET.Append(String.Format("            {0} = {1}", Item.Cells("FromName").Value, defVal) & GetLine())
                                    strSET.Append("                Else" & GetLine())
                                    strSET.Append(String.Format("            {0} = {1}", Item.Cells("FromName").Value, String.Format("{0}.Text", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value))) & GetLine())
                                    strSET.Append("        End If" & GetLine())
                                Case "CheckBox"
                                    strSET.Append(String.Format("        {0} = {1}", Item.Cells("FromName").Value, String.Format("{0}.Checked", GetControlName(Item.Cells("FromType").Value, Item.Cells("FromName").Value))) & GetLine() & GetLine())
                            End Select
                        Next
                    End If
                    strSET.Append(String.Format("        {0}()", BindGV) & GetLine())
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If btnAdd <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    protected void {0}_Click(object sender, EventArgs e)", btnAdd) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If haveMultiView = True Then
                        strSET.Append("        MultiView1.ActiveViewIndex = 1;" & GetLine())
                    End If
                    strSET.Append(String.Format("        {0}();", CleanData) & GetLine())
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Protected Sub {0}_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles {0}.Click", btnAdd) & GetLine())
                    If haveMultiView = True Then
                        strSET.Append("        MultiView1.ActiveViewIndex = 1" & GetLine())
                    End If
                    strSET.Append(String.Format("        {0}()", CleanData) & GetLine())
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If btnSave <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    protected void {0}_Click(object sender, EventArgs e)", btnSave) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If Me.haveAddOnly = True Then
                        strSET.Append(String.Format("        {0}Info {0} = new {0}Info();", mdFile.strClass) & GetLine())
                        If HasAutoIncrement = "" And mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                For Each Item As DataGridViewRow In DataGridView2.Rows
                                    If Item.Cells("chkDOC").Value = True Then
                                        If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                            Dim dataVal As String = String.Format("{0}.Text", Item.Cells("EditName").Value)
                                            strSET.Append(String.Format("        {0}.{1} = {2};", mdFile.strClass, ColumnProperty(mdFile.ArrayKEY(i)), dataVal) & GetLine())
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    Else
                        strSET.Append(String.Format("        {0}Info {0};", mdFile.strClass) & GetLine())
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        if ({0} > 0)", ColumnProperty(mdFile.ArrayKEY(0))) & GetLine())
                        Else
                            strSET.Append("        if (")
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append("&& ")
                                    End If
                                    If GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name) = "0" Then
                                        strSET.Append(String.Format("{0} > 0 ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    Else
                                        strSET.Append(String.Format("{0} != """" ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    End If
                                Next
                            End If
                            strSET.Append(")" & GetLine())
                        End If
                        strSET.Append("        {" & GetLine())
                        strSET.Append(String.Format("            {0} = {0}Biz.GetInfo({1});", mdFile.strClass, mdFile.strKEY) & GetLine())
                        strSET.Append("        }" & GetLine())
                        strSET.Append("        else" & GetLine())
                        strSET.Append("        {" & GetLine())
                        strSET.Append(String.Format("            {0} = new {0}Info();", mdFile.strClass) & GetLine())
                        If HasAutoIncrement = "" And mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                For Each Item As DataGridViewRow In DataGridView2.Rows
                                    If Item.Cells("chkDOC").Value = True Then
                                        If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then

                                            Dim dataType As String = mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name
                                            Dim dataVal As String
                                            Select Case dataType
                                                Case "String"
                                                    dataVal = String.Format("{0}.Text", Item.Cells("EditName").Value)
                                                Case Else
                                                    dataVal = ParseType(String.Format("{0}.Text", Item.Cells("EditName").Value), dataType, "")
                                            End Select

                                            strSET.Append(String.Format("            {0}.{1} = {2};", mdFile.strClass, ColumnProperty(mdFile.ArrayKEY(i)), dataVal) & GetLine())
                                        End If
                                    End If
                                Next
                            Next
                        End If
                        strSET.Append("        }" & GetLine())
                    End If

                    Dim hasEdit As Boolean
                    For Each Item As DataGridViewRow In DataGridView2.Rows
                        hasEdit = True
                        For i = 0 To mdFile.ArrayKEY.Length - 1
                            If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                hasEdit = False
                            End If
                        Next
                        If hasEdit = True Then
                            If Item.Cells("chkDOC").Value = True Then
                                Dim dataType As String = mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name
                                Dim valType As String
                                Dim dataVal As String
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        valType = "Text"
                                    Case "CheckBox"
                                        valType = "Checked"
                                    Case Else
                                        valType = "SelectedValue"
                                End Select
                                Select Case dataType
                                    Case "String", "Boolean"
                                        dataVal = String.Format("{0}.{1}", Item.Cells("EditName").Value, valType)
                                    Case Else
                                        dataVal = ParseType(String.Format("{0}.{1}", Item.Cells("EditName").Value, valType), dataType, "")
                                End Select
                                strSET.Append(String.Format("        {0}.{1} = {2};", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value), dataVal) & GetLine())
                            Else
                                strSET.Append(String.Format("        //{0}.{1} = ;", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                            End If
                        End If
                    Next

                    If Me.haveAddOnly = True Then
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        if ({0}Biz.AddNew({1}) == true)", mdFile.strClass, GetRefParam(strClass)) & GetLine())
                        Else
                            strSET.Append(String.Format("        if ({0}Biz.AddNew({1}) == true)", mdFile.strClass, strClass) & GetLine())
                        End If
                        strSET.Append("        {" & GetLine())
                        strSET.Append("            //穝糤Θ" & GetLine())
                        If haveMultiView = True Then
                            strSET.Append("            MultiView1.ActiveViewIndex = 0;" & GetLine())
                        End If
                        If GridView <> "" Then
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("            {0}();", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("            {0}.DataBind();", GridView) & GetLine())
                            End If
                        End If
                        strSET.Append("        }" & GetLine())
                    Else
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        if ({0} > 0)", ColumnProperty(HasAutoIncrement)) & GetLine())
                            strSET.Append("        {" & GetLine())
                            strSET.Append(String.Format("            if ({0}Biz.Update({0}) == true)", mdFile.strClass) & GetLine())
                            strSET.Append("            {" & GetLine())
                            strSET.Append("                //эΘ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0;" & GetLine())
                            End If
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                {0}();", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                {0}.DataBind();", GridView) & GetLine())
                            End If
                            strSET.Append("            }" & GetLine())
                            strSET.Append("        }" & GetLine())
                            strSET.Append("        else" & GetLine())
                            strSET.Append("        {" & GetLine())
                            strSET.Append(String.Format("            if ({0}Biz.AddNew({1}) == true)", mdFile.strClass, GetRefParam(strClass)) & GetLine())
                            strSET.Append("            {" & GetLine())
                            strSET.Append("                //穝糤Θ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0;" & GetLine())
                            End If
                            If GridView <> "" Then
                                If Me.ObjectDataSource = "" Then
                                    strSET.Append(String.Format("                {0}();", BindGV) & GetLine())
                                Else
                                    strSET.Append(String.Format("                {0}.DataBind();", GridView) & GetLine())
                                End If
                            End If
                            strSET.Append("            }" & GetLine())
                            strSET.Append("        }" & GetLine())
                        Else
                            strSET.Append("        if (")
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append(" && ")
                                    End If
                                    If GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name) = "0" Then
                                        strSET.Append(String.Format("{0} > 0", ColumnProperty(mdFile.ArrayKEY(i))))
                                    Else
                                        strSET.Append(String.Format("{0} != """"", ColumnProperty(mdFile.ArrayKEY(i))))
                                    End If
                                Next
                            End If
                            strSET.Append(")" & GetLine())
                            strSET.Append("        {" & GetLine())
                            strSET.Append(String.Format("            if ({0}Biz.Update({0}) == true)", mdFile.strClass) & GetLine())
                            strSET.Append("            {" & GetLine())
                            strSET.Append("                //эΘ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0;" & GetLine())
                            End If
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                {0}();", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                {0}.DataBind();", GridView) & GetLine())
                            End If
                            strSET.Append("            }" & GetLine())
                            strSET.Append("        }" & GetLine())
                            strSET.Append("        else" & GetLine())
                            strSET.Append("        {" & GetLine())
                            strSET.Append(String.Format("            if ({0}Biz.Exists(", mdFile.strClass))
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append(", ")
                                    End If
                                    For Each Item As DataGridViewRow In DataGridView2.Rows
                                        If Item.Cells("chkDOC").Value = True Then
                                            If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                                Dim dataType As String = mdFile.classDT.Columns(mdFile.ArrayKEY(i)).DataType.Name()
                                                If dataType <> "String" Then
                                                    strSET.Append(ParseType(String.Format("{0}.Text", Item.Cells("EditName").Value), dataType, ""))
                                                Else
                                                    strSET.Append(String.Format("{0}.Text", Item.Cells("EditName").Value))
                                                End If                                                
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                            strSET.Append(") == false)" & GetLine())
                            strSET.Append("            {" & GetLine())
                            strSET.Append(String.Format("                if ({0}Biz.AddNew({0}) == true)", mdFile.strClass) & GetLine())
                            strSET.Append("                {" & GetLine())
                            strSET.Append("                    //穝糤Θ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0;" & GetLine())
                            End If
                            If GridView <> "" Then
                                If Me.ObjectDataSource = "" Then
                                    strSET.Append(String.Format("                {0}();", BindGV) & GetLine())
                                Else
                                    strSET.Append(String.Format("                {0}.DataBind();", GridView) & GetLine())
                                End If
                            End If
                            strSET.Append("                }" & GetLine())
                            strSET.Append("            }" & GetLine())
                            strSET.Append("        }" & GetLine())
                        End If
                    End If
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Protected Sub {0}_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles {0}.Click", btnSave) & GetLine())
                    If Me.haveAddOnly = True Then
                        strSET.Append(String.Format("        Dim {0} As New {0}Info", mdFile.strClass) & GetLine())
                        If HasAutoIncrement = "" And mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                For Each Item As DataGridViewRow In DataGridView2.Rows
                                    If Item.Cells("chkDOC").Value = True Then
                                        If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                            strSET.Append(String.Format("        {0}.{1} = {2}.Text", mdFile.strClass, ColumnProperty(mdFile.ArrayKEY(i)), Item.Cells("EditName").Value) & GetLine())
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    Else
                        strSET.Append(String.Format("        Dim {0} As {0}Info", mdFile.strClass) & GetLine())
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        If {0} > 0 Then", ColumnProperty(mdFile.ArrayKEY(0))) & GetLine())
                        Else
                            strSET.Append("        If ")
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append("And ")
                                    End If
                                    If GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name) = "0" Then
                                        strSET.Append(String.Format("{0} > 0 ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    Else
                                        strSET.Append(String.Format("{0} <> """" ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    End If
                                Next
                            End If
                            strSET.Append("Then" & GetLine())
                        End If
                        strSET.Append(String.Format("            {0} = {0}Biz.GetInfo({1})", mdFile.strClass, mdFile.strKEY) & GetLine())
                        strSET.Append("        Else" & GetLine())
                        strSET.Append(String.Format("            {0} = New {0}Info", mdFile.strClass) & GetLine())
                        If HasAutoIncrement = "" And mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                For Each Item As DataGridViewRow In DataGridView2.Rows
                                    If Item.Cells("chkDOC").Value = True Then
                                        If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                            strSET.Append(String.Format("            {0}.{1} = {2}.Text", mdFile.strClass, ColumnProperty(mdFile.ArrayKEY(i)), Item.Cells("EditName").Value) & GetLine())
                                        End If
                                    End If
                                Next
                            Next
                        End If
                        strSET.Append("        End If" & GetLine())
                    End If

                    Dim hasEdit As Boolean
                    For Each Item As DataGridViewRow In DataGridView2.Rows
                        hasEdit = True
                        For i = 0 To mdFile.ArrayKEY.Length - 1
                            If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                hasEdit = False
                            End If
                        Next
                        If hasEdit = True Then
                            If Item.Cells("chkDOC").Value = True Then
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        strSET.Append(String.Format("        {0}.{1} = {2}.Text", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value), Item.Cells("EditName").Value) & GetLine())
                                    Case "CheckBox"
                                        strSET.Append(String.Format("        {0}.{1} = {2}.Checked", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value), Item.Cells("EditName").Value) & GetLine())
                                    Case Else
                                        strSET.Append(String.Format("        {0}.{1} = {2}.SelectedValue", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value), Item.Cells("EditName").Value) & GetLine())
                                End Select
                            Else
                                strSET.Append(String.Format("        '{0}.{1} = ", mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                            End If
                        End If
                    Next

                    If Me.haveAddOnly = True Then
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        If {0}Biz.AddNew({1}) = True Then", mdFile.strClass, GetRefParam(strClass)) & GetLine())
                        Else
                            strSET.Append(String.Format("        If {0}Biz.AddNew({1}) = True Then", mdFile.strClass, strClass) & GetLine())
                        End If
                        strSET.Append("            '穝糤Θ" & GetLine())
                        If haveMultiView = True Then
                            strSET.Append("            MultiView1.ActiveViewIndex = 0" & GetLine())
                        End If
                        If GridView <> "" Then
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("            {0}()", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("            {0}.DataBind()", GridView) & GetLine())
                            End If
                        End If
                        strSET.Append("        End If" & GetLine())
                    Else
                        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
                            strSET.Append(String.Format("        If {0} > 0 Then", ColumnProperty(HasAutoIncrement)) & GetLine())
                            strSET.Append(String.Format("            If {0}Biz.Update({0}) = True Then", mdFile.strClass) & GetLine())
                            strSET.Append("                'эΘ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0" & GetLine())
                            End If
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                {0}()", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                {0}.DataBind()", GridView) & GetLine())
                            End If
                            strSET.Append("            End If" & GetLine())
                            strSET.Append("        Else" & GetLine())
                            strSET.Append(String.Format("            If {0}Biz.AddNew({1}) = True Then", mdFile.strClass, GetRefParam(strClass)) & GetLine())
                            strSET.Append("                '穝糤Θ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0" & GetLine())
                            End If
                            If GridView <> "" Then
                                If Me.ObjectDataSource = "" Then
                                    strSET.Append(String.Format("                {0}()", BindGV) & GetLine())
                                Else
                                    strSET.Append(String.Format("                {0}.DataBind()", GridView) & GetLine())
                                End If
                            End If
                            strSET.Append("            End If" & GetLine())
                            strSET.Append("        End If" & GetLine())
                        Else
                            strSET.Append("        If ")
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append("And ")
                                    End If
                                    If GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name) = "0" Then
                                        strSET.Append(String.Format("{0} > 0 ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    Else
                                        strSET.Append(String.Format("{0} <> """" ", ColumnProperty(mdFile.ArrayKEY(i))))
                                    End If
                                Next
                            End If
                            strSET.Append("Then" & GetLine())
                            strSET.Append(String.Format("            If {0}Biz.Update({0}) = True Then", mdFile.strClass) & GetLine())
                            strSET.Append("                'эΘ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0" & GetLine())
                            End If
                            If Me.ObjectDataSource = "" Then
                                strSET.Append(String.Format("                {0}()", BindGV) & GetLine())
                            Else
                                strSET.Append(String.Format("                {0}.DataBind()", GridView) & GetLine())
                            End If
                            strSET.Append("            End If" & GetLine())
                            strSET.Append("        Else" & GetLine())
                            strSET.Append(String.Format("            If {0}Biz.Exists(", mdFile.strClass))
                            If mdFile.ArrayKEY.Length > 0 Then
                                For i = 0 To mdFile.ArrayKEY.Length - 1
                                    If i > 0 Then
                                        strSET.Append(", ")
                                    End If
                                    For Each Item As DataGridViewRow In DataGridView2.Rows
                                        If Item.Cells("chkDOC").Value = True Then
                                            If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                                strSET.Append(String.Format("{0}.Text", Item.Cells("EditName").Value))
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                            strSET.Append(") = False Then" & GetLine())
                            strSET.Append(String.Format("                If {0}Biz.AddNew({0}) = True Then", mdFile.strClass) & GetLine())
                            strSET.Append("                    '穝糤Θ" & GetLine())
                            If haveMultiView = True Then
                                strSET.Append("                MultiView1.ActiveViewIndex = 0" & GetLine())
                            End If
                            If GridView <> "" Then
                                If Me.ObjectDataSource = "" Then
                                    strSET.Append(String.Format("                {0}()", BindGV) & GetLine())
                                Else
                                    strSET.Append(String.Format("                {0}.DataBind()", GridView) & GetLine())
                                End If
                            End If
                            strSET.Append("                End If" & GetLine())
                            strSET.Append("            End If" & GetLine())
                            strSET.Append("        End If" & GetLine())
                        End If
                    End If
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If btnBack <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    protected void {0}_Click(object sender, EventArgs e)", btnBack) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If haveMultiView = True Then
                        strSET.Append("        MultiView1.ActiveViewIndex = 0;" & GetLine())
                    End If
                    If CleanData <> "" Then
                        strSET.Append(String.Format("        {0}();", CleanData) & GetLine())
                    End If
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Protected Sub {0}_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles {0}.Click", btnBack) & GetLine())
                    If haveMultiView = True Then
                        strSET.Append("        MultiView1.ActiveViewIndex = 0" & GetLine())
                    End If
                    If CleanData <> "" Then
                        strSET.Append(String.Format("        {0}()", CleanData) & GetLine())
                    End If
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If ReadData <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    private void {0}()", ReadData) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If btnSave <> "" Then
                        strSET.Append(String.Format("        {0}.Text = ""э"";", btnSave) & GetLine())
                    End If
                    strSET.Append(String.Format("        {0}Info {0} = {0}Biz.GetInfo({1});", mdFile.strClass, mdFile.strKEY) & GetLine())

                    If haveReadOnly = False Then
                        Dim hasEdit As Boolean
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            hasEdit = True
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                    hasEdit = False
                                End If
                            Next
                            If Item.Cells("chkDOC").Value = True Then
                                Dim dataType As String = mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name()
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        strSET.Append(String.Format("        {0}.Text = {1}.{2};", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value) & ToStringCS(dataType)) & GetLine())
                                    Case "CheckBox"
                                        strSET.Append(String.Format("        {0}.Checked = {1}.{2};", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                                    Case Else
                                        strSET.Append(String.Format("        {0}.SelectedValue = {1}.{2};", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value) & ToStringCS(dataType)) & GetLine())
                                End Select
                                If hasEdit = False Then
                                    strSET.Append(String.Format("        {0}.ReadOnly = true;", Item.Cells("EditName").Value) & GetLine())
                                End If
                            End If
                        Next
                    Else
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            If Item.Cells("chkDOC").Value = True Then
                                strSET.Append(String.Format("        {0}.Text = {1}.{2};", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value) & ToStringCS(mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name())) & GetLine())
                            End If
                        Next
                    End If
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Private Sub {0}()", ReadData) & GetLine())
                    If btnSave <> "" Then
                        strSET.Append(String.Format("        {0}.Text = ""э""", btnSave) & GetLine())
                    End If
                    strSET.Append(String.Format("        Dim {0} As {0}Info = {0}Biz.GetInfo({1})", mdFile.strClass, mdFile.strKEY) & GetLine())

                    If haveReadOnly = False Then
                        Dim hasEdit As Boolean
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            hasEdit = True
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                    hasEdit = False
                                End If
                            Next
                            If Item.Cells("chkDOC").Value = True Then
                                Dim dataType As String = mdFile.classDT.Columns(Item.Cells("FieldName").Value).DataType.Name()
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        strSET.Append(String.Format("        {0}.Text = {1}.{2}", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                                    Case "CheckBox"
                                        strSET.Append(String.Format("        {0}.Checked = {1}.{2}", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                                    Case Else
                                        strSET.Append(String.Format("        {0}.SelectedValue = {1}.{2}", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                                End Select
                                If hasEdit = False Then
                                    strSET.Append(String.Format("        {0}.ReadOnly = True", Item.Cells("EditName").Value) & GetLine())
                                End If
                            End If
                        Next
                    Else
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            If Item.Cells("chkDOC").Value = True Then
                                strSET.Append(String.Format("        {0}.Text = {1}.{2}", Item.Cells("EditName").Value, mdFile.strClass, ColumnProperty(Item.Cells("FieldName").Value)) & GetLine())
                            End If
                        Next
                    End If
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If CleanData <> "" Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    private void {0}()", CleanData) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If btnSave <> "" And ReadData <> "" Then
                        strSET.Append(String.Format("        {0}.Text = ""穝糤"";", btnSave) & GetLine())
                    End If

                    If ReadData <> "" Then
                        If mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                strSET.Append(String.Format("        {0} = {1};", ColumnProperty(mdFile.ArrayKEY(i)), GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name)) & GetLine())
                            Next
                        End If
                    End If

                    If haveReadOnly = False Then
                        Dim hasEdit As Boolean
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            hasEdit = True
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                    hasEdit = False
                                End If
                            Next
                            If Item.Cells("chkDOC").Value = True Then
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        strSET.Append(String.Format("        {0}.Text = """";", Item.Cells("EditName").Value) & GetLine())
                                    Case "CheckBox"
                                        strSET.Append(String.Format("        {0}.Checked = {1};", Item.Cells("EditName").Value, ConvertValue("False")) & GetLine())
                                    Case Else
                                        strSET.Append(String.Format("        {0}.SelectedIndex = -1;", Item.Cells("EditName").Value) & GetLine())
                                End Select
                                If hasEdit = False And ReadData <> "" Then
                                    strSET.Append(String.Format("        {0}.ReadOnly = false;", Item.Cells("EditName").Value) & GetLine())
                                End If
                            End If
                        Next
                    Else
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            If Item.Cells("chkDOC").Value = True Then
                                strSET.Append(String.Format("        {0}.Text = """";", Item.Cells("EditName").Value) & GetLine())
                            End If
                        Next
                    End If
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Private Sub {0}()", CleanData) & GetLine())
                    If btnSave <> "" And ReadData <> "" Then
                        strSET.Append(String.Format("        {0}.Text = ""穝糤""", btnSave) & GetLine())
                    End If

                    If ReadData <> "" Then
                        If mdFile.ArrayKEY.Length > 0 Then
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                strSET.Append(String.Format("        {0} = {1}", ColumnProperty(mdFile.ArrayKEY(i)), GetDefValue(mdFile.classDT.Columns(mdFile.ArrayKEY(i).ToString()).DataType.Name)) & GetLine())
                            Next
                        End If
                    End If

                    If haveReadOnly = False Then
                        Dim hasEdit As Boolean
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            hasEdit = True
                            For i = 0 To mdFile.ArrayKEY.Length - 1
                                If mdFile.ArrayKEY(i) = Item.Cells("FieldName").Value Then
                                    hasEdit = False
                                End If
                            Next
                            If Item.Cells("chkDOC").Value = True Then
                                Select Case Item.Cells("EditType").Value
                                    Case "TextBox"
                                        strSET.Append(String.Format("        {0}.Text = """"", Item.Cells("EditName").Value) & GetLine())
                                    Case "CheckBox"
                                        strSET.Append(String.Format("        {0}.Checked = {1}", Item.Cells("EditName").Value, ConvertValue("False")) & GetLine())
                                    Case Else
                                        strSET.Append(String.Format("        {0}.SelectedIndex = -1", Item.Cells("EditName").Value) & GetLine())
                                End Select
                                If hasEdit = False And ReadData <> "" Then
                                    strSET.Append(String.Format("        {0}.ReadOnly = False", Item.Cells("EditName").Value) & GetLine())
                                End If
                            End If
                        Next
                    Else
                        For Each Item As DataGridViewRow In DataGridView2.Rows
                            If Item.Cells("chkDOC").Value = True Then
                                strSET.Append(String.Format("        {0}.Text = """"", Item.Cells("EditName").Value) & GetLine())
                            End If
                        Next
                    End If
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If

        If haveBindGV = True Then
            Select Case CodeType
                Case "cs"
                    strSET.Append(String.Format("    private void {0}()", BindGV) & GetLine())
                    strSET.Append("    {" & GetLine())
                    If ObjectDataSource <> "" Then
                        For Each Item As DataGridViewRow In DataGridView1.Rows
                            Select Case Item.Cells("FromType").Value
                                Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                                    If Item.Cells("ParamType").Value = "String" Then
                                        strSET.Append(String.Format("        {0}.SelectParameters[""{1}""].DefaultValue = {1};", ObjectDataSource, Item.Cells("ParamName").Value, Item.Cells("FromName").Value) & GetLine())
                                    Else
                                        strSET.Append(String.Format("        {0}.SelectParameters[""{1}""].DefaultValue = {1}.ToString();", ObjectDataSource, Item.Cells("ParamName").Value, Item.Cells("FromName").Value) & GetLine())
                                    End If
                                Case "other"
                                    strSET.Append(String.Format("        //{0}.SelectParameters[""{1}""].DefaultValue = ;", ObjectDataSource, Item.Cells("ParamName").Value) & GetLine())
                            End Select
                        Next
                    Else
                        strSET.Append(String.Format("        {0}.DataSource = {1}Biz.{2}(", GridView, mdFile.strClass, ComboBox2.SelectedItem("Name")))
                        If DataGridView1.Rows.Count > 0 Then
                            For i = 0 To DataGridView1.Rows.Count - 1
                                If i > 0 Then
                                    strSET.Append(", ")
                                End If
                                Select Case DataGridView1.Rows(i).Cells("FromType").Value
                                    Case "const"
                                        strSET.Append(DataGridView1.Rows(i).Cells("defVal").Value)
                                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                                        strSET.Append(DataGridView1.Rows(i).Cells("FromName").Value)
                                    Case "Request", "Session"
                                        strSET.Append(DataGridView1.Rows(i).Cells("FromName").Value)
                                    Case "other"
                                        strSET.Append("""?""")
                                End Select
                            Next
                        End If
                        strSET.Append(");" & GetLine())
                    End If
                    If Me.ObjectDataSource = "" Then
                        strSET.Append(String.Format("        {0}.DataBind();", GridView) & GetLine())
                    End If
                    strSET.Append("    }" & GetLine())
                Case "vb"
                    strSET.Append(String.Format("    Private Sub {0}()", BindGV) & GetLine())
                    If ObjectDataSource <> "" Then
                        For Each Item As DataGridViewRow In DataGridView1.Rows
                            Select Case Item.Cells("FromType").Value
                                Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                                    strSET.Append(String.Format("        {0}.SelectParameters(""{1}"").DefaultValue = {1}", ObjectDataSource, Item.Cells("ParamName").Value, Item.Cells("FromName").Value) & GetLine())
                                Case "other"
                                    strSET.Append(String.Format("        '{0}.SelectParameters(""{1}"").DefaultValue = ", ObjectDataSource, Item.Cells("ParamName").Value, Item.Cells("FromName").Value) & GetLine())
                            End Select
                        Next
                    Else
                        strSET.Append(String.Format("        {0}.DataSource = {1}Biz.{2}(", GridView, mdFile.strClass, ComboBox2.SelectedItem("Name")))
                        If DataGridView1.Rows.Count > 0 Then
                            For i = 0 To DataGridView1.Rows.Count - 1
                                If i > 0 Then
                                    strSET.Append(", ")
                                End If
                                Select Case DataGridView1.Rows(i).Cells("FromType").Value
                                    Case "const"
                                        strSET.Append(DataGridView1.Rows(i).Cells("defVal").Value)
                                    Case "DropDownList", "ListBox", "RadioButtonList", "TextBox", "CheckBox"
                                        strSET.Append(DataGridView1.Rows(i).Cells("FromName").Value)
                                    Case "Request", "Session"
                                        strSET.Append(DataGridView1.Rows(i).Cells("FromName").Value)
                                    Case "other"
                                        strSET.Append("""?""")
                                End Select
                            Next
                        End If
                        strSET.Append(")" & GetLine())
                    End If
                    If Me.ObjectDataSource = "" Then
                        strSET.Append(String.Format("        {0}.DataBind()", GridView) & GetLine())
                    End If
                    strSET.Append("    End Sub" & GetLine())
            End Select
        End If
        TextBox3.Text = strSET.ToString
    End Sub
End Class