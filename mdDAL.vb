Imports System.Data.Common
Imports System.Data.SQLite
Module mdDAL
#Region "Base Set"
    '專案外部資料庫
    Dim DbOutCom As DbCommand
    Dim DaOut As DbDataAdapter
    Private Sub OpenOutDbCom()
        If IsNothing(DbOutCom) Or IsChgDb Then
            DbOutCom = GetDBOutCommand()
            DbOutCom.Connection = GetDBOutConnection()
            DaOut = GetDBOutDataAdapter()
            DaOut.SelectCommand = DbOutCom
            IsChgDb = False
        End If
        DbOutCom.Connection.Open()
    End Sub
    Public Sub CloseOutDbCom()
        DbOutCom.Parameters.Clear()
        DbOutCom.Connection.Close()
    End Sub
    Private Function GetDBOutDataAdapter() As DbDataAdapter

        Select Case DbType
            'Case "MSSQL"
            '    Return New SqlClient.SqlDataAdapter
            Case "SQLITE"
                Return New SQLiteDataAdapter
            Case Else
                Return New SqlClient.SqlDataAdapter
        End Select

    End Function
    Private Function GetDBOutCommand() As DbCommand
        Select Case DbType
            Case "SQLITE"
                Return New SQLiteCommand
            Case Else
                Return New SqlClient.SqlCommand
        End Select
    End Function
    Private Function GetDBOutConnection() As DbConnection
        Select Case DbType
            Case "SQLITE"
                Return New SQLiteConnection(strConn)
            Case Else
                Return New SqlClient.SqlConnection(strConn)
        End Select
    End Function
    Private Function GetDBOutParameter(ByVal name As String, ByVal value As Object) As DbParameter
        Select Case DbType
            Case "SQLITE"
                Return New SQLiteParameter(name, value)
            Case Else
                Return New SqlClient.SqlParameter(name, value)
        End Select
    End Function

    '3TierDB GenCode內部資料庫
    Dim DbCom As DbCommand
    Dim Da As DbDataAdapter
    Private Sub OpenDbCom()
        If IsNothing(DbCom) Then
            DbCom = GetDBCommand()
            DbCom.Connection = GetDBConnection()
            Da = GetDBDataAdapter()
            Da.SelectCommand = DbCom
        End If
        DbCom.Connection.Open()
    End Sub
    Private Sub CloseDbCom()
        DbCom.Parameters.Clear()
        DbCom.Connection.Close()
    End Sub
    Private Function GetDBDataAdapter() As DbDataAdapter
        Return New OleDb.OleDbDataAdapter
        'Return New SqlClient.SqlDataAdapter
    End Function
    Private Function GetDBCommand() As DbCommand
        Return New OleDb.OleDbCommand
        'Return New SqlClient.SqlCommand
    End Function
    Private Function GetDBConnection() As DbConnection
        Return New OleDb.OleDbConnection(DbConn)
        'Return New SqlClient.SqlConnection(DbConn)
    End Function
    Private Function GetDBParameter(ByVal name As String, ByVal value As Object) As DbParameter
        Return New OleDb.OleDbParameter(name, value)
        'Return New SqlClient.SqlParameter(name, value)
    End Function

    '取得刪除開頭的Sql
    Dim _Del As String
    Private Function GetDel() As String
        If String.IsNullOrEmpty(_Del) Then
            _Del = "DELETE * "
            '_Del = "DELETE "
        End If
        Return _Del
    End Function
#End Region

#Region "DBSource"
    Public Function GetDataTable(ByVal strSQL As String) As DataTable
        OpenOutDbCom()
        DbOutCom.CommandText = strSQL
        Dim dtTable As New DataTable
        DaOut.Fill(dtTable)
        CloseOutDbCom()
        Return dtTable
    End Function

    Public Function GetSchema(ByVal strSQL As String) As DataTable
        OpenOutDbCom()
        DbOutCom.CommandText = strSQL
        Dim dtTable As New DataTable
        DaOut.FillSchema(dtTable, SchemaType.Source)
        CloseOutDbCom()
        Return dtTable
    End Function

    Public Function TablesSource() As DataTable
        Select Case DbType
            Case "SQLITE"
                Return GetDataTable("select name as TABLE_NAME from sqlite_master where type = 'table' and name <> 'sqlite_sequence'")
            Case Else
                Return GetDataTable("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME <> 'dtproperties'")
        End Select
    End Function

    Public Function TableSourceInfo(ByVal TableName As String) As DataTable
        OpenOutDbCom()
        Dim strSQL As New System.Text.StringBuilder

        Select Case DbType
            Case "SQLITE"
                strSQL.Append("select name as TABLE_NAME from sqlite_master where type = 'table' and name <> 'sqlite_sequence' ")
                strSQL.Append("and name = @TABLE_NAME")
            Case Else
                strSQL.Append("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ")
                strSQL.Append("AND TABLE_NAME = @TABLE_NAME")
        End Select

        DbOutCom.CommandText = strSQL.ToString
        DbOutCom.Parameters.Add(GetDBOutParameter("@TABLE_NAME", TableName))
        Dim dtTable As New DataTable
        DaOut.Fill(dtTable)
        CloseOutDbCom()
        Return dtTable
    End Function

    Public Function ColumnsSource(ByVal TableName As String) As DataTable
        OpenOutDbCom()
        Dim strSQL As New System.Text.StringBuilder

        Select Case DbType
            Case "SQLITE"
                strSQL.Append("PRAGMA table_info('" & TableName & "')")
            Case Else
                strSQL.Append("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TABLE_NAME")
        End Select

        DbOutCom.CommandText = strSQL.ToString
        DbOutCom.Parameters.Add(GetDBOutParameter("@TABLE_NAME", TableName))
        Dim dtTable As New DataTable
        DaOut.Fill(dtTable)
        CloseOutDbCom()
        Return dtTable
    End Function

    Public Function ColumnSourceInfo(ByVal TableName As String, ByVal ColumnName As String) As DataTable
        Select Case DbType
            Case "SQLITE"
                Dim dtTable As DataTable = ColumnsSource(TableName)
                dtTable.DefaultView.RowFilter = "name = '" & ColumnName & "'"
                Return dtTable.DefaultView.Table
            Case Else
                OpenOutDbCom()
                Dim strSQL As New System.Text.StringBuilder
                strSQL.Append("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TABLE_NAME ")
                strSQL.Append("AND COLUMN_NAME = @COLUMN_NAME")
                DbOutCom.CommandText = strSQL.ToString
                DbOutCom.Parameters.Add(GetDBOutParameter("@TABLE_NAME", TableName))
                DbOutCom.Parameters.Add(GetDBOutParameter("@COLUMN_NAME", ColumnName))
                Dim dtTable As New DataTable
                DaOut.Fill(dtTable)
                CloseOutDbCom()
                Return dtTable
        End Select
    End Function

    Public Function TablesSchema(ByVal TableName As String) As DataTable
        Dim dt As DataTable
        Select Case DbType
            Case "SQLITE"
                dt = GetSchema("select * from [" & TableName & "] LIMIT 1")
            Case Else
                dt = GetSchema("select top 1 * from [" & TableName & "]")
        End Select
        Return dt
    End Function
#End Region

#Region "DB"
    Public Function GetDbDv() As DataTable
        OpenDbCom()
        DbCom.CommandText = "select * from db"
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetDbInfo(ByVal intDbId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from db ")
        strSQL.Append("where DbId = @DbId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", intDbId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddDb(ByVal strName As String, ByVal strProject As String, ByVal strConnection As String, ByVal strConnName As String, ByVal strWSNameSpace As String, ByVal strDbType As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO DB ")
        strSQL.Append("(Name, Project, Conn, ConnName, WSNameSpace, DbType) ")
        strSQL.Append("VALUES(@Name, @Project, @Conn, @ConnName, @WSNameSpace, @DbType)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Project", strProject))
        DbCom.Parameters.Add(GetDBParameter("@Conn", strConnection))
        DbCom.Parameters.Add(GetDBParameter("@ConnName", strConnName))
        DbCom.Parameters.Add(GetDBParameter("@WSNameSpace", strWSNameSpace))
        DbCom.Parameters.Add(GetDBParameter("@DbType", strDbType))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateDb(ByVal intDbId As Integer, ByVal strName As String, ByVal strProject As String, ByVal strConnection As String, ByVal strConnName As String, ByVal strWSNameSpace As String, ByVal strDbType As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE DB SET ")
        strSQL.Append("Name = @Name")
        strSQL.Append(",Project = @Project")
        strSQL.Append(",Conn = @Conn")
        strSQL.Append(",ConnName = @ConnName ")
        strSQL.Append(",WSNameSpace = @WSNameSpace ")
        strSQL.Append(",DbType = @DbType ")
        strSQL.Append(" WHERE DbId = @DbId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Project", strProject))
        DbCom.Parameters.Add(GetDBParameter("@Conn", strConnection))
        DbCom.Parameters.Add(GetDBParameter("@ConnName", strConnName))
        DbCom.Parameters.Add(GetDBParameter("@WSNameSpace", strWSNameSpace))
        DbCom.Parameters.Add(GetDBParameter("@DbType", strDbType))
        DbCom.Parameters.Add(GetDBParameter("@DbId", intDbId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelDb(ByVal intDbId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM DB ")
        strSQL.Append("WHERE DbId = @DbId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", intDbId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "Table"
    Public Function GetTableDv() As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Table] ")
        strSQL.Append("where DbId = @DbId order by Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetTableInfo() As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Table] ")
        strSQL.Append("where TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetTableInfoByName(ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Table] ")
        strSQL.Append("where DbId = @DbId AND Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddTable(ByVal strName As String, ByVal strClassName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO [Table] ")
        strSQL.Append("(DbId, Name, ClassName, Summary) ")
        strSQL.Append("VALUES(@DbId, @Name, @ClassName, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@ClassName", strClassName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateTable(ByVal intTId As Integer, ByVal strClassName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE [Table] SET ")
        strSQL.Append("ClassName = @ClassName")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(" WHERE TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@ClassName", strClassName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@TId", intTId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelTable(ByVal intTId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM [Table] ")
        strSQL.Append("WHERE TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", intTId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "TableAddInfo"
    Public Function GetTableAddInfoDv() As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [TableAddInfo] ")
        strSQL.Append("where TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddTableAddInfo(ByVal strAddTable As String, ByVal strType As String, ByVal strName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO TableAddInfo ")
        strSQL.Append("(TId, AddTable, Type, Name, Summary) ")
        strSQL.Append("VALUES(@TId, @AddTable, @Type, @Name, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@AddTable", strAddTable))
        DbCom.Parameters.Add(GetDBParameter("@Type", strType))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateTableAddInfo(ByVal intAId As Integer, ByVal strName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE TableAddInfo SET ")
        strSQL.Append("Name = @Name ")
        strSQL.Append(",Summary = @Summary ")
        strSQL.Append(" WHERE AId = @AId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@AId", intAId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelTableAddInfo(ByVal intAId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM TableAddInfo ")
        strSQL.Append("WHERE AId = @AId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@AId", intAId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelTableAddInfoByTId(ByVal intTId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM TableAddInfo ")
        strSQL.Append("WHERE TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", intTId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelTableAddInfoByAddTable(ByVal AddTable As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM TableAddInfo ")
        strSQL.Append("WHERE AddTable = @AddTable")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@AddTable", AddTable))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "Column"
    Public Function GetColumnDv(Optional ByVal intTId As Integer = -1) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Column] ")
        strSQL.Append("where TId = @TId")
        DbCom.CommandText = strSQL.ToString
        If intTId = -1 Then
            intTId = TId
        End If
        DbCom.Parameters.Add(GetDBParameter("@TId", intTId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetColumnInfo(ByVal intColId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Column] ")
        strSQL.Append("where ColId = @ColId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@ColId", intColId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetColumnInfoByName(ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from [Column] ")
        strSQL.Append("where TId = @TId AND Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddColumn(ByVal strName As String,
                         ByVal strPropertyName As String,
                         ByVal strSummary As String,
                         ByVal strInput As String,
                         ByVal min As Object,
                         ByVal max As Object,
                         ByVal required As Boolean,
                         ByVal remark As String)
        OpenDbCom()

        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO [Column] ")
        strSQL.Append("(TId, Name, PropertyName, Summary, InputType, MinVal, MaxVal, Required, Remark) ")
        strSQL.Append("VALUES(@TId, @Name, @PropertyName, @Summary, @InputType, @MinVal, @MaxVal, @Required, Remark)")

        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@PropertyName", strPropertyName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@InputType", strInput))
        DbCom.Parameters.Add(GetDBParameter("@MinVal", min))
        DbCom.Parameters.Add(GetDBParameter("@MaxVal", max))
        DbCom.Parameters.Add(GetDBParameter("@Required", required))
        DbCom.Parameters.Add(GetDBParameter("@Remark", remark))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateColumn(ByVal intColId As Integer,
                            ByVal strPropertyName As String,
                            ByVal strSummary As String,
                            ByVal strInput As String,
                            ByVal min As Object,
                            ByVal max As Object,
                            ByVal required As Boolean,
                            ByVal remark As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE [Column] SET ")
        strSQL.Append("PropertyName = @PropertyName")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(",InputType = @InputType")
        strSQL.Append(",MinVal = @MinVal")
        strSQL.Append(",MaxVal = @MaxVal")
        strSQL.Append(",Required = @Required")
        strSQL.Append(",Remark = @Remark")
        strSQL.Append(" WHERE ColId = @ColId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@PropertyName", strPropertyName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@InputType", strInput))
        DbCom.Parameters.Add(GetDBParameter("@MinVal", min))
        DbCom.Parameters.Add(GetDBParameter("@MaxVal", max))
        DbCom.Parameters.Add(GetDBParameter("@Required", required))
        DbCom.Parameters.Add(GetDBParameter("@Remark", remark))
        DbCom.Parameters.Add(GetDBParameter("@ColId", intColId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateColumnSummary(ByVal intColId As Integer, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE [Column] SET ")
        strSQL.Append("Summary = @Summary")
        strSQL.Append(" WHERE ColId = @ColId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@ColId", intColId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelColumn(ByVal intColId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM [Column] ")
        strSQL.Append("WHERE ColId = @ColId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@ColId", intColId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "Fun"
    Public Function GetFunDv() As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from Fun ")
        strSQL.Append("where TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFunDv(ByVal intTId As Integer, ByVal strReturn As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from Fun ")
        strSQL.Append("where TId = @TId and ReturnType = @ReturnType")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", intTId))
        DbCom.Parameters.Add(GetDBParameter("@ReturnType", strReturn))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFunInfo(ByVal intFId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from Fun ")
        strSQL.Append("where FId = @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFunInfoByName(ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from Fun ")
        strSQL.Append("where TId = @TId and Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        dtTable.Clear()
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFunOtherByName(ByVal intFId As Integer, ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from Fun ")
        strSQL.Append("where TId = @TId and Name = @Name and FId <> @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFId() As Integer
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select max(FId) AS num from Fun ")
        strSQL.Append("where TId = @TId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable.DefaultView(0)(0)
    End Function

    Public Sub AddFun(ByVal strName As String, ByVal Sql As String, ByVal ReturnType As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO Fun ")
        strSQL.Append("(TId, Name, SqlText, ReturnType, Summary) ")
        strSQL.Append("VALUES(@TId, @Name, @SqlText, @ReturnType, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@TId", TId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@SqlText", Sql))
        DbCom.Parameters.Add(GetDBParameter("@ReturnType", ReturnType))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateFun(ByVal intFId As Integer, ByVal strName As String, ByVal Sql As String, ByVal ReturnType As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE Fun SET ")
        strSQL.Append("Name = @Name")
        strSQL.Append(",SqlText = @SqlText")
        strSQL.Append(",ReturnType = @ReturnType")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(" WHERE FId = @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@SqlText", Sql))
        DbCom.Parameters.Add(GetDBParameter("@ReturnType", ReturnType))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelFun(ByVal intFId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM Fun ")
        strSQL.Append("WHERE FId = @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "FunParam"
    Public Function GetFunParamDv(ByVal intFId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from FunParam ")
        strSQL.Append("where FId = @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetFunParamInfoByName(ByVal intFId As Integer, ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from FunParam ")
        strSQL.Append("where FId = @FId and Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddFunParam(ByVal intFId As Integer, ByVal strName As String, ByVal strType As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO FunParam ")
        strSQL.Append("(FId, Name, Type, Summary) ")
        strSQL.Append("VALUES(@FId, @Name, @Type, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Type", strType))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateFunParam(ByVal intMId As Integer, ByVal strName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE FunParam SET ")
        strSQL.Append("Name = @Name")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(" WHERE MId = @MId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@MId", intMId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelFunParam(ByVal intMId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM FunParam ")
        strSQL.Append("WHERE MId = @MId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@MId", intMId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelFunParamByFId(ByVal intFId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM FunParam ")
        strSQL.Append("WHERE FId = @FId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@FId", intFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "CrossFun"
    Public Function GetCrossFunDv() As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFun ")
        strSQL.Append("where DbId = @DbId order by Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetCrossFunInfo(ByVal intCFId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFun ")
        strSQL.Append("where cFId = @cFId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetCrossFunInfoByName(ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFun ")
        strSQL.Append("where DbId = @DbId and Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetCrossOtherByName(ByVal intCFId As Integer, ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFun ")
        strSQL.Append("where DbId = @DbId and Name = @Name and cFId <> @cFId order by Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetCFId() As Integer
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select max(cFId) AS num from cFun ")
        strSQL.Append("where DbId = @DbId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable.DefaultView(0)(0)
    End Function

    Public Sub AddCrossFun(ByVal strName As String, ByVal Sql As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO cFun ")
        strSQL.Append("(DbId, Name, SqlText, Summary) ")
        strSQL.Append("VALUES(@DbId, @Name, @SqlText, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@DbId", DbId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@SqlText", Sql))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateCrossFun(ByVal intCFId As Integer, ByVal strName As String, ByVal Sql As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE cFun SET ")
        strSQL.Append("Name = @Name")
        strSQL.Append(",SqlText = @SqlText")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(" WHERE cFId = @cFId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@SqlText", Sql))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelCrossFun(ByVal intCFId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM cFun ")
        strSQL.Append("WHERE cFId = @cFId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region

#Region "CrossParam"
    Public Function GetCrossParamDv(ByVal intCFId As Integer) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFunParam ")
        strSQL.Append("where cFId = @cFId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Function GetCrossParamInfoByName(ByVal intCFId As Integer, ByVal strName As String) As DataTable
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("select * from cFunParam ")
        strSQL.Append("where cFId = @cFId and Name = @Name")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        Dim dtTable As New DataTable
        Da.Fill(dtTable)
        CloseDbCom()
        Return dtTable
    End Function

    Public Sub AddCrossParam(ByVal intCFId As Integer, ByVal strName As String, ByVal strType As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("INSERT INTO cFunParam ")
        strSQL.Append("(cFId, Name, Type, Summary) ")
        strSQL.Append("VALUES(@cFId, @Name, @Type, @Summary)")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Type", strType))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub UpDateCrossParam(ByVal intCMId As Integer, ByVal strName As String, ByVal strSummary As String)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append("UPDATE cFunParam SET ")
        strSQL.Append("Name = @Name")
        strSQL.Append(",Summary = @Summary")
        strSQL.Append(" WHERE cMId = @cMId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@Name", strName))
        DbCom.Parameters.Add(GetDBParameter("@Summary", strSummary))
        DbCom.Parameters.Add(GetDBParameter("@cMId", intCMId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelCrossParam(ByVal intCMId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM cFunParam ")
        strSQL.Append("WHERE cMId = @cMId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@CMId", intCMId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub

    Public Sub DelCrossParamByCFId(ByVal intCFId As Integer)
        OpenDbCom()
        Dim strSQL As New System.Text.StringBuilder
        strSQL.Append(GetDel() & "FROM cFunParam ")
        strSQL.Append("WHERE cFId = @cFId")
        DbCom.CommandText = strSQL.ToString
        DbCom.Parameters.Add(GetDBParameter("@cFId", intCFId))
        DbCom.ExecuteNonQuery()
        CloseDbCom()
    End Sub
#End Region
End Module
