Imports System.Data
Imports System.Text
Imports System.IO
Module mdFile

#Region "Propertys"
    Dim i, j As Integer
    Public strTable As String
    Public strClass As String
    Dim dt As DataTable
    Dim mArrayKEY As Array
    Dim mArrayNOKEY As Array
    Dim mInputKEY As String
    Public HasAutoIncrement As String
    Public HasMvcInputControl As Boolean
    Public ReadOnly Property ArrayKEY() As Array
        Get
            Return mArrayKEY
        End Get
    End Property
    Public ReadOnly Property strKEY() As String
        Get
            Dim str As String = ""
            For i = 0 To mArrayKEY.Length - 1
                If i > 0 Then
                    str &= ", "
                End If
                str &= ColumnProperty(mArrayKEY(i))
            Next
            Return str
        End Get
    End Property
    Private ReadOnly Property strCrossKEY() As String
        Get
            Dim str As String = ""
            For i = 0 To mArrayKEY.Length - 1
                If i > 0 Then
                    str &= ", "
                End If
                str &= mArrayKEY(i)
            Next
            Return str
        End Get
    End Property
    Private ReadOnly Property ArrayNOKEY() As Array
        Get
            Return mArrayNOKEY
        End Get
    End Property
    Private ReadOnly Property InputKEY() As String
        Get
            Return mInputKEY
        End Get
    End Property

    Public ReadOnly Property classDT() As DataTable
        Get
            Return dt
        End Get
    End Property
#End Region

    Public Sub Table_Initial()
        strTable = TableName()
        strClass = TableClass()

        dt = TablesSchema(TableName)

        HasAutoIncrement = ""
        Dim boolPK As Boolean = False
        Dim thisArrayKEY As New ArrayList
        Dim thisArrayNOKEY As New ArrayList

        For i = 0 To dt.Columns.Count - 1
            For j = 0 To dt.PrimaryKey.GetUpperBound(0)
                If dt.Columns(i).ColumnName = dt.PrimaryKey(j).ColumnName Then
                    boolPK = True
                End If
            Next
            If dt.Columns(i).AutoIncrement = True Then
                boolPK = True
                HasAutoIncrement = dt.Columns(i).ColumnName
            End If
            If boolPK = True Then
                thisArrayKEY.Add(dt.Columns(i).ColumnName)
            Else
                thisArrayNOKEY.Add(dt.Columns(i).ColumnName)
            End If
            boolPK = False
        Next

        mArrayKEY = thisArrayKEY.ToArray
        mArrayNOKEY = thisArrayNOKEY.ToArray

        Dim strSET As New StringBuilder
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(", ")
            End If
            strSET.Append(ByValAs(ColumnProperty(ArrayKEY(i)), dt.Columns(ArrayKEY(i)).DataType.Name))
        Next

        mInputKEY = strSET.ToString
    End Sub

#Region "Table"
    Public Sub GetTableCode()

        Table_Initial()

        SetInfo()
        SetDB()
        SetBiz()
        SetBizAdv()

        If WSNameSpace <> "" Then
            SetWS()
        End If

        'MVC
        SetModel()
        SetController()
        SetViewIndex()
        SetViewDetails()
        SetViewEdit()

        dt.Clear()

    End Sub
    Private Sub SetInfo()

        strNameSpace = "Information"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.ComponentModel"))
        strSET.Append(SetImports("System.ComponentModel.DataAnnotations"))

        strSET.Append(PartialClass(strClass & "Info", TableSummary))

        strSET.Append(GetLine())

        For i = 0 To dt.Columns.Count - 1


            If i > 0 Then
                strSET.Append(GetLine())
            End If
            strSET.Append(SetFunSummary(ColumnSummary(dt.Columns(i).ColumnName)))
            strSET.Append(SetFunRemarks(ColumnRemarks(dt.Columns(i).ColumnName)))

            Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
            If ColumnInfo.Count > 0 Then
                Dim inputType As String = Trim(ColumnInfo(0)("InputType"))

                If inputType <> "" Then

                    If Trim(ColumnInfo(0)("Summary")) <> "" Then
                        strSET.Append("        [DisplayName(""" & ColumnInfo(0)("Summary") & """)]" & vbCrLf)
                    End If

                    If ColumnInfo(0)("Required") Then
                        strSET.Append("        [Required(ErrorMessage = ""請輸入" & ColumnInfo(0)("Summary") & """)]" & vbCrLf)
                    End If

                    'dt.Columns(i).DataType.Name

                    'InputType
                    Select Case inputType
                        Case "Password"
                            strSET.Append("        [DataType(DataType.Password)]" & vbCrLf)
                            strSET.Append("        [RegularExpression(@""[a-zA-Z]+[a-zA-Z0-9]*$"", ErrorMessage = """ & ColumnInfo(0)("Summary") & "僅能有英文或數字，且開頭需為英文字母"")]" & vbCrLf)
                            strSET.Append(GetPropertyRange(ColumnInfo(0), dt.Columns(i).DataType.Name))
                        Case "Date"
                            strSET.Append("        [DataType(DataType.Date)]" & vbCrLf)
                            strSET.Append("        [DisplayFormat(DataFormatString = ""{0:yyyy-MM-dd}"")]" & vbCrLf)
                        Case "Email"
                            strSET.Append("        [EmailAddress(ErrorMessage = """ & ColumnInfo(0)("Summary") & "請輸入電子郵件格式"")]" & vbCrLf)
                            strSET.Append(GetPropertyRange(ColumnInfo(0), dt.Columns(i).DataType.Name))
                        Case "Text"
                            strSET.Append(GetPropertyRange(ColumnInfo(0), dt.Columns(i).DataType.Name))
                        Case "DropDownList", "RadioButtonList", "File"
                            HasMvcInputControl = True
                            'isSelect = True

                            'Case "Hidden"
                            'Case "CheckBox"
                    End Select
                End If
            End If
            strSET.Append(SetProperty(ColumnProperty(dt.Columns(i).ColumnName), dt.Columns(i).DataType.Name))
        Next

        strSET.Append(EndClass)

        If Directory.Exists("Information") = False Then
            Directory.CreateDirectory("Information")
        End If

        If File.Exists("Information\" & strClass & "Info." & CodeType) = True Then
            File.Delete("Information\" & strClass & "Info." & CodeType)
        End If

        File.WriteAllText("Information\" & strClass & "Info." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetDB()

        strNameSpace = "DataAccess"

        Dim strSQL As String
        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("Microsoft.Practices.EnterpriseLibrary.Data"))
        strSET.Append(SetImports("Microsoft.Practices.EnterpriseLibrary.Data.Sql"))
        strSET.Append(SetImports("System.Text"))
        strSET.Append(SetImports("System.Data"))
        strSET.Append(SetImports("System.Data.Common"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.Collections"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("Commons"))
        strSET.Append(PartialClass(strClass & "DB", TableSummary)) 'strSET.Append(PublicClass(strClass & "DB", TableSummary))

        '取得此Table類別的Function
        strSET.Append(GetDBFun)

        'Exists Info
        strSET.Append(SetFunSummary("判斷" & strClass & "此筆資料是否存在"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicFun("Exists", "Boolean", InputKEY))
        strSET.Append(GetDB())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))
        strSQL = "SELECT count(1) FROM " & strTable & " WHERE "
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSQL &= " AND "
            End If
            strSQL &= ArrayKEY(i) & " = @" & ArrayKEY(i)
        Next
        strSET.Append(SqlAppend(strSQL))
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        For i = 0 To ArrayKEY.Length - 1
            strSET.Append(DbAddInParameter(ArrayKEY(i), dt.Columns(ArrayKEY(i)).DataType.Name, ColumnProperty(ArrayKEY(i))))
        Next
        strSET.Append(ExistsResult)
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'Read All
        strSET.Append(SetFunSummary("讀取全部資料"))
        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
        strSET.Append(PublicFun("GetAll", GetIlist(strClass), ""))
        strSET.Append(GetDB())
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))
        strSET.Append(SqlAppend("SELECT * FROM " & strTable))
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        strSET.Append(ReturnValue("GetList(db, dbCommand)"))
        strSET.Append(GetLine())
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'Read Info
        strSET.Append(SetFunSummary("讀取" & strClass & "1筆資料"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary(strClass & "Info"))
        strSET.Append(PublicFun("GetInfo", strClass & "Info", InputKEY))
        strSET.Append(GetDB())
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))

        strSQL = "SELECT * FROM " & strTable & " WHERE "
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSQL &= " AND "
            End If
            strSQL &= ArrayKEY(i) & " = @" & ArrayKEY(i)
        Next
        strSET.Append(SqlAppend(strSQL))

        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        For i = 0 To ArrayKEY.Length - 1
            strSET.Append(DbAddInParameter(ArrayKEY(i), dt.Columns(ArrayKEY(i)).DataType.Name, ColumnProperty(ArrayKEY(i))))
        Next
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("myInfo", strClass & "Info"))
        strSET.Append(DimAsNewList("myList", strClass))
        strSET.Append(SetString("myList = GetList(db, dbCommand)", "        "))
        strSET.Append(ChkDBInfo())
        strSET.Append(GetLine)
        strSET.Append(ReturnValue("myInfo"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'AddNew
        strSET.Append(SetFunSummary("新增" & strClass & "1筆資料"))
        strSET.Append(SetParamSummary(strClass, strClass & "新資料"))
        strSET.Append(SetReturnSummary("Boolean"))

        If HasAutoIncrement <> "" Then
            strSET.Append(PublicFun("AddNew", "Boolean", ByRefAs(strClass, strClass & "Info")))
        Else
            strSET.Append(PublicFun("AddNew", "Boolean", ByValAs(strClass, strClass & "Info")))
        End If

        strSET.Append(GetDB)
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))
        strSET.Append(SqlAppend("INSERT INTO " & strTable & " "))

        Dim strFIELD1 As String = ""
        Dim strFIELD2 As String = ""
        For i = 0 To dt.Columns.Count - 1
            If dt.Columns(i).AutoIncrement = False Then
                If strFIELD1 <> "" Then
                    strFIELD1 &= ","
                    strFIELD2 &= ","
                End If
                strFIELD1 &= dt.Columns(i).ColumnName
                strFIELD2 &= "@" & dt.Columns(i).ColumnName
            End If
        Next

        strSET.Append(SqlAppend("(" & strFIELD1 & ")"))
        strSET.Append(SqlAppend("VALUES(" & strFIELD2 & ")"))

        If HasAutoIncrement <> "" Then
            Select Case DbType
                Case "SQLITE"

                Case Else
                    strSET.Append(SqlAppend(" SELECT @" & HasAutoIncrement & " = SCOPE_IDENTITY()"))
            End Select
        End If

        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))

        If HasAutoIncrement <> "" Then
            Select Case DbType
                Case "SQLITE"

                Case Else
                    strSET.Append(DbAddOutParameter(HasAutoIncrement))
            End Select
        End If

        For i = 0 To dt.Columns.Count - 1
            If dt.Columns(i).AutoIncrement = False Then
                strSET.Append(DbAddInParameter(dt.Columns(i).ColumnName, dt.Columns(i).DataType.Name, strClass & "." & ColumnProperty(dt.Columns(i).ColumnName)))
            End If
        Next

        strSET.Append(GetLine())
        strSET.Append(DimAsValue("result", "Boolean", "False"))
        strSET.Append(GetTry)
        strSET.Append(SetString("db.ExecuteNonQuery(dbCommand)", "            "))

        If HasAutoIncrement <> "" Then
            Select Case DbType
                Case "SQLITE"
                    strSET.Append(SetString("dbCommand = db.GetSqlStringCommand(""SELECT max(" & HasAutoIncrement & ") FROM " & strTable & " "")", "            "))
                    strSET.Append(DbGetAutoIncrementKey(strClass & "." & ColumnProperty(HasAutoIncrement)))
                Case Else
                    strSET.Append(DbGetParameterValue(strClass & "." & ColumnProperty(HasAutoIncrement), HasAutoIncrement))
            End Select
        End If

        strSET.Append(SetString("result = " & ConvertValue("True"), "            "))
        strSET.Append(SetCatch(strClass & "DB.AddNew"))
        strSET.Append(ReturnValue("result"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'Update
        strSET.Append(SetFunSummary("更新" & strClass & "某筆資料"))
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicFun("Update", "Boolean", ByValAs(strClass, strClass & "Info")))
        strSET.Append(GetDB)
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))
        strSET.Append(SqlAppend("UPDATE " & strTable & " SET "))
        For i = 0 To ArrayNOKEY.Length - 1
            If i > 0 Then
                strSET.Append(SqlAppend("," & ArrayNOKEY(i) & " = @" & ArrayNOKEY(i)))
            Else
                strSET.Append(SqlAppend(ArrayNOKEY(i) & " = @" & ArrayNOKEY(i)))
            End If
        Next
        strSQL = " WHERE "
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSQL &= " AND "
            End If
            strSQL &= ArrayKEY(i) & " = @" & ArrayKEY(i)
        Next
        strSET.Append(SqlAppend(strSQL))
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        For i = 0 To dt.Columns.Count - 1
            strSET.Append(DbAddInParameter(dt.Columns(i).ColumnName, dt.Columns(i).DataType.Name, strClass & "." & ColumnProperty(dt.Columns(i).ColumnName)))
        Next
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("result", "Boolean", "False"))
        strSET.Append(GetTry)
        strSET.Append(SetString("db.ExecuteNonQuery(dbCommand)", "            "))
        strSET.Append(SetString("result = " & ConvertValue("True"), "            "))
        strSET.Append(SetCatch(strClass & "DB.Update"))
        strSET.Append(ReturnValue("result"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'Del
        strSET.Append(SetFunSummary("刪除" & strClass & "某筆資料"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicFun("Del", "Boolean", InputKEY))
        strSET.Append(GetDB)
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))

        Select Case DbType
            Case "SQLITE"
                strSQL = "DELETE FROM " & strTable & " WHERE "
            Case Else
                strSQL = "DELETE " & strTable & " WHERE "
        End Select

        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSQL &= " AND "
            End If
            strSQL &= ArrayKEY(i) & " = @" & ArrayKEY(i)
        Next
        strSET.Append(SqlAppend(strSQL))
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        For i = 0 To ArrayKEY.Length - 1
            strSET.Append(DbAddInParameter(ArrayKEY(i), dt.Columns(ArrayKEY(i)).DataType.Name, ColumnProperty(ArrayKEY(i))))
        Next
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("result", "Boolean", "False"))
        strSET.Append(GetTry)
        strSET.Append(SetString("db.ExecuteNonQuery(dbCommand)", "            "))
        strSET.Append(SetString("result = " & ConvertValue("True"), "            "))
        strSET.Append(SetCatch(strClass & "DB.Del"))
        strSET.Append(ReturnValue("result"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'GetList
        strSET.Append(PrivateFun("GetList", GetIlist(strClass), ByValAs("db", "Database") & ", " & ByValAs("dbCommand", "DbCommand")))
        strSET.Append(DimAsNewList("myList", strClass))
        strSET.Append(GetLine())
        strSET.Append(GetTry)
        strSET.Append(UsingValue("dataReader", "IDataReader", "db.ExecuteReader(dbCommand)"))
        strSET.Append(WhileValue("dataReader.Read()"))
        strSET.Append(SetString("myList.Add(BindInfo(dataReader))", "                    "))
        strSET.Append(EndWhile)
        strSET.Append(EndUsing)
        strSET.Append(SetCatch(strClass & "DB.GetList"))
        strSET.Append(ReturnValue("myList"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'BindInfo
        strSET.Append(PrivateFun("BindInfo", strClass & "Info", ByValAs("dr", "IDataReader")))
        strSET.Append(DimAsNew("myInfo", strClass & "Info"))
        strSET.Append(GetLine())
        strSET.Append(GetTry)
        For i = 0 To dt.Columns.Count - 1
            strSET.Append(SetInfoColumns(dt.Columns(i).ColumnName, dt.Columns(i).DataType.Name))
        Next
        strSET.Append(SetCatch(strClass & "DB.BindInfo", "Exception"))
        strSET.Append(ReturnValue("myInfo"))
        strSET.Append(EndFun)
        strSET.Append(EndClass)

        If Directory.Exists("DataAccess") = False Then
            Directory.CreateDirectory("DataAccess")
        End If

        If File.Exists("DataAccess\" & strClass & "DB." & CodeType) = True Then
            File.Delete("DataAccess\" & strClass & "DB." & CodeType)
        End If

        File.WriteAllText("DataAccess\" & strClass & "DB." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetBiz()

        strNameSpace = "Business"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("DataAccess"))
        strSET.Append(PartialClass(strClass & "Biz", TableSummary))

        strSET.Append(PrivateSharedAsNewType("myDB", strClass & "DB"))
        strSET.Append(GetLine())

        '取得此Table類別的Function
        strSET.Append(GetBizFun())

        strSET.Append(SetFunSummary("判斷" & strClass & "此筆資料是否存在"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicSharedFun("Exists", "Boolean", InputKEY))
        strSET.Append(ReturnValue("myDB.Exists(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("讀取全部資料"))
        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
        strSET.Append(PublicSharedFun("GetAll", GetIlist(strClass), ""))
        strSET.Append(ReturnValue("myDB.GetAll()"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("讀取" & strClass & "1筆資料"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary(strClass & "Info"))
        strSET.Append(PublicSharedFun("GetInfo", strClass & "Info", InputKEY))
        'strSET.Append(DimAsNew("myDB", strClass & "DB"))
        'strSET.Append(GetLine())
        strSET.Append(ReturnValue("myDB.GetInfo(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("新增" & strClass & "1筆資料"))
        strSET.Append(SetParamSummary(strClass, strClass & "的新資料"))
        strSET.Append(SetReturnSummary("Boolean"))

        If HasAutoIncrement <> "" Then
            strSET.Append(PublicSharedFun("AddNew", "Boolean", ByRefAs(strClass, strClass & "Info")))
            strSET.Append(ReturnValue("myDB.AddNew(" & GetRefParam(strClass) & ")"))
        Else
            strSET.Append(PublicSharedFun("AddNew", "Boolean", ByValAs(strClass, strClass & "Info")))
            strSET.Append(ReturnValue("myDB.AddNew(" & strClass & ")"))
            'strSET.Append(DimAsNew("myDB", strClass & "DB"))
            'strSET.Append(GetLine())
        End If

        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("更新" & strClass & "某筆資料"))
        strSET.Append(SetParamSummary(strClass, strClass & "的某筆資料"))
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicSharedFun("Update", "Boolean", ByValAs(strClass, strClass & "Info")))
        'strSET.Append(DimAsNew("myDB", strClass & "DB"))
        'strSET.Append(GetLine())
        strSET.Append(ReturnValue("myDB.Update(" & strClass & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("刪除" & strClass & "某筆資料"))
        For i = 0 To mArrayKEY.Length - 1
            Dim strColsmy As String = ColumnSummary(mArrayKEY(i))
            If strColsmy <> "" Then
                strSET.Append(SetParamSummary(ColumnProperty(mArrayKEY(i)), strColsmy))
            End If
        Next
        strSET.Append(SetReturnSummary("Boolean"))
        strSET.Append(PublicSharedFun("Del", "Boolean", InputKEY))
        'strSET.Append(DimAsNew("myDB", strClass & "DB"))
        'strSET.Append(GetLine())
        strSET.Append(ReturnValue("myDB.Del(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(EndClass)

        If Directory.Exists("Business") = False Then
            Directory.CreateDirectory("Business")
        End If

        If File.Exists("Business\" & strClass & "Biz_Bas." & CodeType) = True Then
            File.Delete("Business\" & strClass & "Biz_Bas." & CodeType)
        End If

        File.WriteAllText("Business\" & strClass & "Biz_Bas." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetBizAdv()

        strNameSpace = "Business"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("DataAccess"))
        strSET.Append(PartialClass(strClass & "Biz"))
        strSET.Append(EndClass)

        If Directory.Exists("Business") = False Then
            Directory.CreateDirectory("Business")
        End If

        If File.Exists("Business\" & strClass & "Biz_Adv." & CodeType) = True Then
            File.Delete("Business\" & strClass & "Biz_Adv." & CodeType)
        End If

        File.WriteAllText("Business\" & strClass & "Biz_Adv." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetWS()

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Web"))
        strSET.Append(SetImports("System.Web.Services"))
        strSET.Append(SetImports("System.Web.Services.Protocols"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.Collections"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("Business"))
        strSET.Append(WSClass(strClass & "WS", TableSummary))

        '取得此Table類別的Function
        strSET.Append(GetWSFun())

        strSET.Append(SetFunSummary("判斷" & strClass & "此筆資料是否存在"))
        strSET.Append(GetWebMethod("判斷" & strClass & "此筆資料是否存在"))
        strSET.Append(PublicFun("Exists", "Boolean", InputKEY))
        strSET.Append(ReturnValue(strClass & "Biz.Exists(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("讀取全部資料"))
        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
        strSET.Append(GetWebMethod("讀取全部資料"))
        strSET.Append(PublicFun("GetAll", GetList(strClass), ""))
        strSET.Append(ReturnValue(GetCType(strClass & "Biz.GetAll()", GetList(strClass))))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("讀取" & strClass & "1筆資料"))
        strSET.Append(GetWebMethod("讀取" & strClass & "1筆資料"))
        strSET.Append(PublicFun("GetInfo", strClass & "Info", InputKEY))
        strSET.Append(ReturnValue(strClass & "Biz.GetInfo(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("新增" & strClass & "1筆資料"))
        strSET.Append(GetWebMethod("新增" & strClass & "1筆資料"))

        If HasAutoIncrement <> "" Then
            strSET.Append(PublicFun("AddNew", "Boolean", ByRefAs(strClass, strClass & "Info")))
            strSET.Append(ReturnValue(strClass & "Biz.AddNew(" & GetRefParam(strClass) & ")"))
        Else
            strSET.Append(PublicFun("AddNew", "Boolean", ByValAs(strClass, strClass & "Info")))
            strSET.Append(ReturnValue(strClass & "Biz.AddNew(" & strClass & ")"))
        End If

        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("更新" & strClass & "某筆資料"))
        strSET.Append(GetWebMethod("更新" & strClass & "某筆資料"))
        strSET.Append(PublicFun("Update", "Boolean", ByValAs(strClass, strClass & "Info")))
        strSET.Append(ReturnValue(strClass & "Biz.Update(" & strClass & ")"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        strSET.Append(SetFunSummary("刪除" & strClass & "某筆資料"))
        strSET.Append(GetWebMethod("刪除" & strClass & "某筆資料"))
        strSET.Append(PublicFun("Del", "Boolean", InputKEY))
        strSET.Append(ReturnValue(strClass & "Biz.Del(" & strKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(EndWSClass)

        If Directory.Exists("WebService") = False Then
            Directory.CreateDirectory("WebService")
        End If

        If Directory.Exists("WebService\App_Code") = False Then
            Directory.CreateDirectory("WebService\App_Code")
        End If

        If File.Exists("WebService\" & strClass & "WS.asmx") = True Then
            File.Delete("WebService\" & strClass & "WS.asmx")
        End If
        Dim str As String
        If CodeType = "cs" Then
            str = "<%@ WebService Language=""C#"" CodeBehind=""~/App_Code/" & strClass & "WS.cs"" Class=""" & strClass & "WS"" %>"
        Else
            str = "<%@ WebService Language=""VB"" CodeBehind=""~/App_Code/" & strClass & "WS.vb"" Class=""" & strClass & "WS"" %>"
        End If
        File.WriteAllText("WebService\" & strClass & "WS.asmx", str, Encoding.BigEndianUnicode)

        If File.Exists("WebService\App_Code\" & strClass & "WS." & CodeType) = True Then
            File.Delete("WebService\App_Code\" & strClass & "WS." & CodeType)
        End If
        File.WriteAllText("WebService\App_Code\" & strClass & "WS." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Function GetDBFun() As String
        Dim dvFun As DataView = GetFunDv.DefaultView
        If dvFun.Count > 0 Then
            Dim strSET As New StringBuilder
            Dim strByVal, strParameter, strParamSummary As String
            Dim i, j, k As Integer

            For i = 0 To dvFun.Count - 1

                Dim dvParam As DataView = GetFunParamDv(dvFun(i)("FId")).DefaultView
                strByVal = ""
                strParameter = ""
                strParamSummary = ""
                If dvParam.Count > 0 Then
                    For j = 0 To dvParam.Count - 1

                        If j > 0 Then
                            strByVal &= ", "
                        End If
                        strByVal &= ByValAs(dvParam(j)("Name"), dvParam(j)("Type"))

                        strParameter &= DbAddInParameter(dvParam(j)("Name"), dvParam(j)("Type"), dvParam(j)("Name"))

                        strParamSummary &= SetParamSummary(dvParam(j)("Name"), dvParam(j)("Summary"))

                    Next
                End If

                strSET.Append(SetFunSummary(dvFun(i)("Summary")))
                strSET.Append(strParamSummary)
                Select Case dvFun(i)("ReturnType")
                    Case "Info"
                        strSET.Append(SetReturnSummary(strClass & "Info"))
                        strSET.Append(PublicFun(dvFun(i)("Name"), strClass & "Info", strByVal))
                    Case "List"
                        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
                        strSET.Append(PublicFun(dvFun(i)("Name"), GetIlist(strClass), strByVal))
                    Case Else
                        strSET.Append(SetReturnSummary("Boolean"))
                        strSET.Append(PublicFun(dvFun(i)("Name"), "Boolean", strByVal))
                End Select
                strSET.Append(GetDB)
                strSET.Append(GetLine())
                strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))

                Dim arySQL As Array = dvFun(i)("SqlText").Split(GetLine())
                For k = 0 To arySQL.Length - 1
                    strSET.Append(SqlAppend(Trim(arySQL(k)).ToString.Replace(Chr(10), "") & " "))
                Next

                strSET.Append(GetLine())
                strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))

                strSET.Append(strParameter)

                strSET.Append(GetLine())
                Select Case dvFun(i)("ReturnType")
                    Case "Info"
                        strSET.Append(DimAsNew("myInfo", strClass & "Info"))
                        strSET.Append(DimAsNewList("myList", strClass))
                        strSET.Append(SetString("myList = GetList(db, dbCommand)", "        "))
                        strSET.Append(ChkDBInfo())
                        strSET.Append(GetLine)
                        strSET.Append(ReturnValue("myInfo"))
                    Case "List"
                        strSET.Append(ReturnValue("GetList(db, dbCommand)"))
                    Case Else
                        strSET.Append(DimAsValue("result", "Boolean", "False"))
                        strSET.Append(GetTry)
                        strSET.Append(SetString("db.ExecuteNonQuery(dbCommand)", "            "))
                        strSET.Append(SetString("result = " & ConvertValue("True"), "            "))
                        strSET.Append(SetCatch(strClass & "DB." & dvFun(i)("Name")))
                        strSET.Append(ReturnValue("result"))
                End Select
                strSET.Append(EndFun)
                strSET.Append(GetLine())
            Next
            Return strSET.ToString
        Else
            Return ""
        End If

    End Function
    Private Function GetBizFun() As String

        Dim dvFun As DataView = GetFunDv.DefaultView
        If dvFun.Count > 0 Then
            Dim strSET As New StringBuilder
            Dim strByVal, strInVal, strParamSummary As String
            Dim i, j As Integer

            For i = 0 To dvFun.Count - 1

                Dim dvParam As DataView = GetFunParamDv(dvFun(i)("FId")).DefaultView
                strByVal = ""
                strInVal = ""
                strParamSummary = ""
                If dvParam.Count > 0 Then
                    For j = 0 To dvParam.Count - 1
                        If j > 0 Then
                            strByVal &= ", "
                            strInVal &= ", "
                        End If
                        strByVal &= ByValAs(dvParam(j)("Name"), dvParam(j)("Type"))
                        strInVal &= dvParam(j)("Name")
                        strParamSummary &= SetParamSummary(dvParam(j)("Name"), dvParam(j)("Summary"))
                    Next
                End If

                strSET.Append(SetFunSummary(dvFun(i)("Summary")))
                strSET.Append(strParamSummary)

                Select Case dvFun(i)("ReturnType")
                    Case "Info"
                        strSET.Append(SetReturnSummary(strClass & "Info"))
                        strSET.Append(PublicSharedFun(dvFun(i)("Name"), strClass & "Info", strByVal))
                    Case "List"
                        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
                        strSET.Append(PublicSharedFun(dvFun(i)("Name"), GetIlist(strClass), strByVal))
                    Case Else
                        strSET.Append(SetReturnSummary("Boolean"))
                        strSET.Append(PublicSharedFun(dvFun(i)("Name"), "Boolean", strByVal))
                End Select
                strSET.Append(ReturnValue("myDB." & dvFun(i)("Name") & "(" & strInVal & ")"))
                strSET.Append(EndFun)

                strSET.Append(GetLine())
            Next
            Return strSET.ToString
        Else
            Return ""
        End If

    End Function
    Private Function GetWSFun() As String

        Dim dvFun As DataView = GetFunDv.DefaultView
        If dvFun.Count > 0 Then
            Dim strSET As New StringBuilder
            Dim strByVal, strInVal, strParamSummary As String
            Dim i, j As Integer

            For i = 0 To dvFun.Count - 1

                Dim dvParam As DataView = GetFunParamDv(dvFun(i)("FId")).DefaultView
                strByVal = ""
                strInVal = ""
                strParamSummary = ""
                If dvParam.Count > 0 Then
                    For j = 0 To dvParam.Count - 1
                        If j > 0 Then
                            strByVal &= ", "
                            strInVal &= ", "
                        End If
                        strByVal &= ByValAs(dvParam(j)("Name"), dvParam(j)("Type"))
                        strInVal &= dvParam(j)("Name")
                        strParamSummary &= SetParamSummary(dvParam(j)("Name"), dvParam(j)("Summary"))
                    Next
                End If

                strSET.Append(SetFunSummary(dvFun(i)("Summary")))
                strSET.Append(strParamSummary)
                Dim strReturn As String = strClass & "Biz." & dvFun(i)("Name") & "(" & strInVal & ")"
                Select Case dvFun(i)("ReturnType")
                    Case "Info"
                        strSET.Append(SetReturnSummary(strClass & "Info"))
                        strSET.Append(GetWebMethod(dvFun(i)("Summary")))
                        strSET.Append(PublicFun(dvFun(i)("Name"), strClass & "Info", strByVal))
                    Case "List"
                        strSET.Append(SetReturnSummary(strClass & "Info的物件集合"))
                        strSET.Append(GetWebMethod(dvFun(i)("Summary")))
                        strSET.Append(PublicFun(dvFun(i)("Name"), GetList(strClass), strByVal))

                        strReturn = GetCType(strReturn, GetList(strClass))

                    Case Else
                        strSET.Append(SetReturnSummary("Boolean"))
                        strSET.Append(GetWebMethod(dvFun(i)("Summary")))
                        strSET.Append(PublicFun(dvFun(i)("Name"), "Boolean", strByVal))
                End Select
                strSET.Append(ReturnValue(strReturn))
                strSET.Append(EndFun)

                strSET.Append(GetLine())
            Next
            Return strSET.ToString
        Else
            Return ""
        End If

    End Function
#End Region

#Region "Cross"
    Dim CrossInfo As DataView
    Dim CrossParam As DataView
    Public Sub GetCrossCode(ByVal intCFId As Integer)

        CrossInfo = GetCrossFunInfo(intCFId).DefaultView
        CrossParam = GetCrossParamDv(intCFId).DefaultView

        Dim strSQL As String = CrossInfo(0)("SqlText")
        Dim strSET As New StringBuilder
        Dim thisArrayList As New ArrayList

        If CrossParam.Count > 0 Then
            For i = 0 To CrossParam.Count - 1

                strSQL = ReplaceSQL(strSQL, CrossParam(i)("Name").ToString, CrossParam(i)("Type"))

                If i > 0 Then
                    strSET.Append(", ")
                End If
                strSET.Append(ByValAs(CrossParam(i)("Name"), CrossParam(i)("Type")))

                thisArrayList.Add(CrossParam(i)("Name"))

            Next
        End If

        mInputKEY = strSET.ToString
        mArrayKEY = thisArrayList.ToArray

        strClass = "Cross" & CrossInfo(0)("Name")

        dt = GetSchema(strSQL)

        SetCrossInfo()
        SetCrossDB()
        SetCrossBiz()

        If WSNameSpace <> "" Then
            SetCrossWS()
        End If

        dt.Clear()

    End Sub
    Private Sub SetCrossInfo()

        strNameSpace = "Information"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(PublicClass(strClass & "Info", CrossInfo(0)("Summary")))
        For i = 0 To dt.Columns.Count - 1
            strSET.Append(PrivateAsType("m_" & GetObjectName(dt.Columns(i).ColumnName), dt.Columns(i).DataType.Name))
        Next
        strSET.Append(mdCode.GetLine())
        For i = 0 To dt.Columns.Count - 1
            If i > 0 Then
                strSET.Append(GetLine())
            End If
            strSET.Append(SetProperty(GetObjectName(dt.Columns(i).ColumnName), dt.Columns(i).DataType.Name))
        Next
        strSET.Append(EndClass)

        If Directory.Exists("Information") = False Then
            Directory.CreateDirectory("Information")
        End If

        If File.Exists("Information\" & strClass & "Info." & CodeType) = True Then
            File.Delete("Information\" & strClass & "Info." & CodeType)
        End If

        File.WriteAllText("Information\" & strClass & "Info." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetCrossDB()

        strNameSpace = "DataAccess"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("Microsoft.Practices.EnterpriseLibrary.Data"))
        strSET.Append(SetImports("Microsoft.Practices.EnterpriseLibrary.Data.Sql"))
        strSET.Append(SetImports("System.Text"))
        strSET.Append(SetImports("System.Data"))
        strSET.Append(SetImports("System.Data.Common"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.Collections"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("Commons"))
        strSET.Append(PartialClass("CrossDB"))

        strSET.Append(SetFunSummary(CrossInfo(0)("Summary")))
        If CrossParam.Count > 0 Then
            For i = 0 To CrossParam.Count - 1
                strSET.Append(SetParamSummary(CrossParam(i)("Name"), CrossParam(i)("Summary")))
            Next
        End If
        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
        strSET.Append(PublicFun(CrossInfo(0)("Name"), GetIlist(strClass), InputKEY))
        strSET.Append(GetDB)
        strSET.Append(GetLine())
        strSET.Append(DimAsNew("sqlStatement", "StringBuilder"))

        Dim arySQL As Array = CrossInfo(0)("SqlText").Split(GetLine())
        For i = 0 To arySQL.Length - 1
            strSET.Append(SqlAppend(Trim(arySQL(i)).ToString.Replace(Chr(10), "") & " "))
        Next
        strSET.Append(GetLine())
        strSET.Append(DimAsValue("dbCommand", "DbCommand", "db.GetSqlStringCommand(sqlStatement.ToString())"))
        If CrossParam.Count > 0 Then
            For i = 0 To CrossParam.Count - 1
                strSET.Append(DbAddInParameter(CrossParam(i)("Name"), CrossParam(i)("Type"), CrossParam(i)("Name")))
            Next
        End If
        strSET.Append(GetLine())
        strSET.Append(DimAsNewList("myList", strClass))
        strSET.Append(GetLine())
        strSET.Append(GetTry)
        strSET.Append(UsingValue("dataReader", "IDataReader", "db.ExecuteReader(dbCommand)"))
        strSET.Append(WhileValue("dataReader.Read()"))
        strSET.Append(SetString("myList.Add(Bind" & strClass & "Info(dataReader))", "                    "))
        strSET.Append(EndWhile)
        strSET.Append(EndUsing)

        strSET.Append(SetCatch("CrossDB." & CrossInfo(0)("Name")))
        strSET.Append(ReturnValue("myList"))
        strSET.Append(EndFun)
        strSET.Append(GetLine())

        'BindInfo
        strSET.Append(PrivateFun("Bind" & strClass & "Info", strClass & "Info", ByValAs("dr", "IDataReader")))
        strSET.Append(DimAsNew("myInfo", strClass & "Info"))
        strSET.Append(GetLine())
        strSET.Append(GetTry)
        For i = 0 To dt.Columns.Count - 1
            strSET.Append(SetCrossInfoColumns(dt.Columns(i).ColumnName, dt.Columns(i).DataType.Name))
        Next
        strSET.Append(SetCatch("CrossDB.Bind" & strClass & "Info", "Exception"))
        strSET.Append(ReturnValue("myInfo"))
        strSET.Append(EndFun)
        strSET.Append(EndClass)

        If Directory.Exists("DataAccess") = False Then
            Directory.CreateDirectory("DataAccess")
        End If

        If File.Exists("DataAccess\" & strClass & "DB." & CodeType) = True Then
            File.Delete("DataAccess\" & strClass & "DB." & CodeType)
        End If

        File.WriteAllText("DataAccess\" & strClass & "DB." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetCrossBiz()

        strNameSpace = "Business"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("DataAccess"))
        strSET.Append(PartialClass("CrossBiz"))
        strSET.Append(SetFunSummary(CrossInfo(0)("Summary")))
        If CrossParam.Count > 0 Then
            For i = 0 To CrossParam.Count - 1
                strSET.Append(SetParamSummary(CrossParam(i)("Name"), CrossParam(i)("Summary")))
            Next
        End If
        strSET.Append(SetReturnSummary(strClass & "Info的泛型集合"))
        strSET.Append(PublicSharedFun(CrossInfo(0)("Name"), GetIlist(strClass), InputKEY))
        strSET.Append(DimAsNew("myDB", "CrossDB"))
        strSET.Append(ReturnValue("myDB." & CrossInfo(0)("Name") & "(" & strCrossKEY & ")"))
        strSET.Append(EndFun)
        strSET.Append(EndClass)

        'System.IO.File.WriteAllText(strClass & "Biz." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

        If Directory.Exists("Business") = False Then
            Directory.CreateDirectory("Business")
        End If

        If File.Exists("Business\" & strClass & "Biz." & CodeType) = True Then
            File.Delete("Business\" & strClass & "Biz." & CodeType)
        End If

        File.WriteAllText("Business\" & strClass & "Biz." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
    Private Sub SetCrossWS()

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Web"))
        strSET.Append(SetImports("System.Web.Services"))
        strSET.Append(SetImports("System.Web.Services.Protocols"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.Collections"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("Business"))
        strSET.Append(WSPartialClass("CrossWS"))
        strSET.Append(SetFunSummary(CrossInfo(0)("Summary")))
        If CrossParam.Count > 0 Then
            For i = 0 To CrossParam.Count - 1
                strSET.Append(SetParamSummary(CrossParam(i)("Name"), CrossParam(i)("Summary")))
            Next
        End If
        strSET.Append(SetReturnSummary(strClass & "Info的物件集合"))
        strSET.Append(GetWebMethod(CrossInfo(0)("Summary")))
        strSET.Append(PublicFun(CrossInfo(0)("Name"), GetList(strClass), InputKEY))
        strSET.Append(ReturnValue(GetCType("CrossBiz." & CrossInfo(0)("Name") & "(" & strCrossKEY & ")", GetList(strClass))))
        strSET.Append(EndFun)
        strSET.Append(EndWSClass)

        If Directory.Exists("WebService") = False Then
            Directory.CreateDirectory("WebService")
        End If

        If Directory.Exists("WebService\App_Code") = False Then
            Directory.CreateDirectory("WebService\App_Code")
        End If

        If File.Exists("WebService\" & strClass & "WS.asmx") = True Then
            File.Delete("WebService\" & strClass & "WS.asmx")
        End If
        Dim str As String
        If CodeType = "cs" Then
            str = "<%@ WebService Language=""C#"" CodeBehind=""~/App_Code/" & strClass & "WS.cs"" Class=""CrossWS"" %>"
        Else
            str = "<%@ WebService Language=""VB"" CodeBehind=""~/App_Code/" & strClass & "WS.vb"" Class=""CrossWS"" %>"
        End If
        File.WriteAllText("WebService\" & strClass & "WS.asmx", str, Encoding.BigEndianUnicode)

        If File.Exists("WebService\App_Code\" & strClass & "WS." & CodeType) = True Then
            File.Delete("WebService\App_Code\" & strClass & "WS." & CodeType)
        End If
        File.WriteAllText("WebService\App_Code\" & strClass & "WS." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub
#End Region

#Region "MVC"
    Private Sub SetModel()

        strNameSpace = strProj

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.ComponentModel"))
        strSET.Append(SetImports("System.ComponentModel.DataAnnotations"))
        strSET.Append(SetImports("System.Linq"))
        strSET.Append(SetImports("System.Web"))
        strSET.Append(SetImports("System.Web.Mvc"))
        strSET.Append(SetImports("Information"))

        strSET.Append(PartialClass(strClass & "Model", TableSummary))

        strSET.Append("        public " & strClass & "Info Info { get; set; }" & vbCrLf)

        For i = 0 To dt.Columns.Count - 1
            Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
            Dim columnProperty As String
            If ColumnInfo.Count > 0 Then
                Dim inputType As String = Trim(ColumnInfo(0)("InputType"))

                If inputType <> "" Then

                    If Trim(ColumnInfo(0)("PropertyName")) <> "" Then
                        columnProperty = GetObjectName(ColumnInfo(0)("Name"))
                    Else
                        columnProperty = ColumnInfo(0)("PropertyName")
                    End If

                    Select Case inputType
                        Case "DropDownList", "RadioButtonList"
                            strSET.Append(vbCrLf)
                            strSET.Append("        public IEnumerable<SelectListItem> " & columnProperty & "Select { get; set; }" & vbCrLf)
                        Case "File"
                            strSET.Append(vbCrLf)
                            strSET.Append("        public HttpPostedFileBase " & columnProperty & "File { get; set; }" & vbCrLf)
                            'Case "Password"
                            'Case "Date"
                            'Case "Email"
                            'Case "Text"
                            'Case "Hidden"
                            'Case "CheckBox"
                    End Select
                End If
            End If
        Next

        strSET.Append(EndNoNSClass)
        strSET.Append(vbCrLf)

        strSET.Append("    public class " & strClass & "ListModel" & vbCrLf)
        strSET.Append("    {" & vbCrLf)
        strSET.Append("        public IList<" & strClass & "Info> List { get; set; }" & vbCrLf)
        strSET.Append(EndClass)

        Dim dir As String = "Models"

        If Directory.Exists(dir) = False Then
            Directory.CreateDirectory(dir)
        End If

        If File.Exists(dir & "\" & strClass & "Model." & CodeType) = True Then
            File.Delete(dir & "\" & strClass & "Model." & CodeType)
        End If

        File.WriteAllText(dir & "\" & strClass & "Model." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)
    End Sub

    Private Sub SetController()
        strNameSpace = strProj & ".Controllers"

        Dim strSET As New StringBuilder
        strSET.Append(SetImports("System"))
        strSET.Append(SetImports("System.Collections.Generic"))
        strSET.Append(SetImports("System.Linq"))
        strSET.Append(SetImports("System.Web"))
        strSET.Append(SetImports("System.Web.Mvc"))
        strSET.Append(SetImports("Business"))
        strSET.Append(SetImports("Information"))
        strSET.Append(SetImports("System.IO"))

        strSET.Append(InheritsClass(strClass & "Controller", "", "Controller"))

        Dim _inputKEY As String = InputKEY.Replace("int", "int?")

        strSET.Append("        public ActionResult Index()" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            " & strClass & "ListModel model = new " & strClass & "ListModel();" & vbCrLf)
        strSET.Append("            model.List = " & strClass & "Biz.GetAll();" & vbCrLf)
        strSET.Append("            return View(model);" & vbCrLf)
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("        public ActionResult Details(" & _inputKEY & ")" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            if (")
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(" && ")
            End If
            strSET.Append(ColumnProperty(ArrayKEY(i)) & " != null")
        Next
        strSET.Append(")" & vbCrLf)
        strSET.Append("            {" & vbCrLf)
        strSET.Append("                " & strClass & "Model model = new " & strClass & "Model();" & vbCrLf)
        strSET.Append("                model.Info = " & strClass & "Biz.GetInfo(")
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(", ")
            End If
            Select Case dt.Columns(ArrayKEY(i)).DataType.Name
                Case "Int32", "Int64", "Integer"
                    strSET.Append("(int)" & ColumnProperty(ArrayKEY(i)))
                Case Else
                    strSET.Append(ColumnProperty(ArrayKEY(i)))
            End Select
        Next
        strSET.Append(");" & vbCrLf)
        strSET.Append("                return View(model);" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        strSET.Append("            return RedirectToAction(""Index"");" & vbCrLf)
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("        public ActionResult Create()" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        If HasMvcInputControl Then
            strSET.Append("            " & strClass & "Model model = new " & strClass & "Model();" & vbCrLf)
            If HasMvcInputControl Then
                strSET.Append("            bindEdit(model);" & vbCrLf)
            End If
            strSET.Append("            return View(""Edit"", model);" & vbCrLf)
        Else
            strSET.Append("            return View(""Edit"");" & vbCrLf)
        End If
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("        [HttpPost]" & vbCrLf)
        strSET.Append("        public ActionResult Create(" & strClass & "Model model)" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            if (!ModelState.IsValid)" & vbCrLf)
        strSET.Append("            {" & vbCrLf)
        If HasMvcInputControl Then
            strSET.Append("                bindEdit(model);" & vbCrLf)
        End If
        strSET.Append("                return View(""Edit"", model);" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        For i = 0 To dt.Columns.Count - 1
            Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
            If ColumnInfo.Count > 0 Then
                Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                If inputType = "" Then
                    strSET.Append("            //model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & " = ;" & vbCrLf)
                Else
                    Select Case inputType
                        Case "File"
                            Dim fileProperty As String = ColumnProperty(dt.Columns(i).ColumnName) & "File"
                            strSET.Append("            if (model." & fileProperty & " != null)" & vbCrLf)
                            strSET.Append("            {" & vbCrLf)
                            strSET.Append("                var " & ColumnProperty(dt.Columns(i).ColumnName) & "Path = HttpContext.Server.MapPath(""~/SomeDirectory"");" & vbCrLf)
                            strSET.Append("                if (!Directory.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path))" & vbCrLf)
                            strSET.Append("                {" & vbCrLf)
                            strSET.Append("                    Directory.CreateDirectory(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path);" & vbCrLf)
                            strSET.Append("                }" & vbCrLf)
                            strSET.Append("                var " & ColumnProperty(dt.Columns(i).ColumnName) & "FileName = Path.Combine(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path, model." & fileProperty & ".FileName);" & vbCrLf)
                            strSET.Append("                if (System.IO.File.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName))" & vbCrLf)
                            strSET.Append("                {" & vbCrLf)
                            strSET.Append("                    System.IO.File.Delete(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                            strSET.Append("                }" & vbCrLf)
                            strSET.Append("                model." & ColumnProperty(dt.Columns(i).ColumnName) & "File.SaveAs(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                            strSET.Append("                model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & " = model." & fileProperty & ".FileName;" & vbCrLf)
                            strSET.Append("            }" & vbCrLf)
                    End Select
                End If
            End If
        Next
        strSET.Append("            " & strClass & "Biz.AddNew(model.Info);" & vbCrLf)
        strSET.Append("            return RedirectToAction(""Index"");" & vbCrLf)
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("        public ActionResult Edit(" & _inputKEY & ")" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            if (")
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(" && ")
            End If
            strSET.Append(ColumnProperty(ArrayKEY(i)) & " != null")
        Next
        strSET.Append(")" & vbCrLf)
        strSET.Append("            {" & vbCrLf)

        strSET.Append("                " & strClass & "Model model = new " & strClass & "Model();" & vbCrLf)
        strSET.Append("                model.Info = " & strClass & "Biz.GetInfo(")
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(", ")
            End If
            Select Case dt.Columns(ArrayKEY(i)).DataType.Name
                Case "Int32", "Int64", "Integer"
                    strSET.Append("(int)" & ColumnProperty(ArrayKEY(i)))
                Case Else
                    strSET.Append(ColumnProperty(ArrayKEY(i)))
            End Select
        Next
        strSET.Append(");" & vbCrLf)
        If HasMvcInputControl Then
            strSET.Append("                bindEdit(model);" & vbCrLf)
        End If
        strSET.Append("                return View(model);" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        'strSET.Append("            else" & vbCrLf)
        'strSET.Append("            {" & vbCrLf)
        'strSET.Append("                return View();" & vbCrLf)
        'strSET.Append("            }" & vbCrLf)
        strSET.Append("            return RedirectToAction(""Index"");" & vbCrLf)
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        Dim mvcKEY As String = mdFile.strKEY.Insert(0, "model.Info.")
        mvcKEY = mvcKEY.Replace(", ", ", model.Info.")
        strSET.Append("        [HttpPost]" & vbCrLf)
        strSET.Append("        public ActionResult Edit(" & strClass & "Model model)" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            if (!ModelState.IsValid)" & vbCrLf)
        strSET.Append("            {" & vbCrLf)
        If HasMvcInputControl Then
            strSET.Append("                bindEdit(model);" & vbCrLf)
        End If
        strSET.Append("                return View(model);" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
            strSET.Append("            var info = " & strClass & "Biz.GetInfo(" & mvcKEY & ");" & vbCrLf)
            For i = 0 To dt.Columns.Count - 1
                Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
                If ColumnInfo.Count > 0 Then
                    Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                    If inputType = "" Then
                        strSET.Append("            //info." & ColumnProperty(dt.Columns(i).ColumnName) & " = ;" & vbCrLf)
                    Else
                        Select Case inputType
                            Case "File"
                                Dim fileProperty As String = ColumnProperty(dt.Columns(i).ColumnName) & "File"
                                strSET.Append("            if (model." & fileProperty & " != null)" & vbCrLf)
                                strSET.Append("            {" & vbCrLf)
                                strSET.Append("                var " & ColumnProperty(dt.Columns(i).ColumnName) & "Path = HttpContext.Server.MapPath(""~/SomeDirectory"");" & vbCrLf)
                                strSET.Append("                if (!Directory.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path))" & vbCrLf)
                                strSET.Append("                {" & vbCrLf)
                                strSET.Append("                    Directory.CreateDirectory(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path);" & vbCrLf)
                                strSET.Append("                }" & vbCrLf)
                                strSET.Append("                var " & ColumnProperty(dt.Columns(i).ColumnName) & "FileName = Path.Combine(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path, model." & fileProperty & ".FileName);" & vbCrLf)
                                strSET.Append("                if (System.IO.File.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName))" & vbCrLf)
                                strSET.Append("                {" & vbCrLf)
                                strSET.Append("                    System.IO.File.Delete(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                                strSET.Append("                }" & vbCrLf)
                                strSET.Append("                model." & ColumnProperty(dt.Columns(i).ColumnName) & "File.SaveAs(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                                strSET.Append("                info." & ColumnProperty(dt.Columns(i).ColumnName) & " = model." & fileProperty & ".FileName;" & vbCrLf)
                                strSET.Append("            }" & vbCrLf)
                            Case Else
                                strSET.Append("            info." & ColumnProperty(dt.Columns(i).ColumnName) & " = model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & ";" & vbCrLf)
                        End Select
                    End If
                End If
            Next
            strSET.Append("            " & strClass & "Biz.Update(info);" & vbCrLf)
            strSET.Append("            return RedirectToAction(""Index"");" & vbCrLf)
        Else
            strSET.Append("            if (" & strClass & "Biz.Exists(" & mvcKEY & "))" & vbCrLf)
            strSET.Append("            {" & vbCrLf)
            strSET.Append("                var info = " & strClass & "Biz.GetInfo(" & mvcKEY & ");" & GetLine())

            For i = 0 To dt.Columns.Count - 1
                Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
                If ColumnInfo.Count > 0 Then
                    Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                    If inputType = "" Then
                        strSET.Append("                //info." & ColumnProperty(dt.Columns(i).ColumnName) & " = ;" & vbCrLf)
                    Else
                        Select Case inputType
                            Case "File"
                                Dim fileProperty As String = ColumnProperty(dt.Columns(i).ColumnName) & "File"
                                strSET.Append("                if (model." & fileProperty & " != null)" & vbCrLf)
                                strSET.Append("                {" & vbCrLf)
                                strSET.Append("                    var " & ColumnProperty(dt.Columns(i).ColumnName) & "Path = HttpContext.Server.MapPath(""~/SomeDirectory"");" & vbCrLf)
                                strSET.Append("                    if (!Directory.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path))" & vbCrLf)
                                strSET.Append("                    {" & vbCrLf)
                                strSET.Append("                        Directory.CreateDirectory(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path);" & vbCrLf)
                                strSET.Append("                    }" & vbCrLf)
                                strSET.Append("                    var " & ColumnProperty(dt.Columns(i).ColumnName) & "FileName = Path.Combine(" & ColumnProperty(dt.Columns(i).ColumnName) & "Path, model." & fileProperty & ".FileName);" & vbCrLf)
                                strSET.Append("                    if (System.IO.File.Exists(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName))" & vbCrLf)
                                strSET.Append("                    {" & vbCrLf)
                                strSET.Append("                        System.IO.File.Delete(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                                strSET.Append("                    }" & vbCrLf)
                                strSET.Append("                    model." & ColumnProperty(dt.Columns(i).ColumnName) & "File.SaveAs(" & ColumnProperty(dt.Columns(i).ColumnName) & "FileName);" & vbCrLf)
                                strSET.Append("                    info." & ColumnProperty(dt.Columns(i).ColumnName) & " = model." & fileProperty & ".FileName;" & vbCrLf)
                                strSET.Append("                }" & vbCrLf)
                            Case Else
                                strSET.Append("                info." & ColumnProperty(dt.Columns(i).ColumnName) & " = model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & ";" & vbCrLf)
                        End Select
                    End If
                End If
            Next
            strSET.Append("                " & strClass & "Biz.Update(info);" & vbCrLf)
            strSET.Append("            }" & vbCrLf)
            strSET.Append("            return RedirectToAction(""Index"");" & vbCrLf)
        End If
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("        public ActionResult Delete(" & InputKEY & ")" & vbCrLf)
        strSET.Append("        {" & vbCrLf)
        strSET.Append("            if (" & strClass & "Biz.Del(" & strKEY & "))" & vbCrLf)
        strSET.Append("            {" & vbCrLf)
        strSET.Append("                return RedirectToAction(""Index"");" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        strSET.Append("            else" & vbCrLf)
        strSET.Append("            {" & vbCrLf)
        strSET.Append("                return JavaScript(""alert('刪除失敗');"");" & vbCrLf)
        strSET.Append("            }" & vbCrLf)
        strSET.Append("        }" & vbCrLf)
        strSET.Append(vbCrLf)

        If HasMvcInputControl Then
            strSET.Append("        void bindEdit(" & strClass & "Model model)" & vbCrLf)
            strSET.Append("        {" & vbCrLf)
            For i = 0 To dt.Columns.Count - 1
                Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
                If ColumnInfo.Count > 0 Then
                    Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                    If inputType <> "" Then
                        Select Case inputType
                            Case "DropDownList", "RadioButtonList"
                                strSET.Append("            //model." & ColumnProperty(dt.Columns(i).ColumnName) & "Select = SomeClassBiz.SomeMethod().Select(s => new SelectListItem" & vbCrLf)
                                strSET.Append("            //{" & vbCrLf)
                                strSET.Append("            //    Value = s.Column1," & vbCrLf)
                                strSET.Append("            //    Text = s.Column2" & vbCrLf)
                                strSET.Append("            //});" & vbCrLf)
                        End Select
                    End If
                End If
            Next
            strSET.Append("        }" & vbCrLf)
        End If
        strSET.Append(EndClass)

        Dim dir As String = "Controllers"

        If Directory.Exists(dir) = False Then
            Directory.CreateDirectory(dir)
        End If

        If File.Exists(dir & "\" & strClass & "Controller." & CodeType) = True Then
            File.Delete(dir & "\" & strClass & "Controller." & CodeType)
        End If

        File.WriteAllText(dir & "\" & strClass & "Controller." & CodeType, strSET.ToString, Encoding.BigEndianUnicode)

    End Sub

    Private Sub SetViewIndex()
        Dim strSET As New StringBuilder
        strSET.Append("@model " & strClass & "ListModel" & vbCrLf)
        strSET.Append("@{" & vbCrLf)

        Dim tName As String
        If String.IsNullOrEmpty(TableSummary()) Then
            tName = strClass
        Else
            tName = TableSummary()
        End If

        strSET.Append("    ViewBag.Title = """ & tName & """;" & vbCrLf)
        strSET.Append("}" & vbCrLf)
        strSET.Append(vbCrLf)
        strSET.Append("<h2>" & tName & "</h2>" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("<p>" & vbCrLf)
        strSET.Append("    @Html.ActionLink(""Create New"", ""Create"")" & vbCrLf)
        strSET.Append("</p>" & vbCrLf)
        strSET.Append("<table class=""table"">" & vbCrLf)
        strSET.Append("    <tr>" & vbCrLf)
        For i = 0 To dt.Columns.Count - 1
            strSET.Append("        <th>" & vbCrLf)
            strSET.Append("            @Html.DisplayNameFor(model => model.List.FirstOrDefault()." & ColumnProperty(dt.Columns(i).ColumnName) & ")" & vbCrLf)
            strSET.Append("        </th>" & vbCrLf)
        Next
        strSET.Append("        <th>Edit</th>" & vbCrLf)
        strSET.Append("    </tr>" & vbCrLf)
        strSET.Append("    @foreach (var item in Model.List)" & vbCrLf)
        strSET.Append("    {" & vbCrLf)
        strSET.Append("        <tr>" & vbCrLf)
        For i = 0 To dt.Columns.Count - 1
            strSET.Append("            <td>" & vbCrLf)
            strSET.Append("                @Html.DisplayFor(modelItem => item." & ColumnProperty(dt.Columns(i).ColumnName) & ")" & vbCrLf)
            strSET.Append("            </td>" & vbCrLf)
        Next
        Dim strSET2 As New StringBuilder
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET2.Append(", ")
            End If
            strSET2.Append(ColumnProperty(ArrayKEY(i)) & " = item." & ColumnProperty(ArrayKEY(i)))
        Next
        strSET.Append("            <td>" & vbCrLf)
        strSET.Append("                @Html.ActionLink(""Edit"", ""Edit"", new { " & strSET2.ToString() & " }) |" & vbCrLf)
        strSET.Append("                @Html.ActionLink(""Details"", ""Details"", new { " & strSET2.ToString() & " }) |" & vbCrLf)
        strSET.Append("                @Html.ActionLink(""Delete"", ""Delete"", new { " & strSET2.ToString() & " })" & vbCrLf)
        strSET.Append("            </td>" & vbCrLf)
        strSET.Append("        </tr>" & vbCrLf)
        strSET.Append("    }" & vbCrLf)
        strSET.Append("</table>" & vbCrLf)

        Dim dir As String = "Views" & "\" & strClass

        If Directory.Exists(dir) = False Then
            Directory.CreateDirectory(dir)
        End If

        If File.Exists(dir & "\Index.cshtml") = True Then
            File.Delete(dir & "\Index.cshtml")
        End If

        File.WriteAllText(dir & "\Index.cshtml", strSET.ToString, Encoding.BigEndianUnicode)
    End Sub

    Private Sub SetViewDetails()
        Dim strSET As New StringBuilder
        strSET.Append("@model " & strClass & "Model" & vbCrLf)
        strSET.Append("@{" & vbCrLf)

        Dim tName As String
        If String.IsNullOrEmpty(TableSummary()) Then
            tName = strClass
        Else
            tName = TableSummary()
        End If

        strSET.Append("    ViewBag.Title = """ & tName & """;" & vbCrLf)
        strSET.Append("}" & vbCrLf)
        strSET.Append(vbCrLf)
        strSET.Append("<h2>" & tName & "</h2>" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("<div>" & vbCrLf)
        strSET.Append("    <h4>" & tName & "</h4>" & vbCrLf)
        strSET.Append("    <hr />" & vbCrLf)

        strSET.Append("    <dl class=""dl-horizontal"">" & vbCrLf)
        For i = 0 To dt.Columns.Count - 1
            strSET.Append("        <dt>" & vbCrLf)
            strSET.Append("            @Html.DisplayNameFor(model => model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & ")" & vbCrLf)
            strSET.Append("        </dt>" & vbCrLf)
            strSET.Append("        <dd>" & vbCrLf)


            Dim isFile As Boolean = False
            Dim ColumnInfo As DataView = GetColumnInfoByName(dt.Columns(i).ColumnName).DefaultView
            If ColumnInfo.Count > 0 Then
                Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                If inputType = "File" Then
                    isFile = True
                End If
            End If

            If isFile Then
                strSET.Append("            @if (!string.IsNullOrWhiteSpace(Model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & ")) {" & vbCrLf)
                strSET.Append("                <img src=""~/SomeDirectory/@Model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & """ />" & vbCrLf)
                strSET.Append("            }" & vbCrLf)
            Else
                strSET.Append("            @Html.DisplayFor(model => model.Info." & ColumnProperty(dt.Columns(i).ColumnName) & ")" & vbCrLf)
            End If


            strSET.Append("        </dd>" & vbCrLf)
        Next
        strSET.Append("    </dl>" & vbCrLf)
        strSET.Append("</div>" & vbCrLf)
        strSET.Append("<p>" & vbCrLf)
        strSET.Append("    @Html.ActionLink(""Edit"", ""Edit"", new { ")
        For i = 0 To ArrayKEY.Length - 1
            If i > 0 Then
                strSET.Append(", ")
            End If
            strSET.Append(ColumnProperty(ArrayKEY(i)) & " = Model.Info." & ColumnProperty(ArrayKEY(i)))
        Next
        strSET.Append(" }) |" & vbCrLf)
        strSET.Append("    @Html.ActionLink(""Back to List"", ""Index"")" & vbCrLf)
        strSET.Append("</p>" & vbCrLf)

        Dim dir As String = "Views" & "\" & strClass

        If Directory.Exists(dir) = False Then
            Directory.CreateDirectory(dir)
        End If

        If File.Exists(dir & "\Details.cshtml") = True Then
            File.Delete(dir & "\Details.cshtml")
        End If

        File.WriteAllText(dir & "\Details.cshtml", strSET.ToString, Encoding.BigEndianUnicode)
    End Sub

    Private Sub SetViewEdit()
        Dim strSET As New StringBuilder
        strSET.Append("@model " & strClass & "Model" & vbCrLf)
        strSET.Append("@{" & vbCrLf)

        Dim tName As String
        If String.IsNullOrEmpty(TableSummary()) Then
            tName = strClass
        Else
            tName = TableSummary()
        End If

        strSET.Append("    ViewBag.Title = """ & tName & """;" & vbCrLf)
        strSET.Append("    ViewBag.Edit = (ViewContext.RouteData.Values[""action""].ToString() == ""Edit"");" & vbCrLf)
        strSET.Append("}" & vbCrLf)
        strSET.Append(vbCrLf)
        strSET.Append("<h2>" & tName & "</h2>" & vbCrLf)
        strSET.Append(vbCrLf)

        strSET.Append("[FormString]" & vbCrLf)

        strSET.Append("{" & vbCrLf)
        strSET.Append("    @Html.AntiForgeryToken()" & vbCrLf)
        strSET.Append(vbCrLf)
        strSET.Append("    <div class=""form-horizontal"">" & vbCrLf)
        strSET.Append("        <h4>" & tName & "</h4>" & vbCrLf)
        strSET.Append("        <hr />" & vbCrLf)
        strSET.Append("        @Html.ValidationSummary(true)" & vbCrLf)
        Dim hasFile As Boolean = False

        For i = 0 To ArrayNOKEY.Length - 1
            Dim ColumnInfo As DataView = GetColumnInfoByName(ArrayNOKEY(i)).DefaultView
            If ColumnInfo.Count > 0 Then
                Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                If inputType <> "" Then
                    strSET.Append(getInput(ArrayNOKEY(i), inputType))
                    If inputType = "File" Then
                        hasFile = True
                    End If
                End If
            End If
        Next

        If HasAutoIncrement <> "" And mdFile.ArrayKEY.Length = 1 Then
            strSET.Append("                @if (ViewBag.Edit)" & vbCrLf)
            strSET.Append("                {" & vbCrLf)
            strSET.Append("                    @Html.HiddenFor(model => model.Info." & ColumnProperty(HasAutoIncrement) & ")" & vbCrLf)
            strSET.Append("                    @Html.ValidationMessageFor(model => model.Info." & ColumnProperty(HasAutoIncrement) & ")" & vbCrLf)
            strSET.Append("                }" & vbCrLf)
        Else
            For i = 0 To ArrayKEY.Length - 1
                Dim ColumnInfo As DataView = GetColumnInfoByName(ArrayKEY(i)).DefaultView
                If ColumnInfo.Count > 0 Then
                    Dim inputType As String = Trim(ColumnInfo(0)("InputType"))
                    If inputType <> "" Then
                        strSET.Append(getInput(ArrayKEY(i), inputType, True))
                    End If
                End If
            Next
        End If

        If hasFile Then
            strSET.Replace("[FormString]", "@using (Html.BeginForm(null, null, FormMethod.Post, new { enctype = ""multipart/form-data"" }))")
        Else
            strSET.Replace("[FormString]", "@using (Html.BeginForm())")
        End If

        strSET.Append("        <div class=""form-group"">" & vbCrLf)
        strSET.Append("            <div class=""col-md-offset-2 col-md-10"">" & vbCrLf)
        strSET.Append("                <input type=""submit"" value=""Save"" class=""btn btn-default"" />" & vbCrLf)
        'strSET.Append("                @Html.ActionLink(""Back to List"", ""Index"")" & vbCrLf)
        strSET.Append("                <input" & vbCrLf)
        strSET.Append("                    Type = ""button""" & vbCrLf)
        strSET.Append("                    value = ""Back to List""" & vbCrLf)
        strSET.Append("                    class=""btn btn-default""" & vbCrLf)
        strSET.Append("                    onclick=""location.href='@Url.Action(""Index"")';"" />" & vbCrLf)
        strSET.Append("            </div>" & vbCrLf)
        strSET.Append("        </div>" & vbCrLf)
        strSET.Append("    </div>" & vbCrLf)

        strSET.Append("}" & vbCrLf)
        strSET.Append(vbCrLf)
        strSET.Append("<script src=""~/Scripts/jquery-1.10.2.min.js""></script>" & vbCrLf)
        strSET.Append("<script src=""~/Scripts/jquery.validate.min.js""></script>" & vbCrLf)
        strSET.Append("<script src=""~/Scripts/jquery.validate.unobtrusive.min.js""></script>" & vbCrLf)

        Dim dir As String = "Views" & "\" & strClass

        If Directory.Exists(dir) = False Then
            Directory.CreateDirectory(dir)
        End If

        If File.Exists(dir & "\Edit.cshtml") = True Then
            File.Delete(dir & "\Edit.cshtml")
        End If

        File.WriteAllText(dir & "\Edit.cshtml", strSET.ToString, Encoding.BigEndianUnicode)
    End Sub

    Private Function GetPropertyRange(ByVal row As DataRowView, ByVal dataType As String) As String

        Dim str As String = ""

        Dim min As Integer
        If row("MinVal").Equals(DBNull.Value) Then
            min = 0
        Else
            min = row("MinVal")
        End If

        Dim max As Integer
        If row("MaxVal").Equals(DBNull.Value) Then
            max = 0
        Else
            max = row("MaxVal")
        End If

        If max > 0 Or min > 0 Then
            Select Case dataType
                Case "Int32", "Int64", "Integer", "Decimal", "Double"
                    str = "        [Range(" & min & ", " & max & ", ErrorMessage = ""請輸入數值" & min & "~" & max & """)]"
                Case "String"
                    str = "        [StringLength(" & max & ", MinimumLength = " & min & ", ErrorMessage = """ & row("Summary") & "長度需在" & min & "~" & max & "個字元內"")]"
            End Select
        End If

        If str <> "" Then
            str &= vbCrLf
        End If

        Return str

    End Function

    Function getInput(ByVal colName As String, ByVal inputType As String, Optional ByVal isReadOnly As Boolean = False)
        Dim strSET As New StringBuilder

        strSET.Append("        <div class=""form-group"">" & vbCrLf)
        strSET.Append("            @Html.LabelFor(model => model.Info." & ColumnProperty(colName) & ", new { @class = ""control-label col-md-2"" })" & vbCrLf)
        strSET.Append("            <div class=""col-md-10"">" & vbCrLf)

        Select Case inputType
            Case "Date"
                strSET.Append("                @Html.TextBoxFor(model => model.Info." & ColumnProperty(colName) & ",""{0:yyyy-MM-dd}"", new { type = ""date"" })" & vbCrLf)
            Case "Text", "Email"
                If isReadOnly Then
                    strSET.Append("                @if (ViewBag.Edit)" & vbCrLf)
                    strSET.Append("                {" & vbCrLf)
                    strSET.Append("                    @Html.TextBoxFor(model => model.Info." & ColumnProperty(colName) & ", new { @readonly=""readonly"" })" & vbCrLf)
                    strSET.Append("                }" & vbCrLf)
                    strSET.Append("                else")
                    strSET.Append("                {" & vbCrLf)
                    strSET.Append("                    @Html.TextBoxFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
                    strSET.Append("                }" & vbCrLf)
                Else
                    strSET.Append("                @Html.EditorFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
                End If
            Case "Password"
                strSET.Append("                @Html.PasswordFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
            Case "DropDownList"
                strSET.Append("                @Html.DropDownListFor(model => model.Info." & ColumnProperty(colName) & ", Model." & ColumnProperty(colName) & "Select, ""請選擇"")" & vbCrLf)
            Case "RadioButtonList"
                strSET.Append("                @foreach (var item in Model." & ColumnProperty(colName) & "Select) {" & vbCrLf)
                strSET.Append("                    @Html.RadioButtonFor(model => model.Info." & ColumnProperty(colName) & ", item.Value)@Html.Label(item.Text)" & vbCrLf)
                strSET.Append("                }" & vbCrLf)
            Case "File"
                strSET.Append("                @if (ViewBag.Edit && !string.IsNullOrWhiteSpace(Model.Info." & ColumnProperty(colName) & ")) {" & vbCrLf)
                strSET.Append("                    <img src=""~/SomeDirectory/@Model.Info." & ColumnProperty(colName) & """ />" & vbCrLf)
                strSET.Append("                }" & vbCrLf)
                strSET.Append("                @Html.TextBoxFor(model => model." & ColumnProperty(colName) & "File, new { type = ""file"" })" & vbCrLf)
            Case "Hidden"
                strSET.Append("                @Html.HiddenFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
            Case "CheckBox"
                strSET.Append("                @Html.CheckBoxFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
        End Select

        If inputType = "File" Then
            strSET.Append("                @Html.ValidationMessageFor(model => model." & ColumnProperty(colName) & "File)" & vbCrLf)
        Else
            strSET.Append("                @Html.ValidationMessageFor(model => model.Info." & ColumnProperty(colName) & ")" & vbCrLf)
        End If

        strSET.Append("            </div>" & vbCrLf)
        strSET.Append("        </div>" & vbCrLf)

        Return strSET.ToString
    End Function

#End Region

End Module

