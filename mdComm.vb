Module mdComm
    Public Function GetObjectName(ByVal strTableClass As String) As String

        Dim str As String = ""
        Dim ary1 As Array = strTableClass.Split("_")

        If ary1.Length > 1 Then
            Dim i, j As Integer
            For i = 0 To ary1.Length - 1
                Dim ary2 As String = ary1(i)
                If ary2.Length > 0 Then
                    For j = 0 To ary2.Length - 1
                        If j = 0 Then
                            str &= ary2(j).ToString.ToUpper
                        Else
                            str &= ary2(j).ToString.ToLower
                        End If
                    Next
                End If
            Next
        Else
            str = strTableClass
        End If

        Return str

    End Function

    Public Function GetValue(ByRef mObject As Object) As Object

        If IsDBNull(mObject) = True Then
            mObject = ""
        End If

        Return mObject

    End Function

    Public Sub SetColData(ByVal TableName As String, Optional ByVal intTId As Integer = -1)

        Dim i As Integer

        Dim dv As DataView = GetColumnDv(intTId).DefaultView
        For i = 0 To dv.Count - 1
            If ColumnSourceInfo(TableName, dv(i)("Name")).DefaultView.Count = 0 Then
                DeleteColumn(dv(i)("ColId"))
            End If
        Next

        Dim dvSource As DataView = ColumnsSource(TableName).DefaultView

        Dim dt As DataTable = TablesSchema(TableName)

        For i = 0 To dvSource.Count - 1
 
            Dim colName As String
            Select Case DbType
                Case "SQLITE"
                    colName = "name"
                Case Else
                    colName = "COLUMN_NAME"
            End Select

            Dim defType As String
            If dt.Columns(dvSource(i)(colName)).AutoIncrement Then
                defType = "Hidden"
            Else
                defType = GetDefInputType(dt.Columns(dvSource(i)(colName)).DataType.Name())
            End If

            If GetColumnInfoByName(dvSource(i)(colName)).DefaultView.Count = 0 Then
                AddColumn(dvSource(i)(colName),
                          GetObjectName(dvSource(i)(colName)),
                          "",
                          defType,
                          DBNull.Value,
                          DBNull.Value,
                          False,
                          "")
            End If

        Next

    End Sub

    Public Function TableName() As String

        Dim TableInfo As DataView = GetTableInfo.DefaultView

        Return TableInfo(0)("Name")

    End Function

    Public Function TableClass() As String

        Dim TableInfo As DataView = GetTableInfo.DefaultView

        If Trim(TableInfo(0)("ClassName")) = "" Then
            Return GetObjectName(TableInfo(0)("Name"))
        Else
            Return TableInfo(0)("ClassName")
        End If

    End Function

    Public Function TableClass(ByVal TableName As String) As String

        Dim TableInfo As DataView = GetTableInfoByName(TableName).DefaultView

        If Trim(TableInfo(0)("ClassName")) = "" Then
            Return GetObjectName(TableInfo(0)("Name"))
        Else
            Return TableInfo(0)("ClassName")
        End If

    End Function

    Public Function TableSummary() As String

        Dim TableInfo As DataView = GetTableInfo.DefaultView
        Return TableInfo(0)("Summary")

    End Function

    Public Function ColumnProperty(ByVal strColumn As String) As String

        Dim ColumnInfo As DataView = GetColumnInfoByName(strColumn).DefaultView

        If ColumnInfo.Count > 0 Then
            If Trim(ColumnInfo(0)("PropertyName")) = "" Then
                Return GetObjectName(ColumnInfo(0)("Name"))
            Else
                Return ColumnInfo(0)("PropertyName")
            End If
        Else
            Return GetObjectName(strColumn)
        End If

    End Function

    Public Function ColumnSummary(ByVal strColumn As String) As String

        Dim ColumnInfo As DataView = GetColumnInfoByName(strColumn).DefaultView

        If ColumnInfo.Count > 0 Then
            If Trim(ColumnInfo(0)("Summary")) = "" Then
                Return ""
            Else
                Return ColumnInfo(0)("Summary")
            End If
        Else
            Return ""
        End If

    End Function

    Public Function ColumnRemarks(ByVal strColumn As String) As String

        Dim ColumnInfo As DataView = GetColumnInfoByName(strColumn).DefaultView

        If ColumnInfo.Count > 0 Then
            If IsDBNull(ColumnInfo(0)("Remark")) Then
                Return ""
            Else
                Return ColumnInfo(0)("Remark")
            End If
        Else
            Return ""
        End If

    End Function

    Public Function ReplaceSQL(ByVal strSQL As String, ByVal ParamName As String, ByVal ParamType As String) As String

        Dim strParamName As String = "@" & ParamName

        strSQL = strSQL.Replace(strParamName & " ", GetInitialValue(ParamType) & " ")
        strSQL = strSQL.Replace(strParamName & ">", GetInitialValue(ParamType) & ">")
        strSQL = strSQL.Replace(strParamName & "<", GetInitialValue(ParamType) & "<")
        strSQL = strSQL.Replace(strParamName & "=", GetInitialValue(ParamType) & "=")
        strSQL = strSQL.Replace(strParamName & ")", GetInitialValue(ParamType) & ")")
        strSQL = strSQL.Replace(strParamName & ",", GetInitialValue(ParamType) & ",")
        strSQL = strSQL.Replace(strParamName & Chr(10), GetInitialValue(ParamType) & Chr(10))
        strSQL = strSQL.Replace(strParamName & Chr(13), GetInitialValue(ParamType) & Chr(13))

        If strSQL.EndsWith(strParamName) = True Then
            strSQL = strSQL.Remove(strSQL.LastIndexOf(strParamName), strParamName.Length)
            strSQL &= GetInitialValue(ParamType)
        End If

        Return strSQL

    End Function

    Private Function GetInitialValue(ByVal DBsype As String) As String

        Select Case DBsype
            Case "Int16", "Int32", "Int64", "Decimal", "Double", "Boolean"
                Return "1"
            Case "String"
                Return "''"
            Case "DateTime"
                Return "'9999/12/31'"
            Case Else
                Return "''"
        End Select

    End Function

End Module
