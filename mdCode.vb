Imports System.Text
Module mdCode
    Private str As String

    Public Function PublicClass(ByVal strName As String, Optional ByVal strSummary As String = "") As String

        Select Case CodeType
            Case "vb"
                str = SetClassSummary(strSummary)
                str &= "Public Class " & strName & vbCrLf
            Case "cs"
                str = "namespace " & strNameSpace & vbCrLf
                str &= "{" & vbCrLf
                str &= SetClassSummary(strSummary)
                str &= "    public class " & strName & vbCrLf
                str &= "    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PartialClass(ByVal strName As String, Optional ByVal strSummary As String = "") As String

        Select Case CodeType
            Case "vb"
                str = SetClassSummary(strSummary)
                str &= "Partial Public Class " & strName & vbCrLf
            Case "cs"
                str = "namespace " & strNameSpace & vbCrLf
                str &= "{" & vbCrLf
                str &= SetClassSummary(strSummary)
                str &= "    public partial class " & strName & vbCrLf
                str &= "    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function WSClass(ByVal strName As String, Optional ByVal strSummary As String = "") As String

        Select Case CodeType
            Case "vb"
                str = SetClassSummary(strSummary)
                str &= "<WebService(Namespace:=""" & WSNameSpace & """, Description:=""" & strSummary & """)> _" & vbCrLf
                str &= "<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _" & vbCrLf
                str &= "<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _" & vbCrLf
                str &= "Public Class " & strName & vbCrLf
            Case "cs"
                str = SetClassSummary(strSummary)
                str &= "[WebService(Namespace = """ & WSNameSpace & """, Description = """ & strSummary & """)]" & vbCrLf
                str &= "[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]" & vbCrLf
                str &= "    public class " & strName & vbCrLf
                str &= "    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function WSPartialClass(ByVal strName As String, Optional ByVal strSummary As String = "") As String

        Select Case CodeType
            Case "vb"
                str = SetClassSummary(strSummary)
                str &= "<WebService(Namespace:=""" & WSNameSpace & """, Description:=""" & strSummary & """)> _" & vbCrLf
                str &= "<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _" & vbCrLf
                str &= "<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _" & vbCrLf
                str &= "Partial Public Class " & strName & vbCrLf
            Case "cs"
                str = SetClassSummary(strSummary)
                str &= "[WebService(Namespace = """ & WSNameSpace & """, Description = """ & strSummary & """)]" & vbCrLf
                str &= "[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]" & vbCrLf
                str &= "    public partial class " & strName & vbCrLf
                str &= "    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function InheritsBaseClass(ByVal strName As String, ByVal strSummary As String) As String
        Return InheritsClass(strName, strSummary, "baseInfo")
    End Function

    Public Function InheritsClass(ByVal strName As String, ByVal strSummary As String, ByVal baseClass As String) As String

        Select Case CodeType
            Case "vb"
                str = SetClassSummary(strSummary)
                str &= "Partial Public Class " & strName & vbCrLf
                str &= "    Inherits " & baseClass & vbCrLf
            Case "cs"
                str = "namespace " & strNameSpace & vbCrLf
                str &= "{" & vbCrLf
                str &= SetClassSummary(strSummary)
                str &= "    public partial class " & strName & " : " & baseClass & vbCrLf
                str &= "    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndClass() As String

        Select Case CodeType
            Case "vb"
                str = "End Class" & vbCrLf
            Case "cs"
                str = "    }" & vbCrLf
                str &= "}" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndNoNSClass() As String

        Select Case CodeType
            Case "vb"
                str = "End Class" & vbCrLf
            Case "cs"
                str = "    }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndWSClass() As String

        Select Case CodeType
            Case "vb"
                str = "End Class" & vbCrLf
            Case "cs"
                str = "    }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function ExistsResult() As String

        Select Case CodeType
            Case "vb"
                str = "        Dim cmdResult As Integer = Integer.Parse(db.ExecuteScalar(dbCommand).ToString())" & vbCrLf
                str &= "        If cmdResult = 0 Then" & vbCrLf
                str &= "            Return False" & vbCrLf
                str &= "        Else" & vbCrLf
                str &= "            Return True" & vbCrLf
                str &= "        End If" & vbCrLf
            Case "cs"
                str = "            int cmdResult = int.Parse(db.ExecuteScalar(dbCommand).ToString());" & vbCrLf
                str &= "            if (cmdResult == 0)" & vbCrLf
                str &= "            {" & vbCrLf
                str &= "                return false;" & vbCrLf
                str &= "            }" & vbCrLf
                str &= "            else" & vbCrLf
                str &= "            {" & vbCrLf
                str &= "                return true;" & vbCrLf
                str &= "            }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateAsType(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private " & strName & " As " & ConvertType(strType) & vbCrLf
            Case "cs"
                str = "        private " & ConvertType(strType) & " " & strName & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateAsValue(ByVal strName As String, ByVal strType As String, ByVal strValue As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private " & strName & " As " & ConvertType(strType) & " = " & ConvertValue(strValue) & vbCrLf
            Case "cs"
                str = "        private " & ConvertType(strType) & " " & strName & " = " & ConvertValue(strValue) & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateSharedAsNewType(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private Shared " & strName & " As " & strType & " = New " & strType & vbCrLf
            Case "cs"
                str = "        private static " & strType & " " & strName & " = new " & strType & "();" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateAsNewType(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private " & strName & " As " & strType & " = New " & strType & vbCrLf
            Case "cs"
                str = "        private " & strType & " " & strName & " = new " & strType & "();" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateAsNewList(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private " & strName & " As IList(Of " & strType & "Info) = New List(Of " & strType & "Info)" & vbCrLf
            Case "cs"
                str = "        private IList<" & strType & "Info> " & strName & " = new List<" & strType & "Info>();" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PublicFun(ByVal strName As String, ByVal strType As String, ByVal strParameter As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Public Function " & strName & "(" & strParameter & ") As " & ConvertType(strType) & vbCrLf
                str &= vbCrLf
            Case "cs"
                str = "        public " & ConvertType(strType) & " " & strName & "(" & strParameter & ")" & vbCrLf
                str &= "        {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PrivateFun(ByVal strName As String, ByVal strType As String, ByVal strParameter As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Private Function " & strName & "(" & strParameter & ") As " & ConvertType(strType) & vbCrLf
                str &= vbCrLf
            Case "cs"
                str = "        private " & ConvertType(strType) & " " & strName & "(" & strParameter & ")" & vbCrLf
                str &= "        {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function PublicSharedFun(ByVal strName As String, ByVal strType As String, ByVal strParameter As String) As String

        Select Case CodeType
            Case "vb"
                str = "    Public Shared Function " & strName & "(" & strParameter & ") As " & ConvertType(strType) & vbCrLf
                str &= vbCrLf
            Case "cs"
                str = "        public static " & ConvertType(strType) & " " & strName & "(" & strParameter & ")" & vbCrLf
                str &= "        {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndFun() As String

        Select Case CodeType
            Case "vb"
                str = vbCrLf
                str &= "    End Function" & vbCrLf
            Case "cs"
                str = "        }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function UsingValue(ByVal strName As String, ByVal strType As String, ByVal strValue As String) As String

        Select Case CodeType
            Case "vb"
                str = "            Using " & strName & " As " & ConvertType(strType) & " = " & ConvertValue(strValue) & vbCrLf
            Case "cs"
                str = "                using (" & ConvertType(strType) & " " & strName & " = " & ConvertValue(strValue) & ")" & vbCrLf
                str &= "                {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndUsing() As String

        Select Case CodeType
            Case "vb"
                str = "            End Using" & vbCrLf
            Case "cs"
                str = "                }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function WhileValue(ByVal strValue As String) As String

        Select Case CodeType
            Case "vb"
                str = "                While " & strValue & vbCrLf
            Case "cs"
                str = "                    while (" & strValue & ")" & vbCrLf
                str &= "                    {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function EndWhile() As String

        Select Case CodeType
            Case "vb"
                str = "                End While" & vbCrLf
            Case "cs"
                str = "                    }" & vbCrLf
        End Select

        Return str

    End Function

    Public Function ByValAs(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "ByVal " & strName & " As " & ConvertType(strType)
            Case "cs"
                str = ConvertType(strType) & " " & strName
        End Select

        Return str

    End Function

    Public Function ByRefAs(ByVal strName As String, ByVal strType As String) As String
        Select Case CodeType
            Case "vb"
                str = "ByRef " & strName & " As " & ConvertType(strType)
            Case "cs"
                str = GetRef(ConvertType(strType) & " " & strName)
        End Select
        Return str
    End Function

    Public Function DimAs(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "        Dim " & strName & " As " & ConvertType(strType) & vbCrLf
            Case "cs"
                str = "            " & ConvertType(strType) & " " & strName & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function DimAsValue(ByVal strName As String, ByVal strType As String, ByVal strValue As String) As String

        Select Case CodeType
            Case "vb"
                str = "        Dim " & strName & " As " & ConvertType(strType) & " = " & ConvertValue(strValue) & vbCrLf
            Case "cs"
                str = "            " & ConvertType(strType) & " " & strName & " = " & ConvertValue(strValue) & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function DimAsNew(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "        Dim " & strName & " As New " & ConvertType(strType) & vbCrLf
            Case "cs"
                str = "            " & ConvertType(strType) & " " & strName & " = new " & ConvertType(strType) & "();" & vbCrLf
        End Select

        Return str

    End Function

    Public Function DimAsNewList(ByVal strName As String, ByVal strType As String) As String

        Select Case CodeType
            Case "vb"
                str = "        Dim " & strName & " As IList(Of " & strType & "Info) = New List(Of " & strType & "Info)" & vbCrLf
            Case "cs"
                str = "            IList<" & strType & "Info> " & strName & " = new List<" & strType & "Info>();" & vbCrLf
        End Select

        Return str

    End Function

    Public Function SqlAppend(ByVal strSQL As String) As String

        Select Case CodeType
            Case "vb"
                str = "        sqlStatement.Append(""" & strSQL & """)" & vbCrLf
            Case "cs"
                str = "            sqlStatement.Append(""" & strSQL & """);" & vbCrLf
        End Select

        Return str

    End Function

    Public Function ToStringCS(ByVal str As String) As String
        If str = "String" Then
            Return ""
        Else
            Return ".ToString()"
        End If
    End Function

    Public Function DbAddInParameter(ByVal strName As String, ByVal strType As String, ByVal strValue As String) As String
        Select Case CodeType
            Case "vb"
                Select Case strType
                    Case "DateTime"
                        str = "        db.AddInParameter(dbCommand, ""@" & strName & """, DbType.DateTime, IIf(" & strValue & ".Year = 1, DBNull.Value, " & strValue & "))" & vbCrLf
                    Case "Byte[]"
                        str = "        db.AddInParameter(dbCommand, ""@" & strName & """, DbType.Binary, " & strValue & ")" & vbCrLf
                    Case Else
                        str = "        db.AddInParameter(dbCommand, ""@" & strName & """, DbType." & strType & ", " & strValue & ")" & vbCrLf
                End Select
            Case "cs"
                Select Case strType
                    Case "DateTime"
                        str = "            if (" & strValue & ".Year != 1) { db.AddInParameter(dbCommand, ""@" & strName & """, DbType.DateTime, " & strValue & "); } else { db.AddInParameter(dbCommand, ""@" & strName & """, DbType.DateTime, DBNull.Value); }" & vbCrLf
                    Case "Byte[]"
                        str = "            db.AddInParameter(dbCommand, ""@" & strName & """, DbType.Binary, " & strValue & ");" & vbCrLf
                    Case Else
                        str = "            db.AddInParameter(dbCommand, ""@" & strName & """, DbType." & strType & ", " & strValue & ");" & vbCrLf
                End Select
        End Select
        Return str
    End Function

    Public Function DbAddOutParameter(ByVal strNewID As String) As String
        Select Case CodeType
            Case "vb"
                str = "        db.AddOutParameter(dbCommand, ""@" & strNewID & """, DbType.Int32, 5)" & vbCrLf
            Case "cs"
                str = "            db.AddOutParameter(dbCommand, ""@" & strNewID & """, DbType.Int32, 5);" & vbCrLf
        End Select
        Return str
    End Function

    Public Function DbGetParameterValue(ByVal strValue As String, ByVal strNewID As String) As String
        Select Case CodeType
            Case "vb"
                str = "            " & strValue & " = db.GetParameterValue(dbCommand, """ & strNewID & """)" & vbCrLf
            Case "cs"
                str = "                " & strValue & " = int.Parse(db.GetParameterValue(dbCommand, """ & strNewID & """).ToString());" & vbCrLf
        End Select
        Return str
    End Function

    Public Function DbGetAutoIncrementKey(ByVal strValue As String) As String
        Select Case CodeType
            Case "vb"
                str = "            " & strValue & " = db.ExecuteScalar(dbCommand)" & vbCrLf
            Case "cs"
                str = "                " & strValue & " =  int.Parse(db.ExecuteScalar(dbCommand).ToString());" & vbCrLf
        End Select
        Return str
    End Function

    Public Function ReturnValue(ByVal strValue As String) As String

        Select Case CodeType
            Case "vb"
                str = "        Return " & strValue & vbCrLf
            Case "cs"
                str = "            return " & strValue & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function GetDB()

        Select Case CodeType
            Case "vb"
                If Trim(strConnName) <> "" Then
                    str = "        Dim db As Database = DatabaseFactory.CreateDatabase(""" & strConnName & """)" & vbCrLf
                Else
                    str = "        Dim db As Database = DatabaseFactory.CreateDatabase()" & vbCrLf
                End If
            Case "cs"
                If Trim(strConnName) <> "" Then
                    str = "            Database db = DatabaseFactory.CreateDatabase(""" & strConnName & """);" & vbCrLf
                Else
                    str = "            Database db = DatabaseFactory.CreateDatabase();" & vbCrLf
                End If
        End Select

        Return str

    End Function

    Public Function GetDbCommand()

        Select Case CodeType
            Case "vb"
                str = "        Dim dbCommand As DbCommand = db.GetSqlStringCommand(sqlStatement.ToString())" & vbCrLf
            Case "cs"
                str = "            DbCommand dbCommand = db.GetSqlStringCommand(sqlStatement.ToString());" & vbCrLf
        End Select

        Return str

    End Function

    Public Function GetTry() As String

        Select Case CodeType
            Case "vb"
                str = "        Try" & vbCrLf
            Case "cs"
                str = "            try" & vbCrLf
                str &= "            {" & vbCrLf
        End Select

        Return str

    End Function

    Public Function ChkDBInfo() As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("        If myList.Count > 0 Then" & vbCrLf)
                strSET.Append("            myInfo = myList(0)" & vbCrLf)
                strSET.Append("        End If" & vbCrLf)
            Case "cs"
                strSET.Append("            if (myList.Count > 0) {" & vbCrLf)
                strSET.Append("                myInfo = myList[0];" & vbCrLf)
                strSET.Append("            }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function GetListInfo() As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("            Using dataReader As IDataReader = db.ExecuteReader(dbCommand)" & vbCrLf)
                strSET.Append("                While dataReader.Read()" & vbCrLf)
                strSET.Append("                    myList.Add(BindInfo(dataReader))" & vbCrLf)
                strSET.Append("                End While" & vbCrLf)
                strSET.Append("            End Using" & vbCrLf)
            Case "cs"
                strSET.Append("                using (IDataReader dataReader = db.ExecuteReader(dbCommand))" & vbCrLf)
                strSET.Append("                {" & vbCrLf)
                strSET.Append("                    while (dataReader.Read())" & vbCrLf)
                strSET.Append("                    {" & vbCrLf)
                strSET.Append("                        myList.Add(BindInfo(dataReader));" & vbCrLf)
                strSET.Append("                    }" & vbCrLf)
                strSET.Append("                }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function GetIlist(ByVal strName As String) As String

        Select Case CodeType
            Case "vb"
                str = "IList(Of " & strName & "Info)"
            Case "cs"
                str = "IList<" & strName & "Info>"
        End Select

        Return str

    End Function

    Public Function GetList(ByVal strName As String) As String

        Select Case CodeType
            Case "vb"
                str = "List(Of " & strName & "Info)"
            Case "cs"
                str = "List<" & strName & "Info>"
        End Select

        Return str

    End Function

    Public Function GetLine(Optional ByVal strType As String = "") As String

        If strType = "vb" And CodeType = "cs" Then
            Return ""
        Else
            Return vbCrLf
        End If

    End Function

    Public Function GetWebMethod(ByVal summary As String) As String

        Select Case CodeType
            Case "vb"
                str = String.Format("    <WebMethod(Description:=""{0}"")> _", summary) & vbCrLf
                'str = "    <WebMethod()> _" & vbCrLf
            Case "cs"
                str = String.Format("        [WebMethod(Description = ""{0}"")]", summary) & vbCrLf
                'str = "        [WebMethod]" & vbCrLf
        End Select

        Return str

    End Function

    Public Function GetRef(ByVal strValue As String) As String
        'Return "ref " & strValue
        Return strValue
    End Function

    Public Function GetRefParam(ByVal strValue As String) As String

        Dim str2 As String = ""

        Select Case CodeType
            Case "vb"
                str2 = strValue
            Case "cs"
                str2 = GetRef(strValue)
        End Select

        Return str2

    End Function

    Public Function SetImports(ByVal strName As String) As String

        Select Case CodeType
            Case "vb"
                str = "Imports " & strName & vbCrLf
            Case "cs"
                str = "using " & strName & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function SetProperty(ByVal strName As String, ByVal strType As String) As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("    Public Property " & strName & "() As " & ConvertType(strType) & vbCrLf)
                strSET.Append("        Get" & vbCrLf)
                strSET.Append("            Return m_" & strName & vbCrLf)
                strSET.Append("        End Get" & vbCrLf)
                strSET.Append("        Set(ByVal value As " & ConvertType(strType) & ")" & vbCrLf)
                strSET.Append("            m_" & strName & " = value" & vbCrLf)
                strSET.Append("        End Set" & vbCrLf)
                strSET.Append("    End Property" & vbCrLf)
            Case "cs"
                'strSET.Append("        public " & ConvertType(strType) & " " & strName & "{" & vbCrLf)
                'strSET.Append("            get { return m_" & strName & ";}" & vbCrLf)
                'strSET.Append("            set { m_" & strName & " = value;}" & vbCrLf)
                'strSET.Append("        }" & vbCrLf)
                strSET.Append("        public " & ConvertType(strType) & " " & strName & "{ get; set; }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function SetCatch(ByVal strFun As String, Optional ByVal strEx As String = "DbException") As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("        Catch ex As " & strEx & vbCrLf)
                strSET.Append("            LogController.WriteLog(""" & strFun & """, ex.ToString())" & vbCrLf)
                strSET.Append("            Throw (ex)" & vbCrLf)
                strSET.Append("        End Try" & vbCrLf)
                strSET.Append(vbCrLf)
            Case "cs"
                strSET.Append("            }" & vbCrLf)
                strSET.Append("            catch (" & strEx & " ex)" & vbCrLf)
                strSET.Append("            {" & vbCrLf)
                strSET.Append("                LogController.WriteLog(""" & strFun & """, ex.ToString());" & vbCrLf)
                strSET.Append("                throw (ex);" & vbCrLf)
                strSET.Append("            }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function SetIndex(ByVal intIndex As Integer) As String

        Select Case CodeType
            Case "vb"
                str = "(" & intIndex.ToString & ")"
            Case "cs"
                str = "[" & intIndex.ToString & "]"
        End Select

        Return str

    End Function

    Public Function SetString(ByVal Statement As String, ByVal strRight As String) As String

        Select Case CodeType
            Case "vb"
                str = strRight & Statement & vbCrLf
            Case "cs"
                strRight &= "    "
                str = strRight & Statement & ";" & vbCrLf
        End Select

        Return str

    End Function

    Public Function SetInfoColumns(ByVal strName As String, ByVal strType As String) As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("            If Not dr(""" & strName & """).Equals(DBNull.Value) Then" & vbCrLf)
                strSET.Append("                myInfo." & ColumnProperty(strName) & " = dr(""" & strName & """)" & vbCrLf)
                strSET.Append("            End If" & vbCrLf)
            Case "cs"
                strSET.Append("                if (!dr[""" & strName & """].Equals(DBNull.Value))" & vbCrLf)
                strSET.Append("                {" & vbCrLf)
                strSET.Append("                    myInfo." & ColumnProperty(strName) & " = " & ParseType("dr[""" & strName & """]", strType) & ";" & vbCrLf)
                strSET.Append("                }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function SetCrossInfoColumns(ByVal strName As String, ByVal strType As String) As String

        Dim strSET As New StringBuilder

        Select Case CodeType
            Case "vb"
                strSET.Append("            If Not dr(""" & strName & """).Equals(DBNull.Value) Then" & vbCrLf)
                strSET.Append("                myInfo." & GetObjectName(strName) & " = dr(""" & strName & """)" & vbCrLf)
                strSET.Append("            End If" & vbCrLf)
            Case "cs"
                strSET.Append("                if (!dr[""" & strName & """].Equals(DBNull.Value))" & vbCrLf)
                strSET.Append("                {" & vbCrLf)
                strSET.Append("                    myInfo." & GetObjectName(strName) & " = " & ParseType("dr[""" & strName & """]", strType) & ";" & vbCrLf)
                strSET.Append("                }" & vbCrLf)
        End Select

        Return strSET.ToString

    End Function

    Public Function SetClassSummary(ByVal Summary As String) As String

        Dim strSET As New StringBuilder

        If Trim(Summary) <> "" Then
            Select Case CodeType
                Case "vb"
                    strSET.Append("''' <summary>" & vbCrLf)
                    strSET.Append("''' " & Summary & vbCrLf)
                    strSET.Append("''' </summary>" & vbCrLf)
                Case "cs"
                    strSET.Append("    /// <summary>" & vbCrLf)
                    strSET.Append("    /// " & Summary & vbCrLf)
                    strSET.Append("    /// </summary>" & vbCrLf)
            End Select
            Return strSET.ToString
        Else
            Return ""
        End If

    End Function

    Public Function SetFunSummary(ByVal Summary As String) As String

        Summary = Summary.Replace(Chr(10), "")

        If Trim(Summary) <> "" Then

            Dim strSET As New StringBuilder

            Dim aryStr As Array = Trim(Summary).Split(GetLine())

            Dim leftStr As String = ""

            Select Case CodeType
                Case "vb"
                    leftStr = "    ''' "
                Case "cs"
                    leftStr = "        /// "
            End Select

            strSET.Append(leftStr & "<summary>" & vbCrLf)
            If aryStr.Length > 1 Then
                Dim i As Integer
                For i = 0 To aryStr.Length - 1
                    If Trim(aryStr(i)) <> "" Then
                        strSET.Append(leftStr & aryStr(i) & vbCrLf)
                    End If
                Next
            Else
                strSET.Append(leftStr & Summary & vbCrLf)
            End If
            strSET.Append(leftStr & "</summary>" & vbCrLf)

            Return strSET.ToString
        Else
            Return ""
        End If

        'Dim strSET As New StringBuilder

        'If Trim(Summary) <> "" Then
        '    Select Case CodeType
        '        Case "vb"
        '            strSET.Append("    ''' <summary>" & vbCrLf)
        '            strSET.Append("    ''' " & Summary & vbCrLf)
        '            strSET.Append("    ''' </summary>" & vbCrLf)
        '        Case "cs"
        '            strSET.Append("        /// <summary>" & vbCrLf)
        '            strSET.Append("        /// " & Summary & vbCrLf)
        '            strSET.Append("        /// </summary>" & vbCrLf)
        '    End Select
        '    Return strSET.ToString
        'Else
        '    Return ""
        'End If

    End Function

    Public Function SetParamSummary(ByVal ParamName As String, ByVal Summary As String) As String


        Summary = Summary.Replace(Chr(10), "")

        If Trim(Summary) <> "" Then

            Dim strSET As New StringBuilder

            Dim aryStr As Array = Trim(Summary).Split(GetLine())

            Dim leftStr As String = ""

            Select Case CodeType
                Case "vb"
                    leftStr = "    ''' "
                Case "cs"
                    leftStr = "        /// "
            End Select

            strSET.Append(leftStr & "<param name=""" & ParamName & """>" & vbCrLf)
            If aryStr.Length > 1 Then
                Dim i As Integer
                For i = 0 To aryStr.Length - 1
                    If Trim(aryStr(i)) <> "" Then
                        strSET.Append(leftStr & aryStr(i) & vbCrLf)
                    End If
                Next
            Else
                strSET.Append(leftStr & Summary & vbCrLf)
            End If
            strSET.Append(leftStr & "</param>" & vbCrLf)

            Return strSET.ToString
        Else
            Return ""
        End If


        'Dim strSET As New StringBuilder

        'If Trim(Summary) <> "" Then
        '    Select Case CodeType
        '        Case "vb"
        '            strSET.Append("    ''' <param name=""" & ParamName & """>" & Summary & "</param>" & vbCrLf)
        '        Case "cs"
        '            strSET.Append("        /// <param name=""" & ParamName & """>" & Summary & "</param>" & vbCrLf)
        '    End Select
        '    Return strSET.ToString
        'Else
        '    Return ""
        'End If

    End Function

    Public Function SetReturnSummary(ByVal Summary As String) As String

        Dim strSET As New StringBuilder

        If Trim(Summary) <> "" Then
            Select Case CodeType
                Case "vb"
                    strSET.Append("    ''' <returns>" & Summary & "</returns>" & vbCrLf)
                Case "cs"
                    strSET.Append("        /// <returns>" & Summary & "</returns>" & vbCrLf)
            End Select
            Return strSET.ToString
        Else
            Return ""
        End If

    End Function

    Public Function SetFunRemarks(ByVal Remark As String) As String

        Remark = Remark.Replace(Chr(10), "")

        If Trim(Remark) <> "" Then

            Dim strSET As New StringBuilder

            Dim aryStr As Array = Trim(Remark).Split(GetLine())

            Dim leftStr As String = ""

            Select Case CodeType
                Case "vb"
                    leftStr = "    ''' "
                Case "cs"
                    leftStr = "        /// "
            End Select

            strSET.Append(leftStr & "<remarks>" & vbCrLf)
            If aryStr.Length > 1 Then
                Dim i As Integer
                For i = 0 To aryStr.Length - 1
                    If Trim(aryStr(i)) <> "" Then
                        strSET.Append(leftStr & aryStr(i) & vbCrLf)
                    End If
                Next
            Else
                strSET.Append(leftStr & Remark & vbCrLf)
            End If
            strSET.Append(leftStr & "</remarks>" & vbCrLf)

            Return strSET.ToString
        Else
            Return ""
        End If

    End Function

    Public Function ConvertType(ByVal strType As String) As String

        Dim str2 As String = ""

        Select Case CodeType
            Case "vb"
                Select Case strType
                    Case "Int32", "Int64"
                        str2 = "Integer"
                    Case Else
                        str2 = strType
                End Select
            Case "cs"
                Select Case strType
                    Case "Int32", "Int64", "Integer"
                        str2 = "int"
                    Case "String"
                        str2 = "string"
                    Case "Boolean"
                        str2 = "bool"
                    Case "Decimal"
                        str2 = "decimal"
                    Case "Double"
                        str2 = "double"
                    Case Else
                        str2 = strType
                End Select
        End Select

        Return str2

    End Function

    Public Function ConvertValue(ByVal strValue As String) As String
        Dim str2 As String = ""
        Select Case CodeType
            Case "vb"
                str2 = strValue
            Case "cs"
                Select Case strValue
                    Case "False"
                        str2 = "false"
                    Case "True"
                        str2 = "true"
                    Case Else
                        str2 = strValue
                End Select
        End Select
        Return str2
    End Function

    Public Function GetCType(ByVal strValue As String, ByVal strType As String) As String

        Dim str2 As String = ""

        Select Case CodeType
            Case "vb"
                str2 = "CType(" & strValue & ", " & strType & ")"
            Case "cs"
                str2 = "(" & strType & ")" & strValue
        End Select

        Return str2

    End Function

    Public Function ParseType(ByVal strName As String, ByVal strType As String, Optional ByVal ToString As String = ".ToString()") As String
        Dim str2 As String = ""
        Select Case strType
            Case "Int32", "Int64", "Integer"
                str2 = "int.Parse(" & strName & ToString & ")"
            Case "Int16"
                str2 = "Int16.Parse(" & strName & ToString & ")"
            Case "String"
                str2 = strName & ToString
            Case "Boolean"
                str2 = "bool.Parse(" & strName & ToString & ")"
            Case "Decimal"
                str2 = "decimal.Parse(" & strName & ToString & ")"
            Case "Double"
                str2 = "double.Parse(" & strName & ToString & ")"
            Case "Single"
                str2 = "Single.Parse(" & strName & ToString & ")"
            Case "DateTime"
                'str2 = "DateTime.Parse(" & strName & ToString & ")"
                str2 = "(DateTime)" & strName & "" '20131128 changed by bruce
            Case "Byte"
                str2 = "Byte.Parse(" & strName & ToString & ")"
            Case "Byte[]"
                str2 = "(Byte[])" & strName
        End Select
        Return str2
    End Function
End Module
