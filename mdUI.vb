Imports System.Text
Module mdUI
    Public Function GetViewState(ByVal strName As String, ByVal strType As String, Optional ByVal DefValue As String = "") As String
        Dim strSET As New StringBuilder
        If DefValue = "" Then
            DefValue = GetDefValue(strType)
        End If
        Select Case CodeType
            Case "vb"
                strSET.Append(String.Format("    Private Property {0}() As {1}", strName, ConvertType(strType)) & GetLine())
                strSET.Append("        Get" & GetLine())
                strSET.Append(String.Format("            If IsNothing(ViewState(""{0}"")) = True Then", strName) & GetLine())
                strSET.Append(String.Format("                Return {0}", DefValue) & GetLine())
                strSET.Append("            Else" & GetLine())
                strSET.Append(String.Format("                Return ViewState(""{0}"")", strName) & GetLine())
                strSET.Append("            End If" & GetLine())
                strSET.Append("        End Get" & GetLine())
                strSET.Append(String.Format("        Set(ByVal value As {0})", ConvertType(strType)) & GetLine())
                strSET.Append(String.Format("            ViewState(""{0}"") = value", strName) & GetLine())
                strSET.Append("        End Set" & GetLine())
                strSET.Append("    End Property" & GetLine())
            Case "cs"
                strSET.Append(String.Format("    private {0} {1}", ConvertType(strType), strName) & GetLine())
                strSET.Append("    {" & GetLine())
                strSET.Append("        get" & GetLine())
                strSET.Append("        {" & GetLine())
                strSET.Append(String.Format("            if (ViewState[""{0}""] == null)", strName) & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", DefValue) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("            else" & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", ParseType(String.Format("ViewState[""{0}""]", strName), strType)) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("        }" & GetLine())
                strSET.Append("        set { ")
                strSET.Append(String.Format("ViewState[""{0}""] = value; ", strName))
                strSET.Append("}" & GetLine())
                strSET.Append("    }" & GetLine())
        End Select
        Return strSET.ToString
    End Function

    Public Function GetSession(ByVal strName As String, ByVal strType As String, Optional ByVal DefValue As String = "") As String
        Dim strSET As New StringBuilder
        If DefValue = "" Then
            DefValue = GetDefValue(strType)
        End If
        Select Case CodeType
            Case "vb"
                strSET.Append(String.Format("    Private Property {0}() As {1}", strName, ConvertType(strType)) & GetLine())
                strSET.Append("        Get" & GetLine())
                strSET.Append(String.Format("            If IsNothing(Session(""{0}"")) = True Then", strName) & GetLine())
                strSET.Append(String.Format("                Return {0}", DefValue) & GetLine())
                strSET.Append("            Else" & GetLine())
                strSET.Append(String.Format("                Return Session(""{0}"")", strName) & GetLine())
                strSET.Append("            End If" & GetLine())
                strSET.Append("        End Get" & GetLine())
                strSET.Append(String.Format("        Set(ByVal value As {0})", ConvertType(strType)) & GetLine())
                strSET.Append(String.Format("            Session(""{0}"") = value", strName) & GetLine())
                strSET.Append("        End Set" & GetLine())
                strSET.Append("    End Property" & GetLine())
            Case "cs"
                strSET.Append(String.Format("    private {0} {1}", ConvertType(strType), strName) & GetLine())
                strSET.Append("    {" & GetLine())
                strSET.Append("        get" & GetLine())
                strSET.Append("        {" & GetLine())
                strSET.Append(String.Format("            if (Session[""{0}""] == null)", strName) & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", DefValue) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("            else" & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", ParseType(String.Format("Session[""{0}""]", strName), strType)) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("        }" & GetLine())
                strSET.Append("        set { ")
                strSET.Append(String.Format("Session[""{0}""] = value; ", strName))
                strSET.Append("}" & GetLine())
                strSET.Append("    }" & GetLine())
        End Select
        Return strSET.ToString
    End Function

    Public Function GetReadOnlyProperty(ByVal strName As String, ByVal strType As String, ByVal FromType As String, Optional ByVal DefValue As String = "") As String
        Dim strSET As New StringBuilder
        If DefValue = "" Then
            DefValue = GetDefValue(strType)
        End If
        Dim ValueType, CheckValue, ControlName As String
        Select Case FromType
            Case "TextBox"
                ValueType = "Text"
                CheckValue = "Text.Trim()"
            Case "CheckBox"
                ValueType = "Checked"
                CheckValue = ""
            Case Else
                ValueType = "SelectedValue"
                CheckValue = "SelectedValue"
        End Select

        ControlName = GetControlName(FromType, strName)
        Select Case CodeType
            Case "vb"
                strSET.Append(String.Format("    Private ReadOnly Property {0}() As {1}", strName, ConvertType(strType)) & GetLine())
                strSET.Append("        Get" & GetLine())
                If CheckValue <> "" Then
                    strSET.Append(String.Format("            If {0}.{1} = """" Then", ControlName, CheckValue) & GetLine())
                    strSET.Append(String.Format("                Return {0}", DefValue) & GetLine())
                    strSET.Append("            Else" & GetLine())
                    strSET.Append(String.Format("                Return {0}.{1}", ControlName, ValueType) & GetLine())
                    strSET.Append("            End If" & GetLine())
                Else
                    strSET.Append(String.Format("            Return {0}.{1}", ControlName, ValueType) & GetLine())
                End If
                strSET.Append("        End Get" & GetLine())
                strSET.Append("    End Property" & GetLine())
            Case "cs"
                strSET.Append(String.Format("    private {0} {1}", ConvertType(strType), strName) & GetLine())
                strSET.Append("    {" & GetLine())
                strSET.Append("        get" & GetLine())
                strSET.Append("        {" & GetLine())
                If CheckValue <> "" Then
                    strSET.Append(String.Format("            if ({0}.{1} == """")", ControlName, CheckValue) & GetLine())
                    strSET.Append("            {" & GetLine())
                    strSET.Append(String.Format("                return {0};", DefValue) & GetLine())
                    strSET.Append("            }" & GetLine())
                    strSET.Append("            else" & GetLine())
                    strSET.Append("            {" & GetLine())
                    strSET.Append(String.Format("                return {0};", ParseType(String.Format("{0}.{1}", ControlName, ValueType), strType, "")) & GetLine())
                    strSET.Append("            }" & GetLine())
                Else
                    If strType = "Boolean" Then
                        strSET.Append(String.Format("            return {0}.{1};", ControlName, ValueType) & GetLine())
                    Else
                        strSET.Append(String.Format("            return {0};", ParseType(String.Format("{0}.{1}", ControlName, ValueType), strType)) & GetLine())
                    End If
                End If
                strSET.Append("        }" & GetLine())
                strSET.Append("    }" & GetLine())
        End Select
        Return strSET.ToString
    End Function

    Public Function GetRequest(ByVal strName As String, ByVal strType As String, Optional ByVal DefValue As String = "") As String
        Dim strSET As New StringBuilder
        If DefValue = "" Then
            DefValue = GetDefValue(strType)
        End If
        Select Case CodeType
            Case "vb"
                strSET.Append(String.Format("    Private  ReadOnly Property {0}() As {1}", strName, ConvertType(strType)) & GetLine())
                strSET.Append("        Get" & GetLine())
                strSET.Append(String.Format("            If IsNothing(Request(""{0}"")) = True Then", strName) & GetLine())
                strSET.Append(String.Format("                Return {0}", DefValue) & GetLine())
                strSET.Append("            Else" & GetLine())
                strSET.Append(String.Format("                Return Request(""{0}"")", strName) & GetLine())
                strSET.Append("            End If" & GetLine())
                strSET.Append("        End Get" & GetLine())
                strSET.Append("    End Property" & GetLine())
            Case "cs"
                strSET.Append(String.Format("    private {0} {1}", ConvertType(strType), strName) & GetLine())
                strSET.Append("    {" & GetLine())
                strSET.Append("        get" & GetLine())
                strSET.Append("        {" & GetLine())
                strSET.Append(String.Format("            if (Request[""{0}""] == null)", strName) & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", DefValue) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("            else" & GetLine())
                strSET.Append("            {" & GetLine())
                strSET.Append(String.Format("                return {0};", ParseType(String.Format("Request[""{0}""]", strName), strType)) & GetLine())
                strSET.Append("            }" & GetLine())
                strSET.Append("        }" & GetLine())
                strSET.Append("    }" & GetLine())
        End Select
        Return strSET.ToString
    End Function

    Public Function GetDefValue(ByVal DBsype As String) As String
        Select Case DBsype
            Case "Int16", "Int32", "Int64", "Decimal", "Double"
                Return "0"
            Case "String"
                Return """"""
            Case "DateTime"
                Return "DateTime.Parse(""1900/01/01"")"
            Case "Boolean"
                Return ConvertValue("False")
            Case Else
                Return """"""
        End Select
    End Function

    Public Function GetUIDefValue(ByVal DBsype As String) As String
        Select Case DBsype
            Case "Int16", "Int32", "Int64", "Decimal", "Double"
                Return "0"
            Case "String"
                Return ""
            Case "DateTime"
                Return "1900/01/01"
            Case "Boolean"
                Return mdCode.ConvertValue("False")
            Case Else
                Return ""
        End Select
    End Function

    Public Function GetControlName(ByVal Type As String, ByVal Name As String, Optional ByVal IsQuery As Boolean = True) As String
        Dim ControlName As String
        If IsQuery = True Then
            Type = "qry"
        End If
        Select Case Type
            Case "DropDownList"
                ControlName = String.Format("ddl{0}", Name)
            Case "RadioButtonList"
                ControlName = String.Format("rbl{0}", Name)
            Case "ListBox"
                ControlName = String.Format("lbx{0}", Name)
            Case "CheckBox"
                ControlName = String.Format("cbx{0}", Name)
            Case "qry"
                ControlName = String.Format("qry{0}", Name)
            Case Else
                ControlName = String.Format("txt{0}", Name)
        End Select
        Return ControlName
    End Function

    Public Function GetRegion(ByVal str As String) As String
        Select Case CodeType
            Case "vb"
                Return String.Format("#Region ""{0}""", str) & GetLine()
            Case Else
                Return String.Format("    #region ""{0}""", str) & GetLine()
        End Select
    End Function

    Public Function EndRegion() As String
        Select Case CodeType
            Case "vb"
                Return "#End Region" & GetLine() & GetLine()
            Case Else
                Return "    #endregion" & GetLine() & GetLine()
        End Select
    End Function
End Module
