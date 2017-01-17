Imports System.Text
Module mdCodeCS

    Public Function GetProperty(ByVal strName As String, ByVal inputType As String, ByVal dataType As String) As String
        Dim strSET As New StringBuilder
        strSET.Append("        public " & GetPropertyType(inputType, dataType) & " " & strName & " { get; set; }" & vbCrLf)
        Return strSET.ToString
    End Function

    Function GetPropertyType(ByVal inputType As String, ByVal dataType As String) As String
        Dim str As String = "string"

        Select Case inputType
            Case "Text", "Hidden"
                Return ConvertType(dataType)
            Case "Password", "Email", "DropDownList", "RadioButtonList"
                str = "string"
            Case "Date"
                str = "DateTime?"
            Case "File"
                str = "HttpPostedFileBase"
            Case "CheckBox"
                str = "bool"
        End Select

        Return str

    End Function

    Public Function GetDefInputType(ByVal strType As String) As String
        Dim str2 As String = ""
        Select Case strType
            'Case "Int32", "Int64", "Integer"

            'Case "Int16"

            'Case "String"

            'Case "Decimal"

            'Case "Double"

            'Case "Single"

            'Case "Byte"

            'Case "Byte[]"

            Case "Boolean"
                str2 = "CheckBox"
            Case "DateTime"
                str2 = "Date"
            Case Else
                str2 = "Text"
        End Select
        Return str2
    End Function

    Public Function PartialClassForMVC(ByVal strName As String, Optional ByVal strSummary As String = "") As String
        Dim str As String

        str = "namespace " & strNameSpace & vbCrLf
        str &= "{" & vbCrLf
        str &= SetClassSummary(strSummary)
        str &= "    public partial class " & strName & vbCrLf
        str &= "    {" & vbCrLf

        Return str

    End Function


End Module
