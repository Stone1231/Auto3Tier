Module mdDel

    Public Sub DeleteDb(ByVal intDbId As Integer)

        DbId = intDbId

        Dim i As Integer

        Dim TableDv As DataView = GetTableDv.DefaultView
        For i = 0 To TableDv.Count - 1
            DeleteTb(TableDv(i)("TId"))
        Next

        Dim CrossFunDv As DataView = GetCrossFunDv.DefaultView
        For i = 0 To CrossFunDv.Count - 1
            DeleteCross(CrossFunDv(i)("cFId"))
        Next

        DelDb(intDbId)

        DbId = -1

    End Sub

    Public Sub DeleteTb(ByVal intTId As Integer)

        TId = intTId

        Dim i As Integer

        Dim ColumnDv As DataView = GetColumnDv.DefaultView
        For i = 0 To ColumnDv.Count - 1
            DeleteColumn(ColumnDv(i)("ColId"))
        Next

        Dim FunDv As DataView = GetFunDv.DefaultView
        For i = 0 To FunDv.Count - 1
            DeleteFun(FunDv(i)("FId"))
        Next

        DelTableAddInfoByTId(intTId)

        DelTableAddInfoByAddTable(TableName)

        DelTable(intTId)

        TId = -1

    End Sub

    Public Sub DeleteColumn(ByVal intColId As Integer)

        DelColumn(intColId)

    End Sub

    Public Sub DeleteFun(ByVal intFId As Integer)

        DelFunParamByFId(intFId)
        DelFun(intFId)
       
    End Sub

    Public Sub DeleteCross(ByVal intCFId As Integer)
        DelCrossParamByCFId(intCFId)
        DelCrossFun(intCFId)
    End Sub
End Module
