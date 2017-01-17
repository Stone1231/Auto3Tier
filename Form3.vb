Public Class Form3

    Public Sub LoadData()

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

        Dim dv As DataView = GetTableDv().DefaultView
        Dim i As Integer
        For i = 0 To dv.Count - 1
            Me.ListBox1.Items.Add(dv(i))
        Next

        Dim dv1 As DataView = GetCrossFunDv().DefaultView
        For i = 0 To dv1.Count - 1
            Me.ListBox3.Items.Add(dv1(i))
        Next

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (ListBox2.Items.Count + ListBox4.Items.Count) > 0 Then

            ProgressBar1.Visible = True

            ProgressBar1.Minimum = 0

            ProgressBar1.Maximum = ListBox2.Items.Count + ListBox4.Items.Count

            ProgressBar1.Value = 1

            ProgressBar1.Step = 1

            Dim i As Integer

            If RadioButton1.Checked = True Then
                CodeType = RadioButton1.Tag
            Else
                CodeType = RadioButton2.Tag
            End If

            If ListBox2.Items.Count > 0 Then
                Dim mTId As Integer = TId
                For i = 0 To ListBox2.Items.Count - 1
                    TId = ListBox2.Items(i)("TId")
                    GetTableCode()
                    ProgressBar1.PerformStep()
                Next
                TId = mTId
            End If

            If ListBox4.Items.Count > 0 Then
                Dim mTId As Integer = TId
                TId = -1
                For i = 0 To ListBox4.Items.Count - 1
                    GetCrossCode(ListBox4.Items(i)("cFId"))
                    ProgressBar1.PerformStep()
                Next
                TId = mTId
            End If

            ProgressBar1.Visible = False
            Me.Visible = False

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As Integer
        For i = 0 To ListBox1.SelectedItems.Count - 1
            Me.ListBox2.Items.Add(ListBox1.SelectedItem)
            Me.ListBox1.Items.Remove(ListBox1.SelectedItem)
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim i, j As Integer

        For i = 0 To ListBox2.SelectedItems.Count - 1
            Me.ListBox2.Items.Remove(ListBox2.SelectedItem)
        Next

        ListBox1.Items.Clear()

        Dim dv As DataView = GetTableDv().DefaultView

        For i = 0 To dv.Count - 1
            Me.ListBox1.Items.Add(dv(i))
            For j = 0 To ListBox2.Items.Count - 1
                If ListBox2.Items(j)("Name") = ListBox1.Items(ListBox1.Items.Count - 1)("Name") Then
                    ListBox1.Items.Remove(ListBox1.Items(ListBox1.Items.Count - 1))
                    Exit For
                End If
            Next
        Next

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()

        Dim dv As DataView = GetTableDv().DefaultView
        Dim i As Integer
        For i = 0 To dv.Count - 1
            Me.ListBox2.Items.Add(dv(i))
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()

        Dim dv As DataView = GetTableDv().DefaultView
        Dim i As Integer
        For i = 0 To dv.Count - 1
            Me.ListBox1.Items.Add(dv(i))
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim i As Integer
        For i = 0 To ListBox3.SelectedItems.Count - 1
            Me.ListBox4.Items.Add(ListBox3.SelectedItem)
            Me.ListBox3.Items.Remove(ListBox3.SelectedItem)
        Next
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim i, j As Integer

        For i = 0 To ListBox4.SelectedItems.Count - 1
            Me.ListBox4.Items.Remove(ListBox4.SelectedItem)
        Next

        ListBox3.Items.Clear()

        Dim dv As DataView = GetCrossFunDv().DefaultView

        For i = 0 To dv.Count - 1
            Me.ListBox3.Items.Add(dv(i))
            For j = 0 To ListBox4.Items.Count - 1
                If ListBox4.Items(j)("Name") = ListBox3.Items(ListBox3.Items.Count - 1)("Name") Then
                    ListBox3.Items.Remove(ListBox3.Items(ListBox3.Items.Count - 1))
                    Exit For
                End If
            Next
        Next

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

        Dim i As Integer

        Dim dv As DataView = GetCrossFunDv().DefaultView
        For i = 0 To dv.Count - 1
            Me.ListBox4.Items.Add(dv(i))
        Next

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

        Dim i As Integer

        Dim dv As DataView = GetCrossFunDv().DefaultView
        For i = 0 To dv.Count - 1
            Me.ListBox3.Items.Add(dv(i))
        Next

    End Sub

End Class