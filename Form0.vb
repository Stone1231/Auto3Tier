Imports System.Net.NetworkInformation
Public Class Form0

    Private Sub Form0_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CheckNumber() = False Then
            Me.Close()
        End If

        DbId = -1
        DbConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=3TierDB.mdb;"
        'DbConn = "Data Source=localhost;Initial Catalog=3TierDB;Persist Security Info=True;User ID=sa;Password=test"
        BindDbDv()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckData() = True Then
            If CheckBox1.Checked = False Then
                WSNameSpace = ""
            Else
                WSNameSpace = Trim(TextBox5.Text)
            End If

            AddDb(Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text), Trim(TextBox4.Text), WSNameSpace, ComboBox1.SelectedItem.ToString())
            CleanDbInfo()
            BindDbDv()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DbId = -1 Then
            MessageBox.Show("請選擇一個資料庫連線資料")
        Else
            If CheckData() = True Then

                If CheckBox1.Checked = False Then
                    WSNameSpace = ""
                Else
                    WSNameSpace = Trim(TextBox5.Text)
                End If

                UpDateDb(DbId, Trim(TextBox1.Text), Trim(TextBox2.Text), Trim(TextBox3.Text), Trim(TextBox4.Text), WSNameSpace, ComboBox1.SelectedItem.ToString())

                BindDbDv()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Select Case e.ColumnIndex
            Case 2 '載入
                DbId = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                Dim DbInfo As DataView = GetDbInfo(DbId).DefaultView
                strConn = DbInfo(0)("Conn")
                strConnName = DbInfo(0)("ConnName")

                DbType = DbInfo(0)("DbType")

                If String.IsNullOrEmpty(DbInfo(0)("Project")) Then
                    strProj = ""
                Else
                    strProj = DbInfo(0)("Project") 'DbInfo(0)("Project") & "."
                End If

                If IsDBNull(DbInfo(0)("WSNameSpace")) = False Then
                    WSNameSpace = DbInfo(0)("WSNameSpace")
                Else
                    WSNameSpace = ""
                End If

                IsChgDb = True

                Me.Visible = False
                Form2.Visible = True
                Form2.LoadData()

            Case 3 '讀取
                DbId = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                ReadDbInfo()
            Case 4 '刪除                
                mdDel.DeleteDb(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                BindDbDv()
        End Select
    End Sub

    Private Sub BindDbDv()
        DataGridView1.DataSource = GetDbDv()
    End Sub

    Private Sub ReadDbInfo()
        Dim dt As DataTable = GetDbInfo(DbId)

        TextBox1.Text = dt.DefaultView(0)("Name")
        TextBox2.Text = dt.DefaultView(0)("Project")
        TextBox3.Text = dt.DefaultView(0)("Conn")
        TextBox4.Text = dt.DefaultView(0)("ConnName")

        ComboBox1.SelectedItem = dt.DefaultView(0)("DbType")

        If IsDBNull(dt.DefaultView(0)("WSNameSpace")) = True Then
            TextBox5.Text = "http://tempuri.org/"
            CheckBox1.Checked = False
        Else
            If dt.DefaultView(0)("WSNameSpace") <> "" Then
                TextBox5.Text = dt.DefaultView(0)("WSNameSpace")
                CheckBox1.Checked = True
            Else
                TextBox5.Text = "http://tempuri.org/"
                CheckBox1.Checked = False
            End If
        End If
    End Sub

    Private Sub CleanDbInfo()
        DbId = -1
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub

    Private Function CheckData() As Boolean
        Dim bool As Boolean = True
        Dim str As String = ""

        If Trim(TextBox1.Text) = "" Then
            bool = False
            str = "請輸入名稱"
        End If

        If Trim(TextBox3.Text) = "" Then
            bool = False
            str &= Chr(13) & "請輸入連線字串"
        End If

        If ComboBox1.SelectedItem = Nothing Then
            bool = False
            str &= Chr(13) & "請輸入DB類型"
        End If

        If bool = False Then
            MessageBox.Show(str)
        End If

        Return bool
    End Function

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Label5.Visible = CheckBox1.Checked
        TextBox5.Visible = CheckBox1.Checked
    End Sub

    Private Function CheckNumber() As Boolean
        Dim computerProperties As IPGlobalProperties = IPGlobalProperties.GetIPGlobalProperties()
        Dim nics As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        Dim adapter As NetworkInterface
        Dim isCheck As Boolean = False
        For Each adapter In nics
            If adapter.GetPhysicalAddress().ToString() <> "" Then     
                Select Case adapter.GetPhysicalAddress().ToString()
                    Case "005056C00008", "001D7D91C15E", "001731DDE1ED", "000AE4EE3159", "000AE4FEC923", "001A92C9C813", "0011D894D51F", "001A9296DD5A", "000FB08C29E4", "000EA6A113C2", "001731DDE1F0", "001E371ADC5C", "001A92C9C219", "0015F2D62355", "0022431B4E63", "0040D0790678", "00166F28C9D7", "000AE4FC5193", "0023548310B3", "00235492C831", "0022436218BC ", "00235492C3EB", "001731DDE1F0", "00235492C35E", "0015F2D62509", "00248CDF3490", "00235492C35E", "00248CDF347A", "0026180BBD92", "0022434C4589", "0015F2D6234D ", "000E0CB3786C", "00248CDF32EB", "0013D4AD6DDA", "001BFCB2EEB2"
                        isCheck = True
                End Select
            End If
        Next adapter
        'Return isCheck
        Return True
    End Function
End Class
