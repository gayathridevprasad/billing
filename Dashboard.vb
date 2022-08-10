Imports System
Imports System.Data.OleDb
Public Class Dashboard
    Dim con As OleDbConnection = New OleDbConnection(sql)
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As DataSet = Nothing
    Dim dt As DataTable
    Dim find_customer As String = Nothing
    Dim find_product As String = Nothing
    Dim find_sales As String = Nothing
    Dim sales_sum As String = Nothing
    Dim tax_sum As String = Nothing

    Private Sub Dashboard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        load_details()
        PictureBox1.Image = My.Resources.customer1
        PictureBox2.Image = My.Resources.product
        PictureBox3.Image = My.Resources.sales
        PictureBox4.Image = My.Resources.Rupee
        PictureBox6.Image = My.Resources.close_btn
    End Sub

    Private Sub load_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            'load customer()
            cmd = New OleDbCommand("select COUNT(Ledger_Name) from Tbl_Ledgers", con)
            find_customer = cmd.ExecuteScalar().ToString
            If find_customer = "" Then
                find_customer = "0"
            End If
            Label2.Text = find_customer

            'load product()
            cmd = New OleDbCommand("select COUNT(Item_Name) from Tbl_StockItem", con)
            find_product = cmd.ExecuteScalar().ToString
            If find_product = "" Then
                find_product = "0"
            End If
            Label3.Text = find_product

            'load_sales_count
            cmd = New OleDbCommand("select MAX(Bill_No) from Tbl_Sales", con)
            find_sales = cmd.ExecuteScalar().ToString
            If find_sales = "" Then
                find_sales = "0"
            End If
            Label5.Text = find_sales

            'load bill amount
            cmd = New OleDbCommand("select SUM(Bill_Amount) from Tbl_Sales", con)
            sales_sum = cmd.ExecuteScalar().ToString
            If sales_sum = "" Then
                sales_sum = "0"
            End If
            'load tax
            cmd = New OleDbCommand("select SUM(Item_Tax) from Tbl_Sales", con)
            tax_sum = cmd.ExecuteScalar().ToString
            If tax_sum = "" Then
                tax_sum = "0"
            End If

            Label7.Text = CSng(sales_sum) + CSng(tax_sum)

            con.Close()
        Catch ex As Exception

        End Try
    End Sub

    
    Private Sub Panel14_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel14.Paint

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Me.Close()
    End Sub
End Class