Imports System
Imports System.IO
Imports System.Data.OleDb

Public Class Report
    Dim con As OleDbConnection = New OleDbConnection(sql)
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As DataSet = Nothing
    Dim dt As DataTable
    Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject
    Dim select_row As Integer
    Dim find_bill As Integer

    Private Sub Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        PictureBox6.Image = My.Resources.close_btn
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            Dim date_from As String = Nothing
            Dim date_to As String = Nothing
            date_from = DateTimePicker1.Value.Date
            date_to = DateTimePicker2.Value.Date
            cmd = New OleDbCommand("select DISTINCT Bill_No AS BILLNO,Bill_Date AS BILLDATE,Sup_Name AS CUSTOMER,Net_Amount AS BILLAMOUNT from Tbl_Sales where Bill_Date>=@date_from and Bill_Date<=@date_to and Bill_Status=0 and Group_Name='" & group_name & "'", con)
            cmd.Parameters.AddWithValue("@date_from", Convert.ToDateTime(date_from))
            cmd.Parameters.AddWithValue("@date_to", Convert.ToDateTime(date_to))
            Dim MyDA As New OleDbDataAdapter
            MyDA.SelectCommand = cmd
            Dim MyDataTable As New DataTable
            MyDA.Fill(MyDataTable)
            DataGridView1.DataSource = MyDataTable
            con.Close()
            lbl_count.Text = DataGridView1.RowCount
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Bill_No As String = Nothing
        Dim net_amt As Single
        Dim bill_address As String = Nothing

        If ConnectionState.Open Then
            con.Close()
        End If
        If DataGridView1.RowCount = 0 Then
            MsgBox("No Records found!", MsgBoxStyle.Information, "SS INFO")
            Exit Sub
        End If
        con.Open()
        find_bill = DataGridView1.Rows(select_row).Cells(0).Value
        cmd = New OleDbCommand("select * from Tbl_Sales where Bill_No = " & find_bill, con)
        dr = cmd.ExecuteReader()
        If dr.HasRows Then
            While dr.Read()
                Bill_No = dr("Bill_No")
                net_amt = dr("Net_Amount")
                bill_address = dr("Sup_Name").ToString & Chr(13) & dr("Sup_Address").ToString & Chr(13) & dr("Sup_Place").ToString & "-" & dr("Sup_Pincode").ToString & Chr(13) & Chr(13) & "GSTIN :" & dr("Sup_GSTIN").ToString
            End While
        End If
        con.Close()
        con.Open()

        cmd = New OleDbCommand("select * from Tbl_Sales where Bill_No=" & find_bill & " and Bill_Status=0 and Group_Name='" & group_name & "'", con)

        Dim MyDA As New OleDbDataAdapter
        MyDA.SelectCommand = cmd
        Dim MyDataTable As New DataTable
        MyDA.Fill(MyDataTable)
        con.Close()

        view.CrystalReportViewer1.RefreshReport()
        view.Rpt_Sale1.Load("Rpt_Sale.rpt")
        view.Rpt_Sale1.SetDataSource(MyDataTable)
        view.CrystalReportViewer1.ReportSource = view.Rpt_Sale1

        objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_BillDetails")
        objText.Text = bill_address
        objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Amount_In_Words")
        objText.Text = RupeesToWord(Convert.ToString(net_amt))

        view.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If DataGridView1.RowCount = 0 Then
            MsgBox("No Records found!", MsgBoxStyle.Information, "SS INFO")
            Exit Sub
        End If
        edit_bill = DataGridView1.Rows(select_row).Cells(0).Value
        rec_edit = True
        Sales.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim copy As Integer = 0
        view.CrystalReportViewer1.RefreshReport()
        view.Rpt_Sale1.Refresh()

        Dim net_amt As Single = 0
        If TextBox1.Text = "" Or TextBox1.Text = "0" Then
            MsgBox("Copy input invalid!", MsgBoxStyle.Information, "SS INFO")
            Exit Sub
        End If
        If DataGridView1.RowCount = 0 Then
            MsgBox("No Records found!", MsgBoxStyle.Information, "SS INFO")
            Exit Sub
        End If
        load_report()
        For count As Integer = 0 To CInt(TextBox1.Text) - 1
            If count = 0 Then
                objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_Copy")
                objText.Text = "ORIGINAL"
            ElseIf count = 1 Then
                objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_Copy")
                objText.Text = "DUPLICATE"
            ElseIf count = 2 Then
                objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_Copy")
                objText.Text = "TRIPLICATE"
            ElseIf count > 2 Then
                objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_Copy")
                objText.Text = "Transport Copy"
            End If
            view.Rpt_Sale1.PrintToPrinter(1, False, 0, 0)
        Next
    End Sub

    Private Sub load_report()
        Dim Bill_No As String = Nothing
        Dim net_amt As Single
        Dim bill_address As String = Nothing
        Dim location As String = Nothing
        Dim cust_tin As String = Nothing
        Dim company_name As String = Nothing
        Dim comp_address As String = Nothing
        Dim comp_place As String = Nothing
        Dim comp_pincode As String = Nothing
        Dim comp_contact As String = Nothing
        Dim comp_GSTIN As String = Nothing

        If ConnectionState.Open Then
            con.Close()
        End If
        con.Open()

        find_bill = DataGridView1.Rows(select_row).Cells(0).Value

        cmd = New OleDbCommand("select * from Tbl_Sales where Bill_No = " & find_bill & " and Group_Name='" & group_name & "'", con)

        dr = cmd.ExecuteReader()
        If dr.HasRows Then
            While dr.Read()
                Bill_No = dr("Bill_No")
                net_amt = dr("Net_Amount")
                'location = dr("Jurisdiction").ToString
                bill_address = dr("Sup_Name").ToString & Chr(13) & dr("Sup_Address").ToString & Chr(13) & dr("Sup_Place").ToString & "-" & dr("Sup_Pincode").ToString & Chr(13) & cust_tin & Chr(13) & "GSTIN :" & dr("Sup_GSTIN").ToString
            End While
        End If
        con.Close()
        con.Open()

        cmd = New OleDbCommand("select * from Tbl_Sales where Bill_No=" & find_bill & " and Bill_Status=0 and Group_Name='" & group_name & "'", con)

        Dim MyDA As New OleDbDataAdapter
        MyDA.SelectCommand = cmd
        Dim MyDataTable As New DataTable
        MyDA.Fill(MyDataTable)
        con.Close()

        view.CrystalReportViewer1.RefreshReport()
        view.Rpt_Sale1.Load("Rpt_Sale.rpt")
        view.Rpt_Sale1.SetDataSource(MyDataTable)
        view.CrystalReportViewer1.ReportSource = view.Rpt_Sale1

        objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Txt_BillDetails")
        objText.Text = bill_address
        objText = view.Rpt_Sale1.ReportDefinition.ReportObjects.Item("Amount_In_Words")
        objText.Text = RupeesToWord(Convert.ToString(net_amt))

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton1.BackColor = Color.Green
            load_filter()
            cmb_Customer.Text = ""
        Else
            RadioButton1.BackColor = Color.Transparent
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton2.BackColor = Color.Green
            load_filter()
            cmb_Customer.Text = ""
        Else
            RadioButton2.BackColor = Color.Transparent
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            Dim date_from As String = Nothing
            Dim date_to As String = Nothing
            date_from = DateTimePicker3.Value.Date
            date_to = DateTimePicker4.Value.Date

            If RadioButton1.Checked = True Then
                cmd = New OleDbCommand("select DISTINCT Bill_No,Bill_Date,Sup_Name,Net_Amount from Tbl_Sales where Bill_Date>=@date_from and Bill_Date<=@date_to and Sup_Name='" & cmb_Customer.Text & "' and Bill_Status=0 and Group_Name='" & group_name & "'", con)
            ElseIf RadioButton2.Checked = True Then
                cmd = New OleDbCommand("select DISTINCT Bill_No,Item_Name,Item_Qty,Item_Rate from Tbl_Sales where Bill_Date>=@date_from and Bill_Date<=@date_to and Item_Name='" & cmb_Customer.Text & "' and Bill_Status=0 and Group_Name='" & group_name & "'", con)
            Else
                cmd = New OleDbCommand("select DISTINCT Bill_No,Bill_Date,Sup_Name,Net_Amount from Tbl_Sales where Bill_Date>=@date_from and Bill_Date<=@date_to and Bill_Status=0 and Group_Name='" & group_name & "'", con)
            End If
            cmd.Parameters.AddWithValue("@date_from", Convert.ToDateTime(date_from))
            cmd.Parameters.AddWithValue("@date_to", Convert.ToDateTime(date_to))
            Dim MyDA As New OleDbDataAdapter
            MyDA.SelectCommand = cmd
            Dim MyDataTable As New DataTable
            MyDA.Fill(MyDataTable)
            con.Close()
            If RadioButton1.Checked = True Or ((RadioButton1.Checked = False) And (RadioButton2.Checked = False)) Then
                view.CrystalReportViewer1.RefreshReport()
                view.Rpt_Sales_Date1.Load("Rpt_Sales_Date.rpt")
                view.Rpt_Sales_Date1.SetDataSource(MyDataTable)
                view.CrystalReportViewer1.ReportSource = view.Rpt_Sales_Date1

                objText = view.Rpt_Sales_Date1.ReportDefinition.ReportObjects.Item("Txt_From")
                objText.Text = date_from
                objText = view.Rpt_Sales_Date1.ReportDefinition.ReportObjects.Item("Txt_To")
                objText.Text = date_to
            ElseIf RadioButton2.Checked = True Then
                view.CrystalReportViewer1.RefreshReport()
                view.Rpt_Item_Sales1.Load("Rpt_Item_Sales.rpt")
                view.Rpt_Item_Sales1.SetDataSource(MyDataTable)
                view.CrystalReportViewer1.ReportSource = view.Rpt_Item_Sales1

                objText = view.Rpt_Item_Sales1.ReportDefinition.ReportObjects.Item("Txt_From")
                objText.Text = date_from
                objText = view.Rpt_Item_Sales1.ReportDefinition.ReportObjects.Item("Txt_To")
                objText.Text = date_to
            End If
            
            view.Show()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub load_filter()
        con.Open()
        If RadioButton1.Checked = True Then
            cmd = New OleDbCommand("select distinct Sup_Name from Tbl_Sales", con)
        ElseIf RadioButton2.Checked = True Then
            cmd = New OleDbCommand("select distinct Item_Name from Tbl_Sales", con)
        End If
        dr = cmd.ExecuteReader()
        If dr.HasRows Then
            cmb_Customer.Items.Clear()
            While dr.Read()
                If RadioButton1.Checked = True Then
                    cmb_Customer.Items.Add(dr("Sup_Name").ToString)
                ElseIf RadioButton2.Checked = True Then
                    cmb_Customer.Items.Add(dr("Item_Name").ToString)
                End If
            End While
        End If
        con.Close()
    End Sub
End Class