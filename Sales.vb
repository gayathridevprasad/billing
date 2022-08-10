Imports System
Imports System.Data.OleDb
Public Class Sales
    Dim con As OleDbConnection = New OleDbConnection(sql)
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As DataSet = Nothing
    Dim dt As DataTable
    Dim search_count As String = Nothing
    Dim error_type As String = Nothing
    Dim image_path As String = Nothing
    Dim row_index As Integer = -1
    Dim pos As Integer = 0
    Dim sup_name As String = Nothing
    Dim sup_address As String = Nothing
    Dim sup_place As String = Nothing
    Dim sup_pin As String = Nothing
    Dim sup_contact As String = Nothing
    Dim sup_gstin As String = Nothing
    Dim sup_state As String = Nothing
    Dim round_off_sign As String = Nothing
    Dim previous_closing_stock As String = Nothing
    Dim select_count As Integer = 0

    Private Sub Purchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        PictureBox1.Image = My.Resources.Rupee
        'PictureBox2.Image = My.Resources.close_btn
        PictureBox3.Image = My.Resources.add
        PictureBox4.Image = My.Resources.update
        'PictureBox5.Image = My.Resources.delete
        PictureBox6.Image = My.Resources.close_btn

        If rec_edit = False Then
            load_id()
            Txt_Packing.Text = "0.00"
            Txt_Discount.Text = "0.00"
            Txt_Round.Text = "0.00"
            lbl_NetAmount.Text = "0.00"
            cmb_supplier.Text = ""
        End If
        If rec_edit = True Then
            find_bill_details()
            load_supplier_details()
        End If
        load_details()
    End Sub

    Private Sub find_bill_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim tot_bill As Single = 0

            con.Open()
            cmd = New OleDbCommand("select Item_Name,HSN_Code,Item_Qty,Item_Rate,Item_Rate*Item_Qty AS Amount,Item_Rot as ROT,Item_Tax as Tax_Value,Item_Unit as UOM,Tax_5,Tax_12,Tax_18,Tax_28,Tax_Nil from Tbl_Sales where Bill_No = " & CInt(edit_bill) & " and Group_Name='" & group_name & "'", con)
            dr = cmd.ExecuteReader()
            dr.Close()
            Dim ds As New DataSet
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Rows(0).Cells(0).Selected = False
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            con.Close()
            ' load search sales
            con.Open()
            cmd = New OleDbCommand("select * from Tbl_Sales where Bill_No = " & CInt(edit_bill) & " and Group_Name='" & group_name & "'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    Txt_BillNo.Text = dr("Bill_No")
                    DateTimePicker1.Text = dr("Bill_Date").ToString
                    cmb_supplier.Text = dr("Sup_Name").ToString
                    Txt_Packing.Text = dr("Loading_Forwarding").ToString
                    Txt_Discount.Text = dr("Discount").ToString
                    Txt_Round.Text = dr("Round_Off").ToString
                    lbl_NetAmount.Text = dr("Net_Amount").ToString
                    'Txt_0.Text = dr("Tax_Nil").ToString
                    'Txt_5.Text = dr("Tax_5").ToString
                    'Txt_12.Text = dr("Tax_12").ToString
                    'Txt_18.Text = dr("Tax_18").ToString
                    'Txt_28.Text = dr("Tax_28").ToString
                End While
            End If
            Grid_sum()
            net_amount_calculation()
            'tot_bill = CSng(lbl_NetAmount.Text) - (CSng(Txt_Total.Text) + CSng(Txt_0.Text) + CSng(Txt_5.Text) + CSng(Txt_12.Text) + CSng(Txt_18.Text) + CSng(Txt_28.Text) + CSng(Txt_Packing.Text) - CSng(Txt_Discount.Text))
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub load_supplier_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            cmd = New OleDbCommand("select * from Tbl_Ledgers where Ledger_Name = '" & cmb_supplier.Text & "'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    sup_name = dr("Ledger_Name").ToString
                    sup_address = dr("Ledger_Address").ToString
                    sup_place = dr("Ledger_Place").ToString
                    sup_pin = dr("Ledger_Pincode").ToString
                    sup_contact = dr("Ledger_Contact").ToString
                    sup_gstin = dr("Ledger_GST").ToString
                    sup_state = dr("Ledger_State").ToString
                End While
            Else
                'MsgBox("Supplier is not found!", MsgBoxStyle.Information, "SS INFO")
            End If
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Private Sub load_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            'load stock item
            cmd = New OleDbCommand("select distinct Item_Name from Tbl_StockItem", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_product.Items.Clear()
                While dr.Read()
                    cmb_product.Items.Add(dr("Item_Name"))
                End While
            End If
            cmd = New OleDbCommand("select distinct Item_Name from Tbl_Sales", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    cmb_product.Items.Add(dr("Item_Name"))
                End While
            End If


            'load supplier
            cmd = New OleDbCommand("select distinct Ledger_Name from Tbl_Ledgers where Ledger_Group='Sundry Debtors'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_supplier.Items.Clear()
                While dr.Read()
                    cmb_supplier.Items.Add(dr("Ledger_Name"))
                End While
            End If
            cmd = New OleDbCommand("select distinct Sup_Name from Tbl_Sales", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    cmb_supplier.Items.Add(dr("Sup_Name"))
                End While
            End If


            'load supplier
            cmd = New OleDbCommand("select distinct Item_Unit from Tbl_Sales", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_uom.Items.Clear()
                While dr.Read()
                    cmb_uom.Items.Add(dr("Item_Unit").ToString)
                End While
            End If

            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub load_id()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            cmd = New OleDbCommand("select MAX(Bill_No) from Tbl_Sales", con)
            rec_count = cmd.ExecuteScalar().ToString
            If rec_count = "" Then
                rec_count = "0"
            End If
            rec_count = CInt(rec_count)
            Txt_BillNo.Text = CInt(rec_count + 1)
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Me.Close()
    End Sub

    Private Sub cmb_product_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_product.GotFocus
        'cmb_product.DroppedDown = True
    End Sub

    Private Sub cmb_supplier_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_supplier.GotFocus
        'cmb_supplier.DroppedDown = True
    End Sub

    Private Sub cmb_product_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_product.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_product.SelectedIndex <> -1 Then
                find_product_details()
                Txt_Qty.Focus()
            End If
        End If
    End Sub

    Private Sub cmb_product_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_product.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub


    Private Sub cmb_product_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_product.Leave
        If cmb_product.SelectedIndex <> -1 Then
            find_product_details()
        End If
    End Sub

    Private Sub find_product_details()
        Dim prod_stock As String = Nothing
        Dim entered_stock As Integer = 0
        con.Open()
        cmd = New OleDbCommand("select * from Tbl_StockItem where Item_Name='" & cmb_product.Text & "'", con)
        dr = cmd.ExecuteReader()
        If dr.HasRows Then
            While dr.Read()
                Txt_HSN.Text = dr("HSN_Code").ToString
                Txt_Rate.Text = dr("Item_Rate_Out").ToString
                cmb_tax.Text = dr("Item_ROT").ToString
            End While
        End If
        'find stock
        'cmd = New OleDbCommand("select * from Tbl_Stock Order By Stock_Date ASC", con)
        'cmd.ExecuteReader()
        'cmd = New OleDbCommand("select LAST(Stock_CB) from Tbl_Stock where Item_Name='" & cmb_product.Text & "'", con)
        'prod_stock = cmd.ExecuteScalar().ToString
        'If prod_stock = "" Then
        'prod_stock = "0"
        'End If

        'find entered stock
        For i = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Rows(i).Cells(0).Value.ToString = cmb_product.Text Then
                entered_stock -= DataGridView1.Rows(i).Cells(2).Value
            End If
        Next
        'Txt_Stock.Text = prod_stock + entered_stock
        con.Close()

        'If CInt(Txt_Stock.Text) <= 0 Then
        '    Button4.Enabled = False
        '    Button5.Enabled = False
        'Else
        '    Button4.Enabled = True
        '    Button5.Enabled = True
        'End If
    End Sub

    Private Sub cmb_product_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_product.LostFocus
        If cmb_product.SelectedIndex <> -1 Then
            find_product_details()
        End If
    End Sub

    Private Sub cmb_product_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_product.SelectedIndexChanged
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'get_information = True
        'Product_Master.Show()
        'Product_Master.cmb_name.Text = cmb_product.Text
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'get_information = True
        'Supplier_Master.Show()
        'Supplier_Master.cmb_name.Text = cmb_supplier.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Txt_Qty.Text = "0" Or Txt_Qty.Text = "" Then
            MsgBox("Quantity is invalid!", MsgBoxStyle.Information, "SS INFO")
            Exit Sub
        End If
        'If CInt(Txt_Stock.Text) < CInt(Txt_Qty.Text) Then
        '    MsgBox("Enter quantity is invalid!", MsgBoxStyle.Information, "SS INFO")
        '    Exit Sub
        'End If
        If DataGridView1.RowCount = 0 And DataGridView1.ColumnCount = 0 Then
            add_columns()
        End If
        If row_index <> -1 Then
            Dim total As Single
            DataGridView1.Rows(row_index).Cells(0).Value = cmb_product.Text
            DataGridView1.Rows(row_index).Cells(0).Value = cmb_product.Text
            DataGridView1.Rows(row_index).Cells(1).Value = Txt_HSN.Text
            DataGridView1.Rows(row_index).Cells(2).Value = Txt_Qty.Text
            DataGridView1.Rows(row_index).Cells(3).Value = Txt_Rate.Text
            total = CSng(DataGridView1.Rows(row_index).Cells(2).Value) * CSng(DataGridView1.Rows(row_index).Cells(3).Value)
            DataGridView1.Rows(row_index).Cells(4).Value = total
            DataGridView1.Rows(row_index).Cells(5).Value = CSng(cmb_tax.Text)
            DataGridView1.Rows(row_index).Cells(6).Value = total * CSng(cmb_tax.Text) / 100
            DataGridView1.Rows(row_index).Cells(7).Value = cmb_uom.Text
            If (cmb_tax.Text) = "5" Then
                DataGridView1.Rows(row_index).Cells(8).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(row_index).Cells(9).Value = "0"
                DataGridView1.Rows(row_index).Cells(10).Value = "0"
                DataGridView1.Rows(row_index).Cells(11).Value = "0"
                DataGridView1.Rows(row_index).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "12" Then
                DataGridView1.Rows(row_index).Cells(8).Value = "0"
                DataGridView1.Rows(row_index).Cells(9).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(row_index).Cells(10).Value = "0"
                DataGridView1.Rows(row_index).Cells(11).Value = "0"
                DataGridView1.Rows(row_index).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "18" Then
                DataGridView1.Rows(row_index).Cells(8).Value = "0"
                DataGridView1.Rows(row_index).Cells(9).Value = "0"
                DataGridView1.Rows(row_index).Cells(10).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(row_index).Cells(11).Value = "0"
                DataGridView1.Rows(row_index).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "28" Then
                DataGridView1.Rows(row_index).Cells(8).Value = "0"
                DataGridView1.Rows(row_index).Cells(9).Value = "0"
                DataGridView1.Rows(row_index).Cells(10).Value = "0"
                DataGridView1.Rows(row_index).Cells(11).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(row_index).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "0" Then
                DataGridView1.Rows(row_index).Cells(8).Value = "0"
                DataGridView1.Rows(row_index).Cells(9).Value = "0"
                DataGridView1.Rows(row_index).Cells(10).Value = "0"
                DataGridView1.Rows(row_index).Cells(11).Value = "0"
                DataGridView1.Rows(row_index).Cells(12).Value = total * CSng(cmb_tax.Text) / 100
            End If
            DataGridView1.Rows(row_index).Cells(4).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(4).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(5).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(5).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(6).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(6).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(8).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(8).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(9).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(9).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(10).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(10).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(11).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(11).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
            DataGridView1.Rows(row_index).Cells(12).Value = Strings.FormatNumber(DataGridView1.Rows(row_index).Cells(12).Value.ToString, 2, TriState.UseDefault, TriState.UseDefault)
        Else
            add_grid()
        End If
        Grid_sum()
        net_amount_calculation()
        If DataGridView1.RowCount <> ComboBox1.Items.Count Then
            ComboBox1.Items.Add("")
        End If
        Txt_Rate.Clear()
        Txt_Qty.Clear()
        'Txt_Stock.Clear()
        Txt_HSN.Clear()
        'Txt_Warranty.Clear()
        cmb_product.Text = ""
        cmb_product.SelectedIndex = -1
        cmb_uom.Text = ""
        cmb_uom.SelectedIndex = -1
        cmb_tax.Text = ""
        cmb_tax.SelectedIndex = -1
        cmb_product.Focus()
        DataGridView1.Rows(0).Selected = False
        row_index = -1
    End Sub

    Private Sub add_columns()
        Try
            'add column to datagridview
            With Me.DataGridView1
                .Columns.Add("Item_Name", "PRODUCT")
                .Columns.Add("HSN_Code", "HSN/SAC")
                .Columns.Add("Quantity", "QTY")
                .Columns.Add("Rate", "RATE")
                .Columns.Add("Amount", "AMOUNT")
                .Columns.Add("Tax(%)", "TAX(%)")
                .Columns.Add("Tax_Value", "TAX")
                .Columns.Add("UOM", "UOM")
                .Columns.Add("Tax_5", "5% GST")
                .Columns.Add("Tax_12", "12% GST")
                .Columns.Add("Tax_18", "18% GST")
                .Columns.Add("Tax_28", "28% GST")
                .Columns.Add("Tax_0", "0% GST")
                '.Columns.Add("WY", "WY")
                'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub add_grid()
        Try
            Dim count As Integer
            count = DataGridView1.Rows.Count
            If rec_edit = True Then
                Dim dt As DataTable = DirectCast(DataGridView1.DataSource, DataTable)
                dt.Rows.Add()
            ElseIf rec_edit = False Then
                DataGridView1.Rows.Add()
            End If
            Dim total As Single
            DataGridView1.Rows(count).Cells(0).Value = cmb_product.Text
            DataGridView1.Rows(count).Cells(1).Value = Txt_HSN.Text
            DataGridView1.Rows(count).Cells(2).Value = Txt_Qty.Text
            DataGridView1.Rows(count).Cells(3).Value = Txt_Rate.Text
            total = CSng(DataGridView1.Rows(count).Cells(2).Value) * CSng(DataGridView1.Rows(count).Cells(3).Value)
            DataGridView1.Rows(count).Cells(4).Value = total
            DataGridView1.Rows(count).Cells(5).Value = CSng(cmb_tax.Text)
            DataGridView1.Rows(count).Cells(6).Value = total * CSng(cmb_tax.Text) / 100
            DataGridView1.Rows(count).Cells(7).Value = cmb_uom.Text
            If (cmb_tax.Text) = "5" Then
                DataGridView1.Rows(count).Cells(8).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(count).Cells(9).Value = "0"
                DataGridView1.Rows(count).Cells(10).Value = "0"
                DataGridView1.Rows(count).Cells(11).Value = "0"
                DataGridView1.Rows(count).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "12" Then
                DataGridView1.Rows(count).Cells(8).Value = "0"
                DataGridView1.Rows(count).Cells(9).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(count).Cells(10).Value = "0"
                DataGridView1.Rows(count).Cells(11).Value = "0"
                DataGridView1.Rows(count).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "18" Then
                DataGridView1.Rows(count).Cells(8).Value = "0"
                DataGridView1.Rows(count).Cells(9).Value = "0"
                DataGridView1.Rows(count).Cells(10).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(count).Cells(11).Value = "0"
                DataGridView1.Rows(count).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "28" Then
                DataGridView1.Rows(count).Cells(8).Value = "0"
                DataGridView1.Rows(count).Cells(9).Value = "0"
                DataGridView1.Rows(count).Cells(10).Value = "0"
                DataGridView1.Rows(count).Cells(11).Value = total * CSng(cmb_tax.Text) / 100
                DataGridView1.Rows(count).Cells(12).Value = "0"
            ElseIf (cmb_tax.Text) = "0" Then
                DataGridView1.Rows(count).Cells(8).Value = "0"
                DataGridView1.Rows(count).Cells(9).Value = "0"
                DataGridView1.Rows(count).Cells(10).Value = "0"
                DataGridView1.Rows(count).Cells(11).Value = "0"
                DataGridView1.Rows(count).Cells(12).Value = total * CSng(cmb_tax.Text) / 100
            End If
            'DataGridView1.Rows(count).Cells(13).Value = Txt_Warranty.Text
            DataGridView1.Rows(count).Cells(4).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(5).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(6).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(8).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(9).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(10).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(11).Style.Format = "N2"
            DataGridView1.Rows(count).Cells(12).Style.Format = "N2"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Grid_sum()
        Try
            Txt_0.Text = "0"
            Txt_5.Text = "0"
            Txt_12.Text = "0"
            Txt_18.Text = "0"
            Txt_28.Text = "0"
            Txt_Total.Text = "0"

            If DataGridView1.RowCount <> 0 Then
                For index As Integer = 0 To DataGridView1.RowCount - 1
                    Txt_5.Text += CSng(DataGridView1.Rows(index).Cells(8).Value)
                    Txt_12.Text += CSng(DataGridView1.Rows(index).Cells(9).Value)
                    Txt_18.Text += CSng(DataGridView1.Rows(index).Cells(10).Value)
                    Txt_28.Text += CSng(DataGridView1.Rows(index).Cells(11).Value)
                    Txt_0.Text += CSng(DataGridView1.Rows(index).Cells(12).Value)
                Next
                For index As Integer = 0 To DataGridView1.RowCount - 1
                    Txt_Total.Text += CSng(DataGridView1.Rows(index).Cells(4).Value)
                Next
            End If
            Txt_5.Text = Strings.FormatNumber(Txt_5.Text, 2, TriState.UseDefault, TriState.UseDefault)
            Txt_12.Text = Strings.FormatNumber(Txt_12.Text, 2, TriState.UseDefault, TriState.UseDefault)
            Txt_18.Text = Strings.FormatNumber(Txt_18.Text, 2, TriState.UseDefault, TriState.UseDefault)
            Txt_28.Text = Strings.FormatNumber(Txt_28.Text, 2, TriState.UseDefault, TriState.UseDefault)
            Txt_0.Text = Strings.FormatNumber(Txt_0.Text, 2, TriState.UseDefault, TriState.UseDefault)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub net_amount_calculation()
        lbl_NetAmount.Text = "0.00"
        lbl_NetAmount.Text = (CSng(Txt_Total.Text) + CSng(Txt_0.Text) + CSng(Txt_5.Text) + CSng(Txt_12.Text) + CSng(Txt_18.Text) + CSng(Txt_28.Text) + CSng(Txt_Packing.Text)) - CSng(Txt_Discount.Text)
        lbl_NetAmount.Text = Strings.FormatNumber(lbl_NetAmount.Text, 2, TriState.UseDefault, TriState.UseDefault)
    End Sub

    Private Sub cmb_supplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_supplier.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_supplier.SelectedIndex <> -1 Or cmb_supplier.Text <> "" Then
                load_customer_details()
            End If
        End If
    End Sub

    Private Sub cmb_supplier_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_supplier.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmb_tax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_tax.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_tax.SelectedIndex = -1 Or cmb_tax.Text = "" Then
                MsgBox("Invalid Tax!", MsgBoxStyle.Information, "SS INFO")
                cmb_tax.DroppedDown = True
                Exit Sub
            Else
                cmb_uom.Focus()
            End If
        End If
    End Sub

    Private Sub cmb_tax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_tax.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmb_uom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_uom.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_uom.SelectedIndex = -1 Or cmb_uom.Text = "" Then
                cmb_uom.Text = "NOS"
            End If
            Button4.Focus()
        End If
    End Sub

    Private Sub cmb_uom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_uom.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub Txt_Qty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Qty.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Txt_Qty.Text = "" Or Txt_Qty.Text = "0" Then
                MsgBox("Invalid Quantity!", MsgBoxStyle.Information, "SS INFO")
                Txt_Qty.BackColor = Color.Yellow
                Txt_Qty.Focus()
            Else
                Txt_Qty.BackColor = Color.White
                Txt_Rate.Focus()
            End If
        End If
    End Sub

    Private Sub Txt_Rate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Rate.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Txt_Rate.Text = "" Or Txt_Rate.Text = "0" Then
                MsgBox("Invalid Rate!", MsgBoxStyle.Information, "SS INFO")
                Txt_Rate.BackColor = Color.Yellow
                Txt_Rate.Focus()
            Else
                Txt_Rate.BackColor = Color.White
                cmb_tax.Focus()
            End If
        End If
    End Sub

    Private Sub cmb_uom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_uom.Leave
        If cmb_uom.SelectedIndex = -1 Or cmb_uom.Text = "" Then
            cmb_uom.Text = "NOS"
        End If
    End Sub

    Private Sub cmb_tax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_tax.Leave
        If cmb_tax.SelectedIndex = -1 Or cmb_tax.Text = "" Then
            MsgBox("Invalid Tax!", MsgBoxStyle.Information, "SS INFO")
            cmb_tax.DroppedDown = True
            Exit Sub
        Else
            cmb_uom.Focus()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        row_index = e.RowIndex
        'MsgBox(row_index)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If

            Dim auto_no As Integer = 0
            auto_no = CInt(rec_no) + 1

            If rec_edit = True Then
                con.Open()
                cmd = New OleDbCommand("delete from Tbl_Sales where Bill_No=" & Txt_BillNo.Text & "", con)
                cmd.ExecuteNonQuery()
                con.Close()
            End If
            If Txt_BillNo.Text = "" Then
                MsgBox("Bill No invalid!", MsgBoxStyle.Critical, "SS INFO")
                Txt_BillNo.Focus()
                Txt_BillNo.BackColor = Color.Yellow
                Exit Sub
            Else
                Txt_BillNo.BackColor = Color.White
            End If
            con.Open()
            For i = 0 To DataGridView1.Rows.Count - 1
                cmd = New OleDbCommand("insert into Tbl_Sales values(@auto_no,@billno,@billdate,@billtype,@po_no,@po_date,@sup_name,@sup_address,@sup_place,@sup_contact,@sup_pincode,@sub_state,@sub_GSTIN," _
                                       & "@item_name,@HSN_Code,@item_qty,@item_rate,@item_amount,@item_unit,@item_warranty,@item_ROT,@item_tax_value,@Item_Keys," _
                                       & "@Tax_Nil,@Tax_5,@Tax_12,@Tax_18,@Tax_28," _
                                       & "@freight,@packing,@discount,@round_off,@net_amount,@bill_status,@narration,@group_name)", con)
                cmd.Parameters.AddWithValue("@auto_no", auto_no)
                cmd.Parameters.AddWithValue("@billno", Txt_BillNo.Text)
                cmd.Parameters.AddWithValue("@billdate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@billtype", "CREDIT")
                cmd.Parameters.AddWithValue("@po_no", "")
                cmd.Parameters.AddWithValue("@po_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@sup_name", cmb_supplier.Text)
                cmd.Parameters.AddWithValue("@sup_address", sup_address)
                cmd.Parameters.AddWithValue("@sup_place", sup_place)
                cmd.Parameters.AddWithValue("@sup_contact", sup_contact)
                cmd.Parameters.AddWithValue("@sup_pincode", sup_pin)
                cmd.Parameters.AddWithValue("@sup_state", sup_state)
                cmd.Parameters.AddWithValue("@sup_GSTIN", sup_gstin)
                cmd.Parameters.AddWithValue("@item_name", DataGridView1.Rows(i).Cells(0).Value)
                cmd.Parameters.AddWithValue("@HSN_Code", DataGridView1.Rows(i).Cells(1).Value)
                cmd.Parameters.AddWithValue("@item_qty", DataGridView1.Rows(i).Cells(2).Value)
                cmd.Parameters.AddWithValue("@item_rate", DataGridView1.Rows(i).Cells(3).Value)
                cmd.Parameters.AddWithValue("@item_amount", DataGridView1.Rows(i).Cells(4).Value)
                cmd.Parameters.AddWithValue("@item_unit", DataGridView1.Rows(i).Cells(7).Value)
                cmd.Parameters.AddWithValue("@item_warranty", vbEmpty)
                cmd.Parameters.AddWithValue("@item_ROT", DataGridView1.Rows(i).Cells(5).Value)
                cmd.Parameters.AddWithValue("@item_tax_value", DataGridView1.Rows(i).Cells(6).Value)
                cmd.Parameters.AddWithValue("@item_keys", "")
                cmd.Parameters.AddWithValue("@Tax_Nil", DataGridView1.Rows(i).Cells(12).Value)
                cmd.Parameters.AddWithValue("@Tax_5", DataGridView1.Rows(i).Cells(8).Value)
                cmd.Parameters.AddWithValue("@Tax_12", DataGridView1.Rows(i).Cells(9).Value)
                cmd.Parameters.AddWithValue("@Tax_18", DataGridView1.Rows(i).Cells(10).Value)
                cmd.Parameters.AddWithValue("@Tax_28", DataGridView1.Rows(i).Cells(11).Value)
                cmd.Parameters.AddWithValue("@freight", vbEmpty)
                cmd.Parameters.AddWithValue("@packing", CSng(Txt_Packing.Text))
                cmd.Parameters.AddWithValue("@discount", CSng(Txt_Discount.Text))
                cmd.Parameters.AddWithValue("@round_off", CSng(Txt_Round.Text))
                cmd.Parameters.AddWithValue("@net_amount", CSng(lbl_NetAmount.Text))
                cmd.Parameters.AddWithValue("@bill_status", "0")
                cmd.Parameters.AddWithValue("@narration", Txt_Notes.Text)
                cmd.Parameters.AddWithValue("@group_name", group_name)
                cmd.ExecuteNonQuery()
            Next
            con.Close()
            'add_account()
            'update_stock_summary()
            'load_total_sales()
            MsgBox("Sales Saved Successfully!", MsgBoxStyle.Information, "SS INFO")
            rec_edit = False
            clear_entry()
            edit_bill = Nothing
            'Me.Close()
            Purchase_Load(sender, e)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub add_account()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim i As Integer = 0
            con.Open()

            For i = 0 To 1
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                If i = 0 Then
                    cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                    cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@acc_type", "Sales")
                    cmd.Parameters.AddWithValue("@acc_main", "Sundry Debtors")
                    cmd.Parameters.AddWithValue("@acc_group", "Current Assets")
                    cmd.Parameters.AddWithValue("@acc_nature", "Assets")
                    cmd.Parameters.AddWithValue("@acc_name", sup_name)
                    cmd.Parameters.AddWithValue("@acc_place", sup_place)
                    cmd.Parameters.AddWithValue("@acc_details", sup_name)
                    cmd.Parameters.AddWithValue("@acc_dr", CSng(lbl_NetAmount.Text))
                    cmd.Parameters.AddWithValue("@acc_cr", vbEmpty)
                    cmd.Parameters.AddWithValue("@acc_status", "0")
                    cmd.Parameters.AddWithValue("@group", group_name)
                ElseIf i = 1 Then
                    cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                    cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@acc_type", "Sales")
                    cmd.Parameters.AddWithValue("@acc_main", "Sales Accounts")
                    cmd.Parameters.AddWithValue("@acc_group", "Sales Accounts")
                    cmd.Parameters.AddWithValue("@acc_nature", "Income")
                    cmd.Parameters.AddWithValue("@acc_name", "Sales A/c")
                    cmd.Parameters.AddWithValue("@acc_place", sup_place)
                    cmd.Parameters.AddWithValue("@acc_details", "Sales A/c")
                    cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                    cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_Total.Text))
                    cmd.Parameters.AddWithValue("@acc_status", "0")
                    cmd.Parameters.AddWithValue("@group", group_name)
                End If
                cmd.ExecuteNonQuery()
            Next

            'check 5% tax
            If CInt(Txt_5.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Duties & Taxes")
                cmd.Parameters.AddWithValue("@acc_group", "Current Liabilities")
                cmd.Parameters.AddWithValue("@acc_nature", "Liabilities")
                cmd.Parameters.AddWithValue("@acc_name", "GST@5%")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "GST@5%")
                cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_5.Text))
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If
            'check 12% tax
            If CInt(Txt_12.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Duties & Taxes")
                cmd.Parameters.AddWithValue("@acc_group", "Current Liabilities")
                cmd.Parameters.AddWithValue("@acc_nature", "Liabilities")
                cmd.Parameters.AddWithValue("@acc_name", "GST@12%")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "GST@12%")
                cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_12.Text))
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If

            'check 18% tax
            If CInt(Txt_18.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Duties & Taxes")
                cmd.Parameters.AddWithValue("@acc_group", "Current Liabilities")
                cmd.Parameters.AddWithValue("@acc_nature", "Liabilities")
                cmd.Parameters.AddWithValue("@acc_name", "GST@18%")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "GST@18%")
                cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_18.Text))
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If
            'IGST 28% tax
            If CInt(Txt_28.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Duties & Taxes")
                cmd.Parameters.AddWithValue("@acc_group", "Current Liabilities")
                cmd.Parameters.AddWithValue("@acc_nature", "Liabilities")
                cmd.Parameters.AddWithValue("@acc_name", "GST@28%")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "GST@28%")
                cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_28.Text))
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If
            ''check freight
            'If CInt(Txt_Frieght.Text) > 0 Then
            '    cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
            '    cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
            '    cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
            '    cmd.Parameters.AddWithValue("@acc_type", "Purchase")
            '    cmd.Parameters.AddWithValue("@acc_main", "Direct Expenses")
            '    cmd.Parameters.AddWithValue("@acc_group", "Direct Expenses")
            '    cmd.Parameters.AddWithValue("@acc_nature", "Expenses")
            '    cmd.Parameters.AddWithValue("@acc_name", "Freight Charges")
            '    cmd.Parameters.AddWithValue("@acc_place", "")
            '    cmd.Parameters.AddWithValue("@acc_details", "Freight Charges")
            '    cmd.Parameters.AddWithValue("@acc_dr", CSng(Txt_Frieght.Text))
            '    cmd.Parameters.AddWithValue("@acc_cr", vbEmpty)
            '    cmd.Parameters.AddWithValue("@acc_status", "0")
            '    cmd.Parameters.AddWithValue("@group", group_name)
            '    cmd.ExecuteNonQuery()
            'End If

            'check Packing
            If CInt(Txt_Packing.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_group", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_nature", "Income")
                cmd.Parameters.AddWithValue("@acc_name", "Loading Unloading Expense")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "Loading Unloading Expense")
                cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_Packing.Text))
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If

            'check discount
            If CInt(Txt_Discount.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_group", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_nature", "Income")
                cmd.Parameters.AddWithValue("@acc_name", "Sales Discount")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "Sales Discount")
                cmd.Parameters.AddWithValue("@acc_dr", CSng(Txt_Discount.Text))
                cmd.Parameters.AddWithValue("@acc_cr", vbEmpty)
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If

            'check Round
            If CSng(Txt_Round.Text) > 0 Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
                cmd.Parameters.AddWithValue("@acc_billno", (Txt_BillNo.Text))
                cmd.Parameters.AddWithValue("@acc_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@acc_type", "Sales")
                cmd.Parameters.AddWithValue("@acc_main", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_group", "Direct Incomes")
                cmd.Parameters.AddWithValue("@acc_nature", "Income")
                cmd.Parameters.AddWithValue("@acc_name", "Sales Round-Off")
                cmd.Parameters.AddWithValue("@acc_place", "")
                cmd.Parameters.AddWithValue("@acc_details", "Sales Round-Off")
                If round_off_sign = "plus" Then
                    cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
                    cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_Round.Text))
                ElseIf round_off_sign = "minus" Then
                    cmd.Parameters.AddWithValue("@acc_dr", CSng(Txt_Round.Text))
                    cmd.Parameters.AddWithValue("@acc_cr", vbEmpty)
                End If
                cmd.Parameters.AddWithValue("@acc_status", "0")
                cmd.Parameters.AddWithValue("@group", group_name)
                cmd.ExecuteNonQuery()
            End If
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub update_stock_summary()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim s As String = DateTime.Now.ToString("MM/dd/yyyy")
            Dim previous_day As String = Date.Today.AddDays(-1).ToString("MM/dd/yyyy")
            Dim previous_purchase_qty As String = Nothing
            Dim find_item_brand As String = Nothing
            Dim find_item_group As String = Nothing
            Dim find_item_unit As String = Nothing
            Dim rec_count As String = Nothing

            If rec_edit = True Then
            End If

            con.Open()
            For i = 0 To DataGridView1.Rows.Count - 1
                ' -- get last closing stock for each item
                cmd = New OleDbCommand("select Item_Brand from Tbl_StockItem where Item_Name='" & DataGridView1.Rows(i).Cells(0).Value.ToString & "'", con)
                find_item_brand = cmd.ExecuteScalar().ToString
                If find_item_brand = "" Then
                    find_item_brand = ""
                End If

                cmd = New OleDbCommand("select Item_Group from Tbl_StockItem where Item_Name='" & DataGridView1.Rows(i).Cells(0).Value.ToString & "'", con)
                find_item_group = cmd.ExecuteScalar().ToString
                If find_item_group = "" Then
                    find_item_group = ""
                End If

                cmd = New OleDbCommand("select LAST(Stock_CB) from Tbl_Stock where Item_Name='" & DataGridView1.Rows(i).Cells(0).Value.ToString & "'", con)
                previous_closing_stock = cmd.ExecuteScalar().ToString
                If previous_closing_stock = "" Then
                    previous_closing_stock = "0"
                End If
                cmd = New OleDbCommand("insert into Tbl_Stock values(@stock_date,@item_group,@item_brand,@hsn_code,@item_name,@unit_name,@stock_ob,@p_qty,@tot_qty,@s_qty,@stock_cb,@rate,@amount,@billno,@billtype,@group_name)", con)
                'Else
                'cmd = New OleDbCommand("update Tbl_Stock set Stock_Date=@stock_date,Item_Group=@item_group,Item_Code=@code,Item_Name=@item_name,Printer_Name=@print_name,Stock_Unit=@unit_name,Stock_OB=@stock_ob,P_Qty=@p_qty,Tot_Qty=@tot_qty,S_Qty=@s_qty,Stock_CB=@stock_cb,Stock_Rate=@rate,Stock_Amount=@amount Where Item_Code='" & DataGridView1.Rows(i).Cells(0).Value.ToString & "' and Stock_Date=#" & DateTimePicker1.Value.Date & "#", con)
                'End If
                cmd.Parameters.AddWithValue("@stock_date", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@item_group", find_item_group)
                cmd.Parameters.AddWithValue("@item_brand", find_item_brand)
                cmd.Parameters.AddWithValue("@hsn_code", DataGridView1.Rows(i).Cells(1).Value)
                cmd.Parameters.AddWithValue("@item_name", DataGridView1.Rows(i).Cells(0).Value)
                cmd.Parameters.AddWithValue("@unit_name", DataGridView1.Rows(i).Cells(7).Value)
                cmd.Parameters.AddWithValue("@stock_ob", CInt(previous_closing_stock))
                cmd.Parameters.AddWithValue("@p_qty", "0")
                cmd.Parameters.AddWithValue("@tot_qty", (CInt(previous_closing_stock)))
                cmd.Parameters.AddWithValue("@s_qty", DataGridView1.Rows(i).Cells(2).Value)
                cmd.Parameters.AddWithValue("@stock_cb", (CInt(previous_closing_stock) - DataGridView1.Rows(i).Cells(2).Value))
                cmd.Parameters.AddWithValue("@rate", DataGridView1.Rows(i).Cells(3).Value)
                cmd.Parameters.AddWithValue("@amount", DataGridView1.Rows(i).Cells(3).Value * (CInt(previous_closing_stock) - DataGridView1.Rows(i).Cells(2).Value))
                cmd.Parameters.AddWithValue("@billno", Txt_BillNo.Text)
                cmd.Parameters.AddWithValue("@billtype", "Sales")
                cmd.Parameters.AddWithValue("@group_name", group_name)
                cmd.ExecuteNonQuery()
            Next
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub clear_entry()
        Dim c As Control
        For Each c In Me.Controls
            If TypeOf (c) Is TextBox Then
                c.Text = ""
            End If
        Next
        DataGridView1.DataSource = Nothing
        cmb_product.SelectedIndex = -1
        cmb_supplier.SelectedIndex = -1
        'lbl_customer_details.Text = ""
        lbl_NetAmount.Text = ""
        Txt_0.Clear()
        Txt_5.Clear()
        Txt_12.Clear()
        Txt_18.Clear()
        Txt_28.Clear()
        Txt_Total.Clear()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        load_id()
        'load_total_sales()
    End Sub

    Private Sub load_total_sales()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim sum_sales As String = Nothing
            con.Open()
            cmd = New OleDbCommand("select SUM(Acc_Dr) from Tbl_Accounts where Acc_Type='Sales' and Acc_Main='Sundry Debtors'", con)
            sum_sales = cmd.ExecuteScalar().ToString
            If sum_sales = "" Then
                sum_sales = "0"
            End If
            'lbl_NetSales.Text = sum_sales
            'lbl_NetSales.Text = Strings.FormatNumber(lbl_NetSales.Text, 2, TriState.UseDefault, TriState.UseDefault)
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cmb_supplier_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_supplier.LostFocus
        If cmb_supplier.SelectedIndex <> -1 Or cmb_supplier.Text <> "" Then
            load_customer_details()
        End If
    End Sub

    Private Sub load_customer_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim bill_amt As String = Nothing
            Dim paid_amt As String = Nothing
            con.Open()
            cmd = New OleDbCommand("select * from Tbl_Ledgers where Ledger_Name = '" & cmb_supplier.Text & "'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    sup_name = dr("Ledger_Name").ToString
                    sup_address = dr("Ledger_Address").ToString
                    sup_place = dr("Ledger_Place").ToString
                    sup_pin = dr("Ledger_Pincode").ToString
                    sup_contact = dr("Ledger_Contact").ToString
                    sup_gstin = dr("Ledger_GST").ToString
                    sup_state = dr("Ledger_State").ToString
                End While
            Else
                sup_name = cmb_supplier.Text
                sup_address = ""
                sup_place = ""
                sup_pin = ""
                sup_contact = ""
                sup_gstin = ""
                sup_state = ""
            End If

            'cmd = New OleDbCommand("select sum(Acc_Cr) from Tbl_Accounts where Acc_Name='" & cmb_supplier.Text & "'", con)
            'bill_amt = cmd.ExecuteScalar().ToString
            'If bill_amt = "" Then
            ' bill_amt = "0"
            ' End If
            ' cmd = New OleDbCommand("select sum(Acc_Dr) from Tbl_Accounts where Acc_Name='" & cmb_supplier.Text & "'", con)
            ' paid_amt = cmd.ExecuteScalar().ToString
            ' If paid_amt = "" Then
            ' paid_amt = "0"
            'End If
            'Txt_Inv_Value.Text = bill_amt
            'Txt_Paid_Value.Text = paid_amt
            'Txt_Balance.Text = CSng(bill_amt) - CSng(paid_amt)
            'Txt_Inv_Value.Text = Strings.FormatNumber(Txt_Inv_Value.Text, 2, TriState.UseDefault, TriState.UseDefault)
            'Txt_Paid_Value.Text = Strings.FormatNumber(Txt_Paid_Value.Text, 2, TriState.UseDefault, TriState.UseDefault)
            'Txt_Balance.Text = Strings.FormatNumber(Txt_Balance.Text, 2, TriState.UseDefault, TriState.UseDefault)
            con.Close()
            'lbl_customer_details.Text = sup_address & Chr(13) & sup_place & "-" & sup_pin & ",  Contact:" & sup_contact & Chr(13) & sup_gstin
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cmb_supplier_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_supplier.SelectedIndexChanged

    End Sub

    Private Sub Txt_Packing_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Packing.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Txt_Packing.Text = "" Then
                Txt_Packing.Text = "0"
            End If
            Txt_Packing.Text = Strings.FormatNumber(Txt_Packing.Text, 2, TriState.UseDefault, TriState.UseDefault)
            net_amount_calculation()
        End If
    End Sub

    Private Sub Txt_Packing_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Packing.LostFocus
        If Txt_Packing.Text = "" Then
            Txt_Packing.Text = "0"
        End If
        Txt_Packing.Text = Strings.FormatNumber(Txt_Packing.Text, 2, TriState.UseDefault, TriState.UseDefault)
        net_amount_calculation()
    End Sub

    Private Sub Txt_Discount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Discount.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Txt_Discount.Text = "" Then
                Txt_Discount.Text = "0"
            End If
            Txt_Discount.Text = Strings.FormatNumber(Txt_Discount.Text, 2, TriState.UseDefault, TriState.UseDefault)
            net_amount_calculation()
        End If
    End Sub

    Private Sub Txt_Discount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Discount.LostFocus
        If Txt_Discount.Text = "" Then
            Txt_Discount.Text = "0"
        End If
        Txt_Discount.Text = Strings.FormatNumber(Txt_Discount.Text, 2, TriState.UseDefault, TriState.UseDefault)
        net_amount_calculation()
    End Sub

    Private Sub Txt_Round_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Txt_Round.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Txt_Round.Text = "" Then
                Txt_Round.Text = "0"
            End If
            Txt_Round.Text = Strings.FormatNumber(Txt_Round.Text, 2, TriState.UseDefault, TriState.UseDefault)
            net_amount_calculation()
        End If
    End Sub

    Private Sub Txt_Round_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Round.LostFocus
        If Txt_Round.Text = "" Then
            Txt_Round.Text = "0"
        End If
        Txt_Round.Text = Strings.FormatNumber(Txt_Round.Text, 2, TriState.UseDefault, TriState.UseDefault)
        net_amount_calculation()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If row_index = -1 Then
            MsgBox("Select Row to delete!", MsgBoxStyle.Information, "SS INFO")
        Else
            DataGridView1.Rows.RemoveAt(row_index)
            'ComboBox1.Items.RemoveAt(row_index)
            'ComboBox3.Items.RemoveAt(row_index)
            DataGridView1.Refresh()
        End If
        row_index = -1
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If row_index = -1 Then
            MsgBox("Select Row to Modify!", MsgBoxStyle.Information, "SS INFO")
        Else
            cmb_product.Text = DataGridView1.Rows(row_index).Cells(0).Value.ToString
            Txt_HSN.Text = DataGridView1.Rows(row_index).Cells(1).Value.ToString
            Txt_Qty.Text = DataGridView1.Rows(row_index).Cells(2).Value.ToString
            Txt_Rate.Text = DataGridView1.Rows(row_index).Cells(3).Value.ToString
            cmb_tax.Text = CInt(DataGridView1.Rows(row_index).Cells(5).Value)
            cmb_uom.Text = DataGridView1.Rows(row_index).Cells(7).Value.ToString
            'Txt_Warranty.Text = DataGridView1.Rows(row_index).Cells(13).Value.ToString
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        round_off_sign = "plus"
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        round_off_sign = "minus"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1_Click(sender, e)
    End Sub
End Class