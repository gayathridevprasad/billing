Imports System
Imports System.Data.OleDb

Public Class Product_Master
    Dim con As OleDbConnection = New OleDbConnection(sql)
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As DataSet = Nothing
    Dim dt As DataTable
    Dim search_count As String = Nothing
    Dim error_type As String = Nothing
    Dim image_path As String = Nothing
    Dim row_index As Integer

    Private Sub load_stock()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            If cmb_FindGroup.SelectedIndex <> -1 And cmb_FindName.SelectedIndex = -1 Then
                cmd = New OleDbCommand("select Item_Group as GROUP_NAME,HSN_Code as HSN_CODE,Item_Name as PRODUCT_NAME,Item_Rot as ROT,Item_Rate_Out as SELL_RATE,OP_Stock as MY_STOCK from Tbl_StockItem where Item_Group='" & cmb_FindGroup.Text & "'", con)
            ElseIf cmb_FindGroup.SelectedIndex = -1 And cmb_FindName.SelectedIndex <> -1 Then
                cmd = New OleDbCommand("select Item_Group as GROUP_NAME,HSN_Code as HSN_CODE,Item_Name as PRODUCT_NAME,Item_Rot as ROT,Item_Rate_Out as SELL_RATE,OP_Stock as MY_STOCK from Tbl_StockItem where Item_Name='" & cmb_FindName.Text & "'", con)
            Else
                cmd = New OleDbCommand("select Item_Group as GROUP_NAME,HSN_Code as HSN_CODE,Item_Name as PRODUCT_NAME,Item_Rot as ROT,Item_Rate_Out as SELL_RATE,OP_Stock as MY_STOCK from Tbl_StockItem", con)
            End If

            Dim myDA As New OleDbDataAdapter
            myDA.SelectCommand = cmd
            Dim myDatatable As New DataTable
            myDA.Fill(myDatatable)
            DataGridView1.DataSource = myDatatable
            con.Close()
            'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Me.Close()
    End Sub

    Private Sub load_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            'load stock group
            cmd = New OleDbCommand("select distinct Item_Group from Tbl_StockItem", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_group.Items.Clear()
                cmb_FindGroup.Items.Clear()
                While dr.Read()
                    cmb_group.Items.Add(dr("Item_Group"))
                    cmb_FindGroup.Items.Add(dr("Item_Group"))
                End While
            End If

            'load item name
            cmd = New OleDbCommand("select distinct Item_Name from Tbl_StockItem", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_name.Items.Clear()
                cmb_FindName.Items.Clear()
                While dr.Read()
                    cmb_name.Items.Add(dr("Item_Name"))
                    cmb_FindName.Items.Add(dr("Item_Name"))
                End While
            End If

            'load hsn
            cmd = New OleDbCommand("select distinct HSN_Code from Tbl_StockItem", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_hsn.Items.Clear()
                While dr.Read()
                    cmb_hsn.Items.Add(dr("HSN_Code"))
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
            cmd = New OleDbCommand("select COUNT(Item_Name) from Tbl_StockItem", con)
            rec_count = cmd.ExecuteScalar().ToString
            If rec_count = "" Then
                rec_count = "0"
            End If
            rec_count = CInt(rec_count)
            Txt_TotProducts.Text = rec_count
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ConnectionState.Open Then
            con.Close()
        End If
        If rec_edit = False Then
            check_duplicate()
        End If
        If search_count = 0 Then
            con.Open()
            cmd = New OleDbCommand("insert into Tbl_StockItem values(@item_group,@item_brand,@hsn,@item_name,@item_desc,@item_unit," _
                                   & "@rate_in,@rate_out,@Item_rot,@item_stock,@item_path,@group_name)", con)
        End If
        cmd.Parameters.AddWithValue("@item_group", cmb_group.Text)
        cmd.Parameters.AddWithValue("@item_brand", "")
        cmd.Parameters.AddWithValue("@hsn", cmb_hsn.Text)
        cmd.Parameters.AddWithValue("@item_name", (cmb_name.Text))
        cmd.Parameters.AddWithValue("@item_desc", "")
        cmd.Parameters.AddWithValue("@item_unit", "Nos")
        cmd.Parameters.AddWithValue("@rate_in", vbEmpty)
        cmd.Parameters.AddWithValue("@rate_out", CSng(Txt_Rate_Out.Text))
        cmd.Parameters.AddWithValue("@item_rot", CSng(cmb_Tax.Text))
        cmd.Parameters.AddWithValue("@item_stock", vbEmpty)
        cmd.Parameters.AddWithValue("@item_path", "")
        cmd.Parameters.AddWithValue("@group_name", group_name)
        cmd.ExecuteNonQuery()
        con.Close()
        MsgBox("Stock Item Saved Successfully!", MsgBoxStyle.Information, "SS INFO")
        clear_entry()
        load_details()
        load_id()
        load_stock()
        rec_edit = False
    End Sub

    Private Sub clear_entry()
        Dim c As Control
        For Each c In Me.Controls
            If TypeOf (c) Is TextBox Then
                c.Text = ""
            End If
        Next
        cmb_name.Text = ""
        cmb_group.Text = ""
        cmb_hsn.Text = ""
        cmb_Tax.Text = ""
        cmb_group.Focus()
    End Sub

    Private Sub cmb_FindGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_FindGroup.SelectedIndexChanged
        If cmb_FindGroup.SelectedIndex <> -1 Then
            load_stock()
        End If
    End Sub

    Private Sub cmb_FindName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_FindName.SelectedIndexChanged
        If cmb_FindName.SelectedIndex <> -1 Then
            load_stock()
        End If
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        load_stock()
    End Sub

    Private Sub character_checking()
        'If (Char.IsControl(e.KeyChar) = False) Then
        '    If (Char.IsLetter(e.KeyChar)) Or (Char.IsWhiteSpace(e.KeyChar)) Then
        '        'do nothing
        '    Else
        '        e.Handled = True
        '        MsgBox("Sorry Only Character & Spaces Allowed!!", MsgBoxStyle.Information, "Verify")
        '        cmb_name.Focus()
        '    End If
        'End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Txt_Rate_Out.KeyPress
        If (Char.IsControl(e.KeyChar) = False) Then
            If (Char.IsDigit(e.KeyChar)) Then
                'do nothing
            Else
                e.Handled = True
                MsgBox("Sorry Only Digits Allowed!!", MsgBoxStyle.Information, "Verify")
                Txt_Rate_Out.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_Rate_Out.Leave
        If Txt_Rate_Out.Text = "" Then
            Txt_Rate_Out.Text = "0"
        End If
        Txt_Rate_Out.Text = Strings.FormatNumber(Txt_Rate_Out.Text, 2, TriState.UseDefault, TriState.UseDefault)
    End Sub

    Private Sub check_duplicate()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim str As String = (cmb_group.Text)
            'str = str.Replace(" ", "")
            con.Open()
            cmd = New OleDbCommand("select COUNT(Item_Name) from Tbl_StockItem where Item_Group='" & cmb_group.Text & "' AND Item_Name='" & cmb_name.Text & "'", con)
            search_count = cmd.ExecuteScalar().ToString
            If search_count = "" Then
                search_count = "0"
            End If
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        row_index = e.RowIndex
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            ' load search stock item
            cmd = New OleDbCommand("select * from Tbl_StockItem where Item_Group=@item_group and Item_Name=@item_name", con)
            cmd.Parameters.AddWithValue("@item_group", DataGridView1.Rows(row_index).Cells(0).Value.ToString)
            cmd.Parameters.AddWithValue("@item_name", DataGridView1.Rows(row_index).Cells(2).Value.ToString)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    cmb_group.Text = dr("Item_Group")
                    cmb_name.Text = dr("Item_Name").ToString
                    cmb_hsn.Text = dr("HSN_Code").ToString
                    cmb_Tax.Text = dr("Item_Rot").ToString
                    'Txt_Rate_In.Text = dr("Item_Rate_In").ToString
                    Txt_Rate_Out.Text = dr("Item_Rate_Out").ToString
                    'Txt_Stock.Text = dr("OP_Stock").ToString
                End While
            End If
            Button1.Enabled = False
            cmb_name.Enabled = False
            con.Close()
            rec_edit = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            If rec_edit = True Then
                con.Open()
                cmd = New OleDbCommand("update Tbl_StockItem set Item_Group='" & cmb_group.Text & "',HSN_Code='" & cmb_hsn.Text & "',Item_Rate_Out=" & CSng(Txt_Rate_Out.Text) & ",Item_Rot=" & CSng(cmb_Tax.Text) & " where Item_Name='" & cmb_name.Text & "'", con)
                cmd.ExecuteNonQuery()
                con.Close()
                MsgBox("Product Details Updated!", MsgBoxStyle.Information, "SS INFO")
                clear_entry()
                load_stock()
                load_id()
                load_details()
                rec_edit = False
                Button1.Enabled = True
                cmb_name.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub cmb_group_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_group.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmb_name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_name.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub Product_Master_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        PictureBox3.Image = My.Resources.add
        PictureBox4.Image = My.Resources.update
        PictureBox5.Image = My.Resources.delete
        PictureBox6.Image = My.Resources.close_btn
        PictureBox7.Image = My.Resources.search
        load_stock()
        load_id()
        load_details()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim msg_result As String = Nothing
            msg_result = MsgBox("Sure want to delted this product?", vbYesNo, "SS INFO")
            If msg_result = vbYes Then
                If rec_edit = True Then
                    con.Open()
                    cmd = New OleDbCommand("delete from Tbl_StockItem where Item_Name='" & cmb_name.Text & "'", con)
                    cmd.ExecuteNonQuery()
                    'cmd = New OleDbCommand("update Tbl_Stock set Item_Group='" & cmb_group.Text & "',HSN_Code='" & cmb_hsn.Text & "',Stock_OB=" & CInt(Txt_Stock.Text) & ",Stock_CB=" & CInt(Txt_Stock.Text) & " where Item_Name='" & cmb_name.Text & "' and Stock_Date=#01/01/" & account_year & "#", con)
                    'cmd.ExecuteNonQuery()
                    con.Close()
                    MsgBox("Product Details Deleted", MsgBoxStyle.Information, "SS INFO")
                    clear_entry()
                    load_stock()
                    load_id()
                    load_details()
                    rec_edit = False
                    Button1.Enabled = True
                    cmb_name.Enabled = True
                End If
            ElseIf msg_result = vbNo Then
                MsgBox("Delete Operation Cancelled!", MsgBoxStyle.Information, "SS INFO")
                clear_entry()
                load_stock()
                load_id()
                load_details()
                rec_edit = False
                Button1.Enabled = True
                cmb_name.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

End Class