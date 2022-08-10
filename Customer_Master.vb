Imports System
Imports System.Data.OleDb
Public Class Customer_Master
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

    Private Sub Customer_Master_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        'PictureBox1.Image = My.Resources.NoImage
        'PictureBox2.Image = My.Resources.Add_File
        PictureBox3.Image = My.Resources.add
        PictureBox4.Image = My.Resources.update
        PictureBox5.Image = My.Resources.delete
        PictureBox6.Image = My.Resources.close_btn
        PictureBox7.Image = My.Resources.search
        'ComboBox1.Text = "DR"
        load_customer()
        load_id()
        load_details()
    End Sub

    Private Sub load_customer()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            If cmb_FindPlace.SelectedIndex <> -1 And cmb_FindName.SelectedIndex = -1 Then
                cmd = New OleDbCommand("select Ledger_Name as CUSTOMER_NAME,Ledger_Place as PLACE,Ledger_Contact as CONTACT_NO,Ledger_GST as GSTIN from Tbl_Ledgers where Ledger_Place='" & cmb_FindPlace.Text & "'", con)
            ElseIf cmb_FindPlace.SelectedIndex = -1 And cmb_FindName.SelectedIndex <> -1 Then
                cmd = New OleDbCommand("select Ledger_Name as CUSTOMER_NAME,Ledger_Place as PLACE,Ledger_Contact as CONTACT_NO,Ledger_GST as GSTIN from Tbl_Ledgers where Ledger_Name='" & cmb_FindName.Text & "'", con)
            Else
                cmd = New OleDbCommand("select Ledger_Name as CUSTOMER_NAME,Ledger_Place as PLACE,Ledger_Contact as CONTACT_NO,Ledger_GST as GSTIN from Tbl_Ledgers where Ledger_Group='Sundry Debtors'", con)
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

    Private Sub load_details()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            'load ledger place
            cmd = New OleDbCommand("select distinct Ledger_Place from Tbl_Ledgers where Ledger_Group='Sundry Debtors'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_Place.Items.Clear()
                cmb_FindPlace.Items.Clear()
                While dr.Read()
                    cmb_Place.Items.Add(dr("Ledger_Place"))
                    cmb_FindPlace.Items.Add(dr("Ledger_Place"))
                End While
            End If

            'load ledger name
            cmd = New OleDbCommand("select distinct Ledger_Name from Tbl_Ledgers where Ledger_Group='Sundry Debtors'", con)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                cmb_name.Items.Clear()
                cmb_FindName.Items.Clear()
                While dr.Read()
                    cmb_name.Items.Add(dr("Ledger_Name"))
                    cmb_FindName.Items.Add(dr("Ledger_Name"))
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
            cmd = New OleDbCommand("select COUNT(Ledger_Name) from Tbl_Ledgers where Ledger_Group='Sundry Debtors'", con)
            rec_count = cmd.ExecuteScalar().ToString
            If rec_count = "" Then
                rec_count = "0"
            End If
            rec_count = CInt(rec_count)
            Txt_TotSuppliers.Text = rec_count
            con.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Me.Close()
    End Sub

    Private Sub cmb_name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_name.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmb_Place_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmb_Place.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub Txt_GSTIN_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Txt_GSTIN.Leave
        If Txt_GSTIN.Text <> "" Then
            If Txt_GSTIN.TextLength < 15 Or Txt_GSTIN.TextLength > 15 Then
                MsgBox("Check Your GSTIN!", MsgBoxStyle.Information, "SS INFO")
                Txt_GSTIN.BackColor = Color.Yellow
                Txt_GSTIN.Focus()
            Else
                Txt_GSTIN.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            If rec_edit = False Then
                check_duplicate()
                If search_count > 0 Then
                    MsgBox("Customer already in this name!", MsgBoxStyle.Information, "SS INFO")
                    Exit Sub
                End If
            End If
            con.Open()
            cmd = New OleDbCommand("select MAX(Ledger_Code) from Tbl_Ledgers", con)
            rec_count = cmd.ExecuteScalar().ToString
            If rec_count = "" Then
                rec_count = "0"
            End If
            cmd = New OleDbCommand("insert into Tbl_Ledgers values(@row_index,@ledger_name,@ledger_group,@primary_name,@nature_group,@ledger_date,@ob,@dc,@address,@place,@pincode,@contact,@tin,@cst,@gst,@state,@email,@aadhar,@sname,saddress,splace,spincode,scontact,@group_name)", con)
            cmd.Parameters.AddWithValue("@row_index", CInt(rec_count) + 1)
            cmd.Parameters.AddWithValue("@ledger_name", cmb_name.Text)
            cmd.Parameters.AddWithValue("@ledger_group", "Sundry Debtors")
            cmd.Parameters.AddWithValue("@primary_name", "Current Assets")
            cmd.Parameters.AddWithValue("@nature_group", "Assets")
            cmd.Parameters.AddWithValue("@ledger_date", "01/01/" & account_year)
            cmd.Parameters.AddWithValue("@ob", vbEmpty)
            cmd.Parameters.AddWithValue("@dc", "")
            cmd.Parameters.AddWithValue("@address", Txt_Address.Text)
            cmd.Parameters.AddWithValue("@place", cmb_Place.Text)
            cmd.Parameters.AddWithValue("@pincode", Txt_Pincode.Text)
            cmd.Parameters.AddWithValue("@contact", Txt_Contact.Text)
            cmd.Parameters.AddWithValue("@tin", "")
            cmd.Parameters.AddWithValue("@cst", "")
            cmd.Parameters.AddWithValue("@gst", Txt_GSTIN.Text)
            cmd.Parameters.AddWithValue("@state", Txt_State.Text)
            cmd.Parameters.AddWithValue("@email", Txt_Email.Text)
            cmd.Parameters.AddWithValue("@aadhar", "")
            cmd.Parameters.AddWithValue("@sname", cmb_name.Text)
            cmd.Parameters.AddWithValue("@saddress", Txt_Address.Text)
            cmd.Parameters.AddWithValue("@splace", cmb_Place.Text)
            cmd.Parameters.AddWithValue("@spincode", Txt_Pincode.Text)
            cmd.Parameters.AddWithValue("@scontact", Txt_Contact.Text)
            cmd.Parameters.AddWithValue("@group_name", group_name)
            cmd.ExecuteNonQuery()
            MsgBox("Customer details Saved!", MsgBoxStyle.Information, "SS INFO")
            con.Close()
            'add_account()
            clear_entry()
            load_customer()
            load_id()
            cmb_name.Focus()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub add_account()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim find_group As String = Nothing
            con.Open()
            If rec_edit = False Then
                cmd = New OleDbCommand("insert into Tbl_Accounts values(@acc_billno,@acc_date,@acc_type,@acc_main,@acc_group,@acc_nature,@acc_name,@acc_place,@acc_details,@acc_dr,@acc_cr,@acc_status,@group_name)", con)
            End If
            cmd.Parameters.AddWithValue("@acc_billno", "NEW")
            cmd.Parameters.AddWithValue("@acc_date", "01/01/" & account_year)
            cmd.Parameters.AddWithValue("@acc_type", "NEW LEDGER")
            cmd.Parameters.AddWithValue("@acc_main", "Sundry Debtors")
            cmd.Parameters.AddWithValue("@acc_group", "Current Assets")
            cmd.Parameters.AddWithValue("@acc_nature", "Assets")
            cmd.Parameters.AddWithValue("@acc_name", cmb_name.Text)
            cmd.Parameters.AddWithValue("@acc_place", cmb_Place.Text)
            cmd.Parameters.AddWithValue("@acc_details", "NEW LEDGER")
            'If ComboBox1.Text = "DR" Then
            'cmd.Parameters.AddWithValue("@acc_dr", CSng(Txt_OB.Text))
            'cmd.Parameters.AddWithValue("@acc_cr", vbEmpty)
            'ElseIf ComboBox1.Text = "CR" Then
            'cmd.Parameters.AddWithValue("@acc_dr", vbEmpty)
            'cmd.Parameters.AddWithValue("@acc_cr", CSng(Txt_OB.Text))
            'End If
            cmd.Parameters.AddWithValue("@acc_status", "0")
            cmd.Parameters.AddWithValue("@group_name", group_name)
            cmd.ExecuteNonQuery()
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
        cmb_name.Text = ""
        cmb_Place.Text = ""
        cmb_name.Focus()
    End Sub

    Private Sub check_duplicate()
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim str As String = (cmb_name.Text)
            'str = str.Replace(" ", "")
            con.Open()
            cmd = New OleDbCommand("select COUNT(Ledger_Name) from Tbl_Ledgers where Ledger_Name='" & cmb_name.Text & "' and Ledger_Group='Sundry Debtors'", con)
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
            ' load search customer
            cmd = New OleDbCommand("select * from Tbl_Ledgers where Ledger_Name=@ledger_Name", con)
            cmd.Parameters.AddWithValue("@item_group", DataGridView1.Rows(row_index).Cells(0).Value.ToString)
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    cmb_name.Text = dr("Ledger_Name").ToString
                    Txt_Address.Text = dr("Ledger_Address").ToString
                    cmb_Place.Text = dr("Ledger_Place").ToString
                    Txt_Pincode.Text = dr("Ledger_Pincode").ToString
                    Txt_Contact.Text = dr("Ledger_Contact").ToString
                    'Txt_Tin.Text = dr("Ledger_Tin").ToString
                    'Txt_Cst.Text = dr("Ledger_Cst").ToString
                    Txt_GSTIN.Text = dr("Ledger_GST").ToString
                    Txt_Email.Text = dr("Ledger_Email").ToString
                    Txt_State.Text = dr("Ledger_State").ToString
                    'Txt_Aadhar.Text = dr("Ledger_Aadhar").ToString
                    'Txt_OB.Text = dr("Opening_Balance").ToString
                    'If dr("DC").ToString = "DR" Then
                    'ComboBox1.Text = "DR"
                    'ElseIf dr("DC").ToString = "CR" Then
                    'ComboBox1.Text = "CR"
                    'End If
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
                cmd = New OleDbCommand("update Tbl_Ledgers set Ledger_Address='" & Txt_Address.Text & "',Ledger_Place='" & cmb_Place.Text & "',Ledger_Pincode='" & Txt_Pincode.Text & "',Ledger_Contact='" & Txt_Contact.Text & "',Ledger_GST='" & Txt_GSTIN.Text & "',Ledger_State='" & Txt_State.Text & "',Ledger_Email='" & Txt_Email.Text & "' where Ledger_Name='" & cmb_name.Text & "'", con)
                cmd.ExecuteNonQuery()
                'If ComboBox1.Text = "DR" Then
                '    cmd = New OleDbCommand("update Tbl_Accounts set Acc_Place='" & cmb_Place.Text & "',Acc_Dr=" & Txt_OB.Text & ",Acc_Cr=0 where Acc_Name='" & cmb_name.Text & "' and Acc_Date=#01/01/" & account_year & "#", con)
                'ElseIf ComboBox1.Text = "CR" Then
                '    cmd = New OleDbCommand("update Tbl_Accounts set Acc_Place='" & cmb_Place.Text & "',Acc_Cr=" & Txt_OB.Text & ",Acc_Dr=0 where Acc_Name='" & cmb_name.Text & "' and Acc_Date=#01/01/" & account_year & "#", con)
                'End If
                cmd.ExecuteNonQuery()
                con.Close()
                MsgBox("Customer Details Updated!", MsgBoxStyle.Information, "SS INFO")
                clear_entry()
                load_customer()
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If ConnectionState.Open Then
                con.Close()
            End If
            Dim msg_result As String = Nothing
            msg_result = MsgBox("Sure want to delted this customer?", vbYesNo, "SS INFO")
            If msg_result = vbYes Then
                If rec_edit = True Then
                    con.Open()
                    cmd = New OleDbCommand("delete from Tbl_Ledgers where Ledger_Name='" & cmb_name.Text & "'", con)
                    cmd.ExecuteNonQuery()
                    'cmd = New OleDbCommand("delete Tbl_Stock set Item_Group='" & cmb_group.Text & "',HSN_Code='" & cmb_hsn.Text & "',Stock_OB=" & CInt(Txt_Stock.Text) & ",Stock_CB=" & CInt(Txt_Stock.Text) & " where Item_Name='" & cmb_name.Text & "' and Stock_Date=#01/01/" & account_year & "#", con)
                    'cmd.ExecuteNonQuery()
                    con.Close()
                    MsgBox("Customer Details Deleted", MsgBoxStyle.Information, "SS INFO")
                    clear_entry()
                    load_customer()
                    load_id()
                    load_details()
                    rec_edit = False
                    Button1.Enabled = True
                    cmb_name.Enabled = True
                End If
            ElseIf msg_result = vbNo Then
                MsgBox("Delete Operation Cancelled!", MsgBoxStyle.Information, "SS INFO")
                clear_entry()
                load_customer()
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

    Private Sub cmb_FindPlace_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_FindPlace.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_FindPlace.SelectedIndex = -1 Then
                MsgBox("Place Not Found!", MsgBoxStyle.Information, "SS INFO")
            End If
        End If
    End Sub

    Private Sub cmb_FindPlace_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_FindPlace.SelectedIndexChanged
        If cmb_FindPlace.SelectedIndex <> -1 Then
            load_customer()
        End If
    End Sub

    Private Sub cmb_FindName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmb_FindName.KeyDown
        If e.KeyCode = Keys.Enter Then
            If cmb_FindName.SelectedIndex = -1 Then
                MsgBox("Customer Not Found!", MsgBoxStyle.Information, "SS INFO")
            End If
        End If
    End Sub

    Private Sub cmb_FindName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_FindName.SelectedIndexChanged
        If cmb_FindName.SelectedIndex <> -1 Then
            load_customer()
        End If
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        load_customer()
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub
End Class