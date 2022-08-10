Imports System.IO
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Module Module1
    Public sql As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\db.accdb;"
    Public group_name As String = "SSINFO"
    Public rec_no As String = Nothing
    Public rep_id As String = Nothing
    Public rec_edit As Boolean = False
    Public rec_del As Boolean
    Public rec_count As String = Nothing
    Public acc_type As String = Nothing
    Public edit_bill As String = Nothing
    Public account_year As String = "2017"
    Public get_information As Boolean = False
End Module
