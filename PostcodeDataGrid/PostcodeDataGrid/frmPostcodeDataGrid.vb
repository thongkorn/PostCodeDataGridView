#Region "ABOUT"
' / --------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/
' /
' / Purpose: Postcode Thailand on DataGridView
' / Microsoft Visual Basic .NET (2010) + MS Access 2007
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------
#End Region

Imports System.Data.OleDb

Public Class frmPostcodeDataGrid
    Dim Conn As New OleDbConnection
    Dim Cmd As New OleDbCommand
    Dim DR As OleDbDataReader
    Dim DA As New OleDbDataAdapter
    Dim Sql As String = String.Empty
    Dim ColProvince As New DataGridViewComboBoxColumn
    '//
    Private Sub frmPostcodeDataGrid_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim ColAmphur As New DataGridViewComboBoxColumn
        Dim ColTumbon As New DataGridViewComboBoxColumn
        Dim ColPostCode As New DataGridViewTextBoxColumn
        '// กำหนด Path ของไฟล์ฐานข้อมูล
        Dim strPath As String = Application.StartupPath
        strPath = strPath.ToLower()
        strPath = strPath.Replace("\bin\debug", "\").Replace("\bin\release", "\")

        Dim strConn As String = _
            " Provider=Microsoft.ACE.OLEDB.12.0;" & _
            " Data Source = " & strPath & "Data\PostCode2555.accdb; " & _
            " Persist Security Info=False;"
        Try
            '// เปิดการเชื่อมต่อไฟล์ฐานข้อมูล
            Conn = New OleDb.OleDbConnection(strConn)
            Conn.Open()
            With DataGridView1
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .SelectionMode = DataGridViewSelectionMode.CellSelect
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                '.AutoResizeColumns()
                .MultiSelect = False
            End With
            With ColProvince
                .Name = "Province"
                .HeaderText = "จังหวัด"
                .ReadOnly = False
                .MaxDropDownItems = 10
            End With
            DataGridView1.Columns.Add(ColProvince)
            Call LoadProvice()
            '//
            With ColAmphur
                .Name = "Amphur"
                .HeaderText = "อำเภอ/เขต"
                .ReadOnly = False
                .MaxDropDownItems = 10
            End With
            DataGridView1.Columns.Add(ColAmphur)
            '//
            With ColTumbon
                .Name = "Tumbon"
                .HeaderText = "ตำบล"
                .ReadOnly = False
                .MaxDropDownItems = 10
            End With
            DataGridView1.Columns.Add(ColTumbon)
            '//
            With ColPostCode
                .ReadOnly = True
                .HeaderText = "รหัสไปรษณีย์"
                .Name = "PostCode"
            End With
            DataGridView1.Columns.Add(ColPostCode)
            '// Add new row.
            DataGridView1.Rows.Add("", "", "", "")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Application.Exit()
        End Try

    End Sub

    ' / โหลดรายชื่อจังหวัดเข้าสู่ ComboBox
    Sub LoadProvice()
        '// DISTINCT คือ หากชื่อรายการมันซ้ำ ต้องตัดให้เหลือเพียงรายการเดียว
        Sql = _
            " SELECT DISTINCT PostCode.Province " & _
            " From PostCode ORDER BY PostCode.Province "
        Cmd = New OleDbCommand(Sql, Conn)
        DR = Cmd.ExecuteReader
        While DR.Read()
            ColProvince.Items.Add(DR("Province").ToString)
        End While
        DR.Close()
        Cmd.Dispose()
    End Sub

    Sub LoadAmphur(ByVal Province As String, ByVal row As Integer)
        Sql = _
                " SELECT DISTINCT PostCode.Amphur, PostCode.Province " & _
                " From PostCode  " & _
                " WHERE " & _
                " Province = " & "'" & Province & "'" & _
                " ORDER BY PostCode.Amphur "
        Try
            Dim Grow As DataGridViewRow = DataGridView1.Rows(row)
            Dim MyCell As DataGridViewComboBoxCell
            MyCell = Grow.Cells.Item("Amphur")
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            DA = New OleDb.OleDbDataAdapter(Sql, Conn)
            Dim DT As New DataTable
            DA.Fill(DT)
            MyCell.Items.Clear()
            For Each DR As DataRow In DT.Rows
                MyCell.Items.Add(DR("Amphur"))
            Next
            DT.Dispose()
            DA.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub LoadTumbon(ByVal Province As String, ByVal Amphur As String, ByVal row As Integer)
        Sql = _
            " SELECT PostCode.Amphur, PostCode.Province, PostCode.Tumbon " & _
            " From PostCode  " & _
            " WHERE " & _
            " Province = " & "'" & Province & "'" & _
            " AND " & _
            " Amphur = " & "'" & Amphur & "'" & _
            " ORDER BY PostCode.Tumbon "
        Try
            Dim Grow As DataGridViewRow = DataGridView1.Rows(row)
            Dim MyCell As DataGridViewComboBoxCell
            MyCell = Grow.Cells.Item("Tumbon")
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            DA = New OleDb.OleDbDataAdapter(Sql, Conn)
            Dim DT As New DataTable
            DA.Fill(DT)
            MyCell.Items.Clear()
            For Each DR As DataRow In DT.Rows
                MyCell.Items.Add(DR("Tumbon"))
            Next
            DT.Dispose()
            DA.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub LoadPostCode(ByVal Province As String, ByVal Amphur As String, ByVal Tumbon As String, ByVal row As Integer)
        Sql = _
            " SELECT PostCode.Province, PostCode.Amphur, PostCode.Tumbon, " & _
            " PostCode.PostCode, PostCode.Remark " & _
            " From PostCode  " & _
            " WHERE " & _
            " Province = " & "'" & Province & "'" & _
            " AND " & _
            " Amphur = " & "'" & Amphur & "'" & _
            " AND " & _
            " Tumbon = " & "'" & Tumbon & "'" & _
            " ORDER BY PostCode.Tumbon "
        Try
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Cmd = New OleDbCommand(Sql, Conn)
            DR = Cmd.ExecuteReader
            While DR.Read()
                DataGridView1.Rows(row).Cells(3).Value = DR.Item("PostCode").ToString
            End While
            DR.Close()
            Cmd.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Remove selected row
    Private Sub btnRemoveRow_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveRow.Click
        If DataGridView1.RowCount = 0 Then Exit Sub
        DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
    End Sub

    Private Sub btnAddRow_Click(sender As System.Object, e As System.EventArgs) Handles btnAddRow.Click
        DataGridView1.Rows.Add("", "", "", "")
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim Province As String = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Dim Amphur As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Dim Tumbon As String = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Select Case e.ColumnIndex
            Case 0
                DataGridView1(1, e.RowIndex).Value = Nothing
                '// ส่งค่าจังหวัดและแถวปัจจุบัน
                Call LoadAmphur(Province, e.RowIndex)
            Case 1
                DataGridView1(2, e.RowIndex).Value = Nothing
                '// ส่งค่าจังหวัด อำเภอและแถวปัจจุบัน
                If Province <> Nothing And Amphur <> Nothing Then Call LoadTumbon(Province, Amphur, e.RowIndex)
            Case 2
                DataGridView1(3, e.RowIndex).Value = Nothing
                '// ส่งค่าจังหวัด อำเภอ ตำบลและแถวปัจจุบัน
                If Province <> Nothing AndAlso Amphur <> Nothing AndAlso Tumbon <> Nothing Then
                    Call LoadPostCode(Province, Amphur, Tumbon, e.RowIndex)
                End If
        End Select
    End Sub

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        Select Case e.ColumnIndex
            '// Update
            Case 2
                DataGridView1.EndEdit()
        End Select
    End Sub
End Class
