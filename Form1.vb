Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Public elements(1) As String
    Public elements2(1) As String
    Public elements3(1) As String
    Public rowcount As Integer
    Public loaded As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub
    Sub invcalc()
        Dim index As Integer = 0
        Dim total As Decimal = 0
        Dim scanned As Integer = 0
        Dim cost As Decimal
        On Error Resume Next
        'flat goods
        While index < DataGridView2.RowCount
            total = total + (DataGridView2.Rows(index).Cells("Unreturned").Value * DataGridView2.Rows(index).Cells(22).Value)
            index = index + 1
        End While
        'unreturned
        index = 0
        While index < DataGridView1.RowCount
            If DataGridView1.Rows(index).Cells(23).Value.ToString = "Y" Then cost = DataGridView1.Rows(index).Cells(25).Value Else cost = DataGridView1.Rows(index).Cells(22).Value
            total = total + (DataGridView1.Rows(index).Cells("Unreturned2").Value * cost)
            scanned = scanned + DataGridView1.Rows(index).Cells("CI").Value - DataGridView1.Rows(index).Cells("AXF045").Value
            'MsgBox(DataGridView1.Rows(index).Cells("unret").Value)
            index = index + 1
        End While
        'damaged
        index = 0
        While index < DataGridView1.RowCount
            If DataGridView1.Rows(index).Cells(23).Value = "Y" Then cost = DataGridView1.Rows(index).Cells(25).Value Else cost = DataGridView1.Rows(index).Cells(22).Value
            If DataGridView1.Rows(index).Cells("Damaged").Value > DataGridView1.Rows(index).Cells("Unreturned2").Value Then
                total = total + ((DataGridView1.Rows(index).Cells("CI").Value - DataGridView1.Rows(index).Cells("Unreturned2").Value) * cost)
            Else
                total = total + (DataGridView1.Rows(index).Cells("Damaged").Value * cost)
            End If
            index = index + 1
        End While
        TextBox3.Text = FormatCurrency(total)
        TextBox4.Text = scanned
    End Sub
    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim path As String
        Dim strm As System.IO.Stream
        strm = OpenFileDialog1.OpenFile()
        path = OpenFileDialog1.FileName.ToString()
        If Not (strm Is Nothing) Then
            'DataSet1.Clear()
            'DataSet2.Clear()
            'DataSet3.Clear()
            loaded = 0
            DataGridView1.ColumnHeadersVisible = True
            DataGridView2.ColumnHeadersVisible = True
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim MyCommand, mycommand2, mycommand3 As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & path & "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where CI > 0 AND AXF129 <> ''", MyConnection)
            MyCommand.TableMappings.Add("Table", "Andrew")
            strm.Close()
            MyCommand.Fill(DataSet2)
            DataGridView1.DataSource = DataSet2.Tables(0)
            mycommand2 = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where CI > 0 AND AXF129 = ''", MyConnection)
            mycommand2.TableMappings.Add("Table2", "Andrew")
            mycommand2.Fill(DataSet3)
            DataGridView2.DataSource = DataSet3.Tables(0)
            mycommand3 = New System.Data.OleDb.OleDbDataAdapter("select [Sheet2$].SDI,[Sheet2$].Barcode,[Sheet1$].AXAXCD,[Sheet1$].AXA1CD from [Sheet2$],[Sheet1$] WHERE [Sheet2$].SDI = [Sheet1$].AXBYCD", MyConnection)
            mycommand3.TableMappings.Add("Table3", "Andrew")
            mycommand3.Fill(DataSet1)
            Form2.DataGridView1.DataSource = DataSet1.Tables(0)
            MyConnection.Close()
            Dim index As Integer = 0
            DataGridView2.CurrentCell = Nothing
            Do While index < DataGridView2.RowCount
                DataGridView2.Rows(index).Cells(31).Value = DataGridView2.Rows(index).Cells(29).Value
                index = index + 1
            Loop
            index = 0
            Dim remaining As Integer = 0
            DataGridView1.CurrentCell = Nothing
            Do While index < DataGridView1.RowCount
                DataGridView1.Rows(index).Cells(31).Value = DataGridView1.Rows(index).Cells(29).Value
                DataGridView1.Rows(index).Cells(30).Value = DataGridView1.Rows(index).Cells(29).Value
                DataGridView1.Rows(index).Cells(32).Value = 0
                remaining = remaining + DataGridView1.Rows(index).Cells(31).Value
                index = index + 1
            Loop
            TextBox2.Text = remaining
            rowcount = Form2.DataGridView1.RowCount
            ReDim elements(rowcount)
            ReDim elements2(rowcount)
            ReDim elements3(rowcount)
            index = 0
            Form2.DataGridView1.CurrentCell = Nothing
            Do While index < rowcount
                elements(index) = Form2.DataGridView1.Rows(index).Cells(1).Value.ToString
                elements2(index) = "U"
                elements3(index) = Form2.DataGridView1.Rows(index).Cells(0).Value.ToString & Form2.DataGridView1.Rows(index).Cells(2).Value.ToString & Form2.DataGridView1.Rows(index).Cells(3).Value.ToString
                index = index + 1
            Loop
            loaded = 1
            Call invcalc()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "SCAN HERE" Then
            TextBox1.Text = ""
            TextBox1.ForeColor = Color.Black
            TextBox1.TextAlign = HorizontalAlignment.Left
        End If
    End Sub
    Private Sub txtUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim index As Integer = 0
            Dim index2 As Integer = 0
            Dim thing, thing2 As String
            If IsNumeric(TextBox1.Text) Then thing = TextBox1.Text + 0 Else thing = TextBox1.Text
            Dim dr As String
            If RadioButton2.Checked = True Then dr = "R" Else dr = "D"
            While index <= rowcount
                thing2 = elements(index)
                'MsgBox(thing & " " & thing2)
                If thing = thing2 Then
                    Label1.Visible = False
                    While index2 < DataGridView1.RowCount
                        If DataGridView1.Rows(index2).Cells("AXBYCD").Value & DataGridView1.Rows(index2).Cells("AXAXCD").Value & DataGridView1.Rows(index2).Cells("AXA1CD").Value = elements3(index) Then
                            If DataGridView1.Rows(index2).Cells("Unreturned2").Value > 0 And elements2(index) = "U" Then
                                TextBox2.Text = TextBox2.Text - 1
                                DataGridView1.Rows(index2).Cells("Unreturned2").Value = DataGridView1.Rows(index2).Cells("Unreturned2").Value - 1
                            End If
                            If dr = "D" And elements2(index) = "U" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value + 1
                            If dr = "R" And elements2(index) = "D" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value - 1
                            If dr = "D" And elements2(index) = "R" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value + 1
                            If Form2.Visible = True And Form2.Label4.Text = "SDI: " & elements3(index).Substring(0, 3) Then
                                If dr = "R" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Green
                                If dr = "D" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Red
                            End If
                            If DataGridView1.Rows(index2).Cells("AXBYCD").Value & DataGridView1.Rows(index2).Cells("AXAXCD").Value & DataGridView1.Rows(index2).Cells("AXA1CD").Value = elements3(index) And elements2(index) = "U" Then DataGridView1.Rows(index2).Cells(30).Value = DataGridView1.Rows(index2).Cells(30).Value - 1
                            elements2(index) = dr
                            DataGridView1.ClearSelection()
                            DataGridView1.FirstDisplayedScrollingRowIndex = index2
                            DataGridView1.Rows(index2).Selected = True
                            TextBox1.Clear()
                            Call invcalc()
                            Exit Sub
                        End If
                        index2 = index2 + 1
                    End While
                End If
                index = index + 1
            End While
            If index > rowcount Then Label1.Visible = True
            TextBox1.Clear()
        End If
    End Sub
    Private Sub DataGridView1_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        On Error GoTo theend
        Form2.Close()
        Form2.Visible = True
        Dim selectedRow = Me.DataGridView1.Rows(e.RowIndex)
        Form2.Label3.Text = selectedRow.Cells("AXF131").Value.ToString & " " & selectedRow.Cells("AXF132").Value.ToString
        Form2.Label1.Text = "Locker: " & selectedRow.Cells("AXBJNB").Value.ToString & "-" & selectedRow.Cells("AXF133").Value.ToString
        Form2.Label2.Text = "Wearer: " & selectedRow.Cells("AXF129").Value.ToString
        Form2.Label4.Text = "SDI: " & selectedRow.Cells("AXBYCD").Value.ToString
        Form2.DataGridView1.DataSource = DataSet1.Tables(0)
        Form2.DataGridView1.Refresh()
        Dim index As Integer = 0
        Form2.DataGridView1.CurrentCell = Nothing
        While index < rowcount
            If selectedRow.Cells("AXBYCD").Value & selectedRow.Cells("AXAXCD").Value & selectedRow.Cells("AXA1CD").Value <> Form2.DataGridView1.Rows(index).Cells(0).Value & Form2.DataGridView1.Rows(index).Cells(2).Value & Form2.DataGridView1.Rows(index).Cells(3).Value Then
                Form2.DataGridView1.Rows(index).Visible = False
            Else
                If elements2(index) = "R" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Green
                If elements2(index) = "D" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Red
            End If
            index = index + 1
        End While
        Exit Sub
theend:
        Form2.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim index As Integer = 0
        DataGridView2.CurrentCell = Nothing
        Do While index < DataGridView2.RowCount
            DataGridView2.Rows(index).Cells(31).Value = 0
            index = index + 1
        Loop
        Call invcalc()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For Each row As DataGridViewRow In DataGridView2.SelectedRows
            row.Cells(31).Value = 0
        Next
        Call invcalc()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If loaded = 1 Then
            Dim index As Integer = 0
            Dim remaining As Integer = 0
            ' DataGridView1.CurrentCell = Nothing
            Do While index < DataGridView1.RowCount
                remaining = remaining + DataGridView1.Rows(index).Cells(31).Value
                index = index + 1
            Loop
            TextBox2.Text = remaining
            Call invcalc()
        End If
    End Sub
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        If loaded = 1 Then Call invcalc()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles Button4.Click
        If loaded = 1 Then
            Dim excel_app As New Excel.ApplicationClass()
            ' Make Excel visible (optional).
            excel_app.Visible = True

            ' Open the workbook.
            Dim workbook As Excel.Workbook = _
                excel_app.Workbooks.Open(Application.StartupPath & "\template.xls")

            ' See if the worksheet already exists.
            Dim sheet_name As String = "Sheet1"

            Dim sheet As Excel.Worksheet = FindSheet(workbook, _
                sheet_name)
            Dim value_range As Excel.Range = sheet.Range("A9")
            value_range.Value2 = DataGridView1.Rows(0).Cells(2).Value
            value_range = sheet.Range("C14")
            value_range.Value2 = DataGridView1.Rows(0).Cells(3).Value & "-" & DataGridView1.Rows(0).Cells(4).Value
            value_range = sheet.Range("C15")
            value_range.Value2 = DataGridView1.Rows(0).Cells(0).Value
            'lost
            Dim index As Integer = 0
            Do While index < DataGridView1.RowCount
                If DataGridView1.Rows(index).Cells(31).Value > 0 Then
                    sheet.Rows("22:22").insert()
                    value_range = sheet.Range("A22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(19).Value
                    value_range = sheet.Range("B22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(5).Value
                    value_range = sheet.Range("C22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(6).Value
                    value_range = sheet.Range("D22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(7).Value
                    value_range = sheet.Range("E22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(14).Value
                    value_range = sheet.Range("F22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(23).Value
                    value_range = sheet.Range("G22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(31).Value
                    value_range = sheet.Range("H22")
                    If DataGridView1.Rows(index).Cells(23).Value.ToString = "Y" Then value_range.Value2 = DataGridView1.Rows(index).Cells(25).Value
                    If DataGridView1.Rows(index).Cells(23).Value.ToString <> "Y" Then value_range.Value2 = DataGridView1.Rows(index).Cells(22).Value
                    value_range = sheet.Range("I22")
                    value_range.Value2 = "=G22*H22"
                    value_range = sheet.Range("J22")
                    value_range.Value2 = "LM"
                End If
                index = index + 1
            Loop
            index = 0
            'damaged
            Do While index < DataGridView1.RowCount
                Dim total As Integer = 0
                If DataGridView1.Rows(index).Cells("Damaged").Value > DataGridView1.Rows(index).Cells("Unreturned2").Value Then
                    total = DataGridView1.Rows(index).Cells("CI").Value - DataGridView1.Rows(index).Cells("Unreturned2").Value
                Else
                    total = DataGridView1.Rows(index).Cells("Damaged").Value
                End If
                If total > 0 Then
                    sheet.Rows("22:22").insert()
                    value_range = sheet.Range("A22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(19).Value
                    value_range = sheet.Range("B22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(5).Value
                    value_range = sheet.Range("C22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(6).Value
                    value_range = sheet.Range("D22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(7).Value
                    value_range = sheet.Range("E22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(14).Value
                    value_range = sheet.Range("F22")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(23).Value
                    value_range = sheet.Range("G22")
                    value_range.Value2 = total
                    value_range = sheet.Range("H22")
                    If DataGridView1.Rows(index).Cells(23).Value.ToString = "Y" Then value_range.Value2 = DataGridView1.Rows(index).Cells(25).Value
                    If DataGridView1.Rows(index).Cells(23).Value.ToString <> "Y" Then value_range.Value2 = DataGridView1.Rows(index).Cells(22).Value
                    value_range = sheet.Range("I22")
                    value_range.Value2 = "=G22*H22"
                    value_range = sheet.Range("J22")
                    value_range.Value2 = "AB"
                End If
                index = index + 1
            Loop
            'flat goods
            index = 0
            Do While index < DataGridView2.RowCount
                If DataGridView2.Rows(index).Cells(31).Value > 0 Then
                    sheet.Rows("22:22").insert()
                    value_range = sheet.Range("A22")
                    value_range.Value2 = DataGridView2.Rows(index).Cells(19).Value
                    value_range = sheet.Range("E22")
                    value_range.Value2 = DataGridView2.Rows(index).Cells(14).Value
                    value_range = sheet.Range("G22")
                    value_range.Value2 = DataGridView2.Rows(index).Cells(31).Value
                    value_range = sheet.Range("H22")
                    value_range.Value2 = DataGridView2.Rows(index).Cells(22).Value
                    value_range = sheet.Range("I22")
                    value_range.Value2 = "=G22*H22"
                    value_range = sheet.Range("J22")
                    value_range.Value2 = "LM"
                End If
                index = index + 1
            Loop
        End If
    End Sub
    Private Function FindSheet(ByVal workbook As Excel.Workbook, _
        ByVal sheet_name As String) As Excel.Worksheet
        For Each sheet As Excel.Worksheet In workbook.Sheets
            If (sheet.Name = sheet_name) Then Return sheet
        Next sheet

        Return Nothing
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If loaded = 1 Then
            Dim excel_app As New Excel.ApplicationClass()
            ' Make Excel visible (optional).
            excel_app.Visible = True
            ' Open the workbook.
            Dim workbook As Excel.Workbook = _
                excel_app.Workbooks.Add
            ' See if the worksheet already exists.
            Dim sheet_name As String = "Sheet1"
            Dim sheet As Excel.Worksheet = FindSheet(workbook, _
                sheet_name)
            Dim value_range As Excel.Range = sheet.Range("A1")
            value_range.Value2 = "Wearer #"
            value_range = sheet.Range("B1")
            value_range.Value2 = "First Name"
            value_range = sheet.Range("C1")
            value_range.Value2 = "Last Name"
            value_range = sheet.Range("D1")
            value_range.Value2 = "Bank"
            value_range = sheet.Range("E1")
            value_range.Value2 = "Locker"
            value_range = sheet.Range("F1")
            value_range.Value2 = "Item"
            value_range = sheet.Range("G1")
            value_range.Value2 = "CI"
            value_range = sheet.Range("H1")
            value_range.Value2 = "Damaged"
            value_range = sheet.Range("I1")
            value_range.Value2 = "Scanned"
            Dim index As Integer = 0
            Dim row As Integer = 2
            sheet.Columns("A:A").NumberFormat = "@"
            Do While index < DataGridView1.RowCount
                If DataGridView1.Rows(index).Cells(29).Value - DataGridView1.Rows(index).Cells(30).Value > 0 Then
                    value_range = sheet.Range("A" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(5).Value
                    value_range = sheet.Range("B" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(6).Value
                    value_range = sheet.Range("C" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(7).Value
                    value_range = sheet.Range("D" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(8).Value
                    value_range = sheet.Range("E" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(9).Value
                    value_range = sheet.Range("F" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(14).Value
                    value_range = sheet.Range("G" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(29).Value
                    value_range = sheet.Range("I" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(29).Value - DataGridView1.Rows(index).Cells(30).Value
                    value_range = sheet.Range("H" & row & "")
                    value_range.Value2 = DataGridView1.Rows(index).Cells(32).Value
                    row = row + 1
                End If
                index = index + 1
            Loop
            sheet_name = "Sheet2"
            sheet = FindSheet(workbook, _
                sheet_name)
            index = 0
            row = 2
            Form2.DataGridView1.CurrentCell = Nothing
            sheet.Columns("A:A").NumberFormat = "@"
            value_range = sheet.Range("A1")
            value_range.Value2 = "Barcode"
            value_range = sheet.Range("B1")
            value_range.Value2 = "Status"
            Do While index < elements.GetLength(0)
                If elements2(index) = "D" Or elements2(index) = "R" Then
                    value_range = sheet.Range("A" & row & "")
                    value_range.Value2 = elements(index)
                    value_range = sheet.Range("B" & row & "")
                    value_range.Value2 = elements2(index)
                    row = row + 1
                End If
                index = index + 1
            Loop
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If loaded = 1 Then OpenFileDialog2.ShowDialog()
    End Sub
    Private Sub OpenFileDialog2_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        Dim path As String
        Dim strm As System.IO.Stream
        strm = OpenFileDialog2.OpenFile()
        path = OpenFileDialog2.FileName.ToString()
        If Not (strm Is Nothing) Then
            DataSet4.Clear()
            Dim MyConnection4 As System.Data.OleDb.OleDbConnection
            Dim MyCommand4 As System.Data.OleDb.OleDbDataAdapter
            MyConnection4 = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & path & "';Extended Properties=Excel 8.0;")
            MyCommand4 = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet2$]", MyConnection4)
            MyCommand4.TableMappings.Add("Table4", "Andrew")
            strm.Close()
            MyCommand4.Fill(DataSet4)
            MyConnection4.Close()
            Dim indexseven As Integer = 0
            Do While indexseven < DataSet4.Tables(0).Rows.Count
                Dim index As Integer = 0
                Dim index2 As Integer = 0
                Dim thing, thing2 As String
                thing = DataSet4.Tables(0).Rows(indexseven).Item(0).ToString
                Dim dr As String
                If DataSet4.Tables(0).Rows(indexseven).Item(1) = "R" Then dr = "R" Else dr = "D"
                While index <= rowcount
                    thing2 = elements(index)
                    If thing = thing2 Then
                        Label1.Visible = False
                        While index2 < DataGridView1.RowCount
                            If DataGridView1.Rows(index2).Cells("AXBYCD").Value & DataGridView1.Rows(index2).Cells("AXAXCD").Value & DataGridView1.Rows(index2).Cells("AXA1CD").Value = elements3(index) Then
                                If DataGridView1.Rows(index2).Cells("Unreturned2").Value > 0 And elements2(index) = "U" Then
                                    TextBox2.Text = TextBox2.Text - 1
                                    DataGridView1.Rows(index2).Cells("Unreturned2").Value = DataGridView1.Rows(index2).Cells("Unreturned2").Value - 1
                                End If
                                If dr = "D" And elements2(index) = "U" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value + 1
                                If dr = "R" And elements2(index) = "D" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value - 1
                                If dr = "D" And elements2(index) = "R" Then DataGridView1.Rows(index2).Cells("Damaged").Value = DataGridView1.Rows(index2).Cells("Damaged").Value + 1
                                If Form2.Visible = True And Form2.Label4.Text = "SDI: " & elements3(index).Substring(0, 3) Then
                                    If dr = "R" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Green
                                    If dr = "D" Then Form2.DataGridView1.Rows(index).DefaultCellStyle.BackColor = Color.Red
                                End If
                                If DataGridView1.Rows(index2).Cells("AXBYCD").Value = elements3(index) & DataGridView1.Rows(index2).Cells("AXAXCD").Value & DataGridView1.Rows(index2).Cells("AXA1CD").Value And elements2(index) = "U" Then DataGridView1.Rows(index2).Cells(30).Value = DataGridView1.Rows(index2).Cells(30).Value - 1
                                elements2(index) = dr
                                DataGridView1.ClearSelection()
                                DataGridView1.FirstDisplayedScrollingRowIndex = index2
                                DataGridView1.Rows(index2).Selected = True
                                Call invcalc()
                                GoTo here
                            End If
                            index2 = index2 + 1
                        End While
                    End If
                    index = index + 1
                End While
                If index > rowcount Then Label1.Visible = True
here:
                indexseven = indexseven + 1
            Loop
            Call invcalc()
        End If
    End Sub

End Class
