Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Main
    'Excel變量
    Public app As New Excel.Application 'app 是操作 Excel 的變數
    Public worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
    Public workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
    Public xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
    Public misvalue As Object = System.Reflection.Missing.Value
    '輸入輸出文件變量
    Public output_folder As String
    Public files() As String
    Public count As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        selectallfiles.Checked = True
    End Sub
    Private Sub InputButton_Click(sender As Object, e As EventArgs) Handles InputButton.Click
        '清理列表
        ListBox1.Items.Clear()
        files = Nothing
        If selectallfiles.Checked = True Then
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                TextBox1.Text = FolderBrowserDialog1.SelectedPath
                files = IO.Directory.GetFiles(FolderBrowserDialog1.SelectedPath)
                '更新列表顯示路徑内檔案
                For Each file As String In files
                    ListBox1.Items.Add(file)
                Next
            End If
        Else
            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
                TextBox1.Text = OpenFileDialog1.SafeFileName
                files = {OpenFileDialog1.FileName}
            End If
        End If
    End Sub
    '主要處理程序
    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        '檢查Excel是否安裝妥當
        If output_folder = "" Or TextBox1.Text = "" Then
            MsgBox("Please select the destinations first.")
        Else

            On Error Resume Next
            xlApp = GetObject(, "Excel.Application")
            If Err.Number() <> 0 Then
                Err.Clear()
                xlApp = CreateObject("Excel.Application")
                If Err.Number() <> 0 Then
                    MsgBox("Excel is not properly installed!!")
                    End
                End If
            End If
            'Check paths selected
            Dim result1 As DialogResult = MessageBox.Show("Are you sure you want to process files and save to " + output_folder + " ?",
                "Confirmation",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
            If result1 = DialogResult.Yes Then

                If (TextBox1.Text <> "") And (TextBox2.Text <> "") And Err.Number() = 0 Then

                    xlApp.DisplayAlerts = False
                    InputButton.Enabled = False
                    RunButton.Enabled = False
                    OutputButton.Enabled = False
                    exitbut.Enabled = False
                    '清理列表
                    ListBox1.Items.Clear()
                    '提取Excel文檔及數量
                    For Each file As String In files
                        If (Path.GetExtension(file) = ".xls") Or (Path.GetExtension(file) = ".xlsx") Then
                            ListBox1.Items.Add(file)
                            count += 1
                        End If
                    Next
                    'Main Workbook loop
                    For Each file As String In files
                        If (Path.GetExtension(file) = ".xlsx") Or (Path.GetExtension(file) = ".xls") Then
                            ''Open workbooks
                            ListBox2.Items.Add(file)
                            workbook = app.Workbooks.Open(file)
                            ''Worksheet loop
                            For i As Integer = 1 To workbook.Sheets.Count
                                Dim testbook As Excel.Workbook = app.Workbooks.Add(1)
                                workbook.Sheets(i).copy(testbook.Sheets(1))
                                testbook.SaveAs(output_folder + "\" + workbook.Sheets(i).Name + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook)
                                ToolStripStatusLabel1.Text = "Now processing [" + workbook.Name + "] to [" + workbook.Sheets(i).Name + "]"
                                If ToolStripProgressBar1.Value <= ToolStripProgressBar1.Maximum - 1 Then
                                    ToolStripProgressBar1.Value += 100 / count
                                End If
                                testbook = Nothing
                                testbook.Close()
                            Next
                        End If
                    Next
                    MessageBox.Show("Output done.")

                    'Enable the buttons
                    exitbut.Enabled = True
                    InputButton.Enabled = True
                    RunButton.Enabled = True
                    OutputButton.Enabled = True

                    'Quit Excel & garbage collect
                    xlApp.Quit()
                    workbook.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                    xlApp = Nothing
                    workbook = Nothing
                    worksheet = Nothing
                    GC.Collect()

                    'Error conditions
                ElseIf Err.Number() <> 0 Then
                    MessageBox.Show("Please deal with the excel error problem first. Error=" + Err.Number(),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1)
                Else
                    MsgBox("Please select the location first.")
                End If
            End If
        End If

    End Sub
    Private Sub OutputButton_Click(sender As Object, e As EventArgs) Handles OutputButton.Click
        If (FolderBrowserDialog2.ShowDialog() = DialogResult.OK) Then
            TextBox2.Text = FolderBrowserDialog2.SelectedPath
            output_folder = FolderBrowserDialog2.SelectedPath
        End If
    End Sub

    Private Sub selectallfiles_CheckedChanged(sender As Object, e As EventArgs) Handles selectallfiles.CheckedChanged
        TextBox1.Text = Nothing
        ListBox1.Items.Clear()
    End Sub

    Private Sub selectsinglefile_CheckedChanged(sender As Object, e As EventArgs) Handles selectsinglefile.CheckedChanged
        TextBox1.Text = Nothing
        ListBox1.Items.Clear()
    End Sub

    'Quit app & garbage collect again
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles exitbut.Click
        xlApp = Nothing
        workbook = Nothing
        worksheet = Nothing
        GC.Collect()
        Close()
    End Sub
End Class




