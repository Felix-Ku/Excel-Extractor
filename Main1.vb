Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Main
    'Excel變量 Variables for excel
    Public xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application() 'Excel main application
    Public worksheet As Excel.Worksheet 'Excel worksheet
    Public workbook As Excel.Workbook 'Excel workbook
    Public misvalue As Object = System.Reflection.Missing.Value

    '輸入輸出文件變量 Variables for file operation
    Public output_folder As String 'Output target location
    Public files() As String 'For storing  list of files under target location
    Public count As Integer = 0 'Store number of files to be processed
    Public overwritetoall As Boolean = False 'Switch for overwrite all or not
    Public overwrite As Boolean = False 'Switch for overwrite in the next procedure or not
    Public runclicked As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        selectallfiles.Checked = True '預設選擇輸出模式 Default output mode to "select all files"
    End Sub

    '來源路徑選擇 Choose the source files location
    Private Sub InputButton_Click(sender As Object, e As EventArgs) Handles InputButton.Click
        ToolStripProgressBar1.Value = 0
        files = Nothing '清理檔案 Clear files 
        If selectallfiles.Checked = True Then '選擇文件夾 Select all files mode
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                TextBox1.Text = FolderBrowserDialog1.SelectedPath
                files = IO.Directory.GetFiles(FolderBrowserDialog1.SelectedPath)
            End If
        Else '選擇單一文件 Select single file mode
            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
                TextBox1.Text = OpenFileDialog1.SafeFileName
                files = {OpenFileDialog1.FileName}
            End If
        End If

        If TextBox1.Text = "" Then
            ToolStripStatusLabel1.Text = "Please select the source destination"
        ElseIf TextBox2.Text = "" Then
            ToolStripStatusLabel1.Text = "Please select the output destination"
        Else
            ToolStripStatusLabel1.Text = "Please press Run to start"
        End If
    End Sub

    '輸出路徑選擇 Choose the target output location
    Private Sub OutputButton_Click(sender As Object, e As EventArgs) Handles OutputButton.Click
        ToolStripProgressBar1.Value = 0
        If (FolderBrowserDialog2.ShowDialog() = DialogResult.OK) Then
            TextBox2.Text = FolderBrowserDialog2.SelectedPath
            output_folder = FolderBrowserDialog2.SelectedPath
        End If
        If TextBox1.Text = "" Then
            ToolStripStatusLabel1.Text = "Please select the source destination"
        ElseIf TextBox2.Text = "" Then
            ToolStripStatusLabel1.Text = "Please select the output destination"
        Else
            ToolStripStatusLabel1.Text = "Please press Run to start"
        End If
    End Sub

    '選擇文件夾 Choice for selecting all files under folder
    Private Sub selectallfiles_CheckedChanged(sender As Object, e As EventArgs) Handles selectallfiles.CheckedChanged
        TextBox1.Text = Nothing
    End Sub

    '選擇單一文件 Choice for selecting single file
    Private Sub selectsinglefile_CheckedChanged(sender As Object, e As EventArgs) Handles selectsinglefile.CheckedChanged
        TextBox1.Text = Nothing
    End Sub

    '主要處理程序 Main program
    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        runclicked = True
        ToolStripProgressBar1.Value = 0
        '檢查Excel是否安裝妥當 Check whether excel is installed properly
        If output_folder = "" Or TextBox1.Text = "" Then
            MsgBox("Please select the destinations first.")
        Else

            '提取Excel文檔及數量 Checked the amount of excel files 
            If FolderBrowserDialog1.SelectedPath <> "" Or System.IO.File.Exists(OpenFileDialog1.FileName) And files IsNot Nothing Then
                For Each file As String In files
                    If Path.GetFileName(file)(0) <> "~" Then
                        If (Path.GetExtension(file) = ".xls") Or (Path.GetExtension(file) = ".xlsx") Then
                            workbook = xlApp.Workbooks.Open(file)
                            For i As Integer = 1 To workbook.Sheets.Count
                                count += 1
                            Next
                        End If
                    End If
                Next
            End If

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

            '確認是否繼續 Ask user to confirm proceeding or not
            Dim result1 As DialogResult = MessageBox.Show("Are you sure you want to process files and save to " + output_folder + " ?",
                "Confirmation",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)

            If result1 = DialogResult.Yes And count <> 0 Then

                If (TextBox1.Text <> "") And (TextBox2.Text <> "") And Err.Number() = 0 Then

                    '禁止進行中的按鈕使用 Stop buttons from functioning
                    xlApp.DisplayAlerts = False
                    InputButton.Enabled = False
                    RunButton.Enabled = False
                    OutputButton.Enabled = False
                    exitbut.Enabled = False

                    '主要處理程序 Main Workbook loop
                    For Each file As String In files
                        If (Path.GetExtension(file) = ".xlsx") Or (Path.GetExtension(file) = ".xls") And Path.GetFileName(file)(0) <> "~" Then
                            '打開工作本 Open workbooks
                            workbook = xlApp.Workbooks.Open(file)
                            '處理工作表 Worksheet loop
                            For i As Integer = 1 To workbook.Sheets.Count
                                Dim testbook As Excel.Workbook = xlApp.Workbooks.Add(1)
                                workbook.Sheets(i).copy(testbook.Sheets(1))
                                Dim SavePath As String = output_folder + "\" + workbook.Sheets(i).Name + ".xlsx"
                                '檢查文件重復性 Check file exist and ask overwrite or not
                                If System.IO.File.Exists(SavePath) And overwritetoall = False Then
                                    overwrite = False
                                    Dim exist1 As DialogResult = MessageBox.Show("File " + SavePath + " already exists, do you want to overwrite it?",
                "Overwrite?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
                                    If exist1 = DialogResult.Yes Then
                                        overwrite = True
                                        Dim exist2 As DialogResult = MessageBox.Show("Do you want to overwrite all remaining file(s)?",
                "Overwrite?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
                                        If exist2 = DialogResult.Yes Then
                                            overwritetoall = True
                                        End If
                                    End If
                                Else overwrite = True
                                End If
                                '主要轉存程序 Main saving operation

                                If overwrite = True Or overwritetoall = True Then
                                    ListBox2.Items.Add(workbook.Sheets(i).Name)
                                    testbook.SaveAs(SavePath, Excel.XlFileFormat.xlOpenXMLWorkbook)
                                    ToolStripStatusLabel1.Text = "Now processing [" + workbook.Name + "] to [" + workbook.Sheets(i).Name + ".xlsx]"
                                    If ToolStripProgressBar1.Value <= ToolStripProgressBar1.Maximum - 1 Then
                                        ToolStripProgressBar1.Value += 100 / count
                                    End If
                                End If
                                testbook = Nothing
                                testbook.Close()
                                GC.Collect()
                                workbook.Close()
                                workbook = Nothing
                                worksheet = Nothing
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                            Next
                        End If

                    Next
                    MessageBox.Show("Output done.")

                    '回復按鈕狀態 Enable the buttons and reset the variables
                    ToolStripProgressBar1.Value = 100
                    ToolStripStatusLabel1.Text = "Output done."
                    exitbut.Enabled = True
                    InputButton.Enabled = True
                    RunButton.Enabled = True
                    OutputButton.Enabled = True
                    overwritetoall = False




                    '錯誤情況 Error conditions
                ElseIf Err.Number() <> 0 Then
                    Dim error2 As DialogResult = MessageBox.Show("Please deal with the excel error problem first. Error=" + Err.Number(),
                        "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
                Else
                    MsgBox("Please select the location first.")
                End If
            ElseIf count = 0 Then
                Dim error1 As DialogResult = MessageBox.Show("No excel files exist in the source destination! Please check the destination path",
                    "Error Not found",
            MessageBoxButtons.OK,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2)
            End If
        End If

    End Sub

    '退出程式及再次垃圾收集 Quit app & garbage collect again
    Private Sub Exit_Click(sender As Object, e As EventArgs) Handles exitbut.Click
        Close()
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '清理Excel相關程序及資源回收 Quit Excel & garbage collect
        If runclicked Then
            xlApp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            xlApp = Nothing
            workbook = Nothing
            worksheet = Nothing
            GC.Collect()
        End If
    End Sub

End Class




