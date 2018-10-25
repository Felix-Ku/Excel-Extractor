'Function: Extract Excel Workbooks
'Version: 3.0
'Last updated date: 24/10/2018

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO


Public Class Main

    'Excel變量 Excel variables
    Private app As New Excel.Application 'app 是操作 Excel 的變數
    Private worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
    Private workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
    Private xlApp As Excel.Application
    Private misvalue As Object = System.Reflection.Missing.Value
    '輸入輸出文件變量 Files
    Private output_folder As String
    Private files() As String
    Private count As Integer = 0
    Private filenums As Integer = 0

    Private sheetname As New List(Of String)
    Private boxname As New List(Of Object)

    'Messages
    Private result1 As DialogResult
    Private resultoverwrite As DialogResult
    Private resultoverwriteall As DialogResult

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

    '主要處理程序
    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles RunButton.Click

        xlApp = New Microsoft.Office.Interop.Excel.Application()

        If output_folder = "" Or TextBox1.Text = "" Then
            MsgBox("Please select the destinations first.")
        Else
            'Initializing
            ToolStripProgressBar1.Value = 0
            verifyexcelApp()
            numexcel() 'Calculate number of excel files
            If count > 0 And Err.Number() = 0 Then
                'If confirmed proceeding
                messageConfirm()
                If result1 = DialogResult.Yes Then

                    xlApp.DisplayAlerts = True
                    app.DisplayAlerts = True

                    buttonhide() 'Disable buttons
                    'Main Workbook loop

                    For Each file As String In files
                        If ((Path.GetExtension(file) = ".xlsx") Or (Path.GetExtension(file) = ".xls")) Then
                            ''Open workbooks
                            ListBox2.Items.Add(Path.GetFileName(file))
                            workbook = app.Workbooks.Open(file)
                            'Initialize progress bar for next file
                            ToolStripProgressBar1.Value = 0
                            ToolStripProgressBar1.Minimum = 0
                            ''Worksheet loop
                            For i As Integer = 1 To workbook.Sheets.Count

                                ToolStripProgressBar1.Maximum = workbook.Sheets.Count


                                Dim SavePath As String = output_folder + "\" + workbook.Sheets(i).Name + ".xlsx"
                                '檢查文件重復性 Check file exist and ask overwrite or not
                                If System.IO.File.Exists(SavePath) = True And xlApp.DisplayAlerts = True Or app.DisplayAlerts = True Then
                                    Dim exist1 As DialogResult = MessageBox.Show("File " + SavePath + " already exists, do you want to overwrite all remainings?",
                "Overwrite?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
                                    If exist1 = DialogResult.Yes Then
                                        xlApp.DisplayAlerts = False
                                        app.DisplayAlerts = False
                                    End If
                                End If

                                Dim problemname As String = workbook.Sheets(i).Name

                                If problemname <> "Ls_XLB_WorkbookFile" And problemname <> "Ls_AgXLB_WorkbookFile" Then

                                    worksheet = workbook.Sheets(i)
                                    sheetname.Add(worksheet.Name)
                                    boxname.Add(worksheet.Range("M101:M101").Value)

                                    Dim testbook As Excel.Workbook = app.Workbooks.Add(1)
                                    filenums += 1
                                    workbook.Sheets(i).copy(testbook.Sheets(1))

                                    Try
                                        testbook.Sheets(2).delete
                                    Catch ex As Exception
                                    End Try
                                    testbook.SaveAs(SavePath, Excel.XlFileFormat.xlOpenXMLWorkbook)
                                    testbook.Close(True)
                                    testbook = Nothing
                                    ToolStripStatusLabel1.Text = "Now processing [" + workbook.Name + "] to [" + workbook.Sheets(i).Name + "]..."
                                    ToolStripProgressBar1.Value += 1


                                End If
                            Next
                            ToolStripProgressBar1.Value = ToolStripProgressBar1.Maximum
                        End If
                    Next
                    savetoini()
                    messagesuccess() 'Message:Success

                    'Release objects
                    ''sheetname.Clear()
                    ''boxname.Clear()
                    closeObject(workbook)
                    quitObject(xlApp)
                    quitObject(app)
                    releaseObject(worksheet)
                    releaseObject(workbook)



                    buttonshow() 'Enable the buttons

                    '錯誤情況 Error conditions

                End If

            ElseIf Err.Number() <> 0 Then
                messageExcelError()
            Else
                MsgBox("No Excel file(s) exist in the source destination! Please select destination again!")
            End If
        End If
    End Sub

    'Excel app & Files loop----------------------------------------------------
    '檢查Excel是否安裝妥當 Check whether excel is installed properly
    Private Sub verifyexcelApp()
        Try
            xlApp = GetObject(, "Excel.Application")
            If Err.Number() <> 0 Then
                Err.Clear()
                xlApp = CreateObject("Excel.Application")
                If Err.Number() <> 0 Then
                    MsgBox("Excel is not properly installed!!")
                    End
                End If
            End If
        Catch ex As Exception
            MsgBox("Excel is not properly installed!!")
        End Try
    End Sub
    '數Excel文件數量及驗證Excel檔案存在性 Check whether and how many excel files exist
    Private Sub numexcel()
        If FolderBrowserDialog1.SelectedPath <> "" Or System.IO.File.Exists(OpenFileDialog1.FileName) And files IsNot Nothing Then
            For Each file As String In files
                If ((Path.GetExtension(file) = ".xls") Or (Path.GetExtension(file) = ".xlsx")) And Path.GetFileName(file)(0) <> "~" Then
                    count += 1
                End If
            Next
        Else
            MsgBox("Destinations invalid! Please check again!")
        End If
    End Sub
    Private Sub inifile()


    End Sub

    'Buttons-------------------------------------------------------
    'Buttons show
    Private Sub buttonshow()
        exitbut.Enabled = True
        InputButton.Enabled = True
        RunButton.Enabled = True
        OutputButton.Enabled = True
    End Sub
    'Buttons hide
    Private Sub buttonhide()
        InputButton.Enabled = False
        RunButton.Enabled = False
        OutputButton.Enabled = False
        exitbut.Enabled = False
    End Sub

    'Messages------------------------------------------------------
    'Message: Ask user to confirm proceeding or not 確認是否繼續
    Private Sub messageConfirm()
        result1 = MessageBox.Show("Are you sure you want to process files and save to " + output_folder + " ?",
            "Confirmation",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2)
    End Sub
    'Message: Ask user
    Private Sub messageExcelError()
        Dim error2 As DialogResult = MessageBox.Show("Please deal with the excel error problem first. Error=" + Err.Number(),
                       "Error",
               MessageBoxButtons.OK,
               MessageBoxIcon.Question,
               MessageBoxDefaultButton.Button2)
    End Sub
    'Message: Ask user 
    Private Sub messagesuccess()
        MessageBox.Show("Output done.")
        ToolStripStatusLabel1.Text = "Output done. Press exit to leave or choose other file(s) to restart"
        Dim numfiles As DialogResult = MessageBox.Show(filenums & " file(s) output done",
                "Done",
                MessageBoxButtons.OK,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
    End Sub
    'Message: Ask user 
    Private Sub messageConfirm4()

    End Sub

    'System-------------------------------------------------------------
    '資源回收 Quit object
    Private Sub quitObject(ByVal obj As Object)
        Try
            obj.Quit()
        Catch ex As Exception
        End Try
    End Sub
    '資源回收 close object
    Private Sub closeObject(ByVal obj As Object)
        Try
            obj.Close()
        Catch ex As Exception
        End Try
    End Sub
    '資源回收 release object
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    '退出程式 Quit app
    Private Sub Exit_Click(sender As Object, e As EventArgs) Handles exitbut.Click
        Close()
    End Sub
    '清理Excel相關程序及資源回收 Quit Excel & garbage collect
    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            releaseObject(xlApp)
        Catch ex As Exception
        End Try
        closeObject(workbook)
        quitObject(xlApp)
        quitObject(app)
        releaseObject(worksheet)
        releaseObject(workbook)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub savetoini()
        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim inipath As String = baseDir + "tranpath.ini"
        Dim iniwrite As System.IO.StreamWriter

        For i As Integer = 0 To boxname.Count - 1
            ListBox2.Items.Add(sheetname(i) & "," & boxname(i))
        Next

        Try
            'If inifile originally not exist
            If Not System.IO.File.Exists(inipath) Then
                File.Create(inipath).Dispose() 'Create ini
                iniwrite = My.Computer.FileSystem.OpenTextFileWriter(inipath, True)
                For i As Integer = 0 To boxname.Count - 1
                    iniwrite.WriteLine(sheetname(i) & "," & boxname(i))
                Next
                iniwrite.Close()

                'If inifile originally exist
            Else
                Dim lines() As String = IO.File.ReadAllLines(inipath)
                Dim linesdy As List(Of String) = New List(Of String)(lines)

                'For j=number of lines in new excels | Fori=number of lines in existing ini file
                For j As Integer = 0 To boxname.Count - 1
                    Dim written As Integer = 0
                    For i As Integer = 0 To lines.Length - 1
                        If linesdy(i).Contains(sheetname(j)) Then
                            linesdy(i) = (sheetname(j) & "," & boxname(j))
                            written = 1
                        ElseIf written = 0 And i = lines.Length - 1 Then
                            linesdy.Add(sheetname(j) & "," & boxname(j))
                        End If

                    Next
                Next
                IO.File.WriteAllLines(inipath, linesdy)
                'iniwrite = My.Computer.FileSystem.OpenTextFileWriter(inipath, True)
            End If

        Catch ex As Exception
            MessageBox.Show("The saving process failed: {0}", ex.ToString())

        End Try

    End Sub

End Class