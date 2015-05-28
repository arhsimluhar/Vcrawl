Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDocumentName As System.Windows.Forms.TextBox
    Friend WithEvents btnParse As System.Windows.Forms.Button
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents lblFindCount As System.Windows.Forms.Label
    Friend WithEvents txtResults As System.Windows.Forms.TextBox
    Friend WithEvents txtDocContents As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.txtResults = New System.Windows.Forms.TextBox
        Me.lblFindCount = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.txtDocContents = New System.Windows.Forms.TextBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtDocumentName = New System.Windows.Forms.TextBox
        Me.btnParse = New System.Windows.Forms.Button
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(8, 48)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(472, 360)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtResults)
        Me.TabPage1.Controls.Add(Me.lblFindCount)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(464, 334)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "EMail Addresses"
        '
        'txtResults
        '
        Me.txtResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtResults.Location = New System.Drawing.Point(8, 40)
        Me.txtResults.Multiline = True
        Me.txtResults.Name = "txtResults"
        Me.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResults.Size = New System.Drawing.Size(448, 288)
        Me.txtResults.TabIndex = 2
        Me.txtResults.Text = ""
        '
        'lblFindCount
        '
        Me.lblFindCount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFindCount.Location = New System.Drawing.Point(8, 8)
        Me.lblFindCount.Name = "lblFindCount"
        Me.lblFindCount.Size = New System.Drawing.Size(448, 23)
        Me.lblFindCount.TabIndex = 1
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtDocContents)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(464, 334)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Document Content"
        '
        'txtDocContents
        '
        Me.txtDocContents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDocContents.Location = New System.Drawing.Point(8, 8)
        Me.txtDocContents.Multiline = True
        Me.txtDocContents.Name = "txtDocContents"
        Me.txtDocContents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDocContents.Size = New System.Drawing.Size(448, 320)
        Me.txtDocContents.TabIndex = 3
        Me.txtDocContents.Text = ""
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(400, 424)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.TabIndex = 1
        Me.btnExit.Text = "Exit"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Document:"
        '
        'txtDocumentName
        '
        Me.txtDocumentName.Location = New System.Drawing.Point(72, 8)
        Me.txtDocumentName.Name = "txtDocumentName"
        Me.txtDocumentName.Size = New System.Drawing.Size(264, 20)
        Me.txtDocumentName.TabIndex = 3
        Me.txtDocumentName.Text = ""
        '
        'btnParse
        '
        Me.btnParse.Location = New System.Drawing.Point(416, 8)
        Me.btnParse.Name = "btnParse"
        Me.btnParse.Size = New System.Drawing.Size(56, 23)
        Me.btnParse.TabIndex = 4
        Me.btnParse.Text = "Parse"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(352, 8)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(56, 23)
        Me.btnBrowse.TabIndex = 5
        Me.btnBrowse.Text = "Browse"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 462)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.btnParse)
        Me.Controls.Add(Me.txtDocumentName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form1"
        Me.Text = "Extract EMail Addresses from Word Documents"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub enableParseButton()
        btnParse.Enabled = (txtDocumentName.Text.Length > 0)
    End Sub

    Private Function ExtractEmailAddressesFromString(ByVal source As String) As String()
        Dim mc As MatchCollection
        Dim i As Integer

        ' expression garnered from www.regexlib.com - thanks guys!
        mc = Regex.Matches(source, "([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")
        Dim results(mc.Count - 1) As String
        For i = 0 To results.Length - 1
            results(i) = mc(i).Value
        Next

        Return results
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtDocumentName.Text = ""
        enableParseButton()
    End Sub

    Private Sub txtDocumentName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocumentName.TextChanged
        enableParseButton()
    End Sub

    Private Sub btnParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParse.Click
        ' Dim app As Word.Application
        ' Dim doc As Word.Document
        Dim app As Object
        Dim doc As Object
        Dim docFileName As String
        Dim docPath As String
        Dim contents As String

        Cursor.Current = Cursors.WaitCursor

        Try
            ' init UI controls
            lblFindCount.Text = ""
            txtResults.Text = ""
            txtDocContents.Text = ""

            ' validate file name
            docFileName = txtDocumentName.Text
            If docFileName.Length = 0 Then
                MsgBox("Please enter a file name")
                txtDocumentName.Focus()
                Return
            End If

            ' if no path use APP_BASE
            docPath = Path.GetDirectoryName(docFileName)
            If docPath.Length = 0 Then
                docFileName = Application.StartupPath & "\" & docFileName
            End If

            ' ensure file exists
            If Not File.Exists(docFileName) Then
                MsgBox("File does not exist")
                txtDocumentName.SelectAll()
                txtDocumentName.Focus()
                Return
            End If

            ' extract contents of file
            contents = ""
            If Path.GetExtension(docFileName).ToLower = ".txt" Then
                Dim fs As StreamReader

                Try
                    fs = New StreamReader(docFileName)
                    contents = fs.ReadToEnd
                Catch ex As Exception
                    MsgBox("Unable to read from text input file")
                    contents = ""
                Finally
                    If Not fs Is Nothing Then fs.Close()
                End Try
            Else
                Try
                    Try
                        'app = New Word.Application
                        app = CreateObject("Word.Application")
                    Catch ex As Exception
                        MsgBox("Unable to start Word")
                        Throw ex
                    End Try

                    Try
                        doc = app.Documents.Open(docFileName)
                    Catch ex As Exception
                        MsgBox("Unable to load document in Word")
                        Throw ex
                    End Try

                    contents = doc.Content.Text
                Catch ex As Exception
                    contents = ""
                Finally
                    If Not app Is Nothing Then app.Quit()
                End Try
            End If

            If contents.Length = 0 Then Return

            ' search for email addresses
            Dim emails As String()
            Dim email As String
            Dim results As New StringBuilder
            emails = ExtractEmailAddressesFromString(contents)
            For Each email In emails
                results.Append(email & vbNewLine)
            Next

            ' display results
            lblFindCount.Text = String.Format("{0} match(es) found.", emails.Length)
            txtResults.Text = results.ToString
            txtDocContents.Text = contents
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim ofd As OpenFileDialog

        Try
            ofd = New OpenFileDialog
            ofd.CheckFileExists = True
            ofd.CheckPathExists = True
            ofd.Filter = "Word Documents (*.doc)|*.doc|Rich Text Documents (*.rtf)|*.rtf|Text Documents (*.txt)|*.txt"
            ofd.Title = "Select Document"
            If ofd.ShowDialog = DialogResult.OK Then
                txtDocumentName.Text = ofd.FileName
            End If
        Finally
            If Not ofd Is Nothing Then ofd.Dispose()
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class
