Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False
    Private imagecontroller As Integer = 8
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Private letterlist As ArrayList
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReScanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReScanToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AssignDriveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Private initialsteps As Boolean

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String, ByVal statustag As Integer)


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        letterlist = New ArrayList
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FullError As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents flashresult As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Screen))
        Me.Label1 = New System.Windows.Forms.Label
        Me.FullError = New System.Windows.Forms.CheckBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.flashresult = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ReScanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AssignDriveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ReScanToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.3!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(510, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "BUILD 20050819.2"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Visible = False
        '
        'FullError
        '
        Me.FullError.BackColor = System.Drawing.Color.Transparent
        Me.FullError.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FullError.Location = New System.Drawing.Point(606, 27)
        Me.FullError.Name = "FullError"
        Me.FullError.Size = New System.Drawing.Size(16, 16)
        Me.FullError.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.FullError, "If checked, full error reporting mode is initiated.")
        Me.FullError.UseVisualStyleBackColor = False
        Me.FullError.Visible = False
        '
        'TreeView1
        '
        Me.TreeView1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TreeView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeView1.ImageIndex = 0
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Location = New System.Drawing.Point(454, 128)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 1
        Me.TreeView1.Size = New System.Drawing.Size(168, 160)
        Me.TreeView1.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.TreeView1, "Displays the current drives found on your system.")
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        Me.ImageList1.Images.SetKeyName(5, "")
        Me.ImageList1.Images.SetKeyName(6, "")
        Me.ImageList1.Images.SetKeyName(7, "")
        Me.ImageList1.Images.SetKeyName(8, "")
        Me.ImageList1.Images.SetKeyName(9, "")
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Silver
        Me.Button1.Enabled = False
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(328, 121)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Assign Drive"
        Me.ToolTip1.SetToolTip(Me.Button1, "Assign your USB disk to the selected Unassigned letter above")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Silver
        Me.Button2.Enabled = False
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(570, 96)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(50, 18)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "RE:SCAN"
        Me.ToolTip1.SetToolTip(Me.Button2, "Restart operation")
        Me.Button2.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.Color.White
        Me.ComboBox1.Enabled = False
        Me.ComboBox1.ForeColor = System.Drawing.Color.Black
        Me.ComboBox1.Location = New System.Drawing.Point(48, 96)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(144, 21)
        Me.ComboBox1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.ComboBox1, "Select an unassigned drive letter from this list.")
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(9, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(613, 64)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(9, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(538, 40)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Please remember to REMOVE your USB Flash Disk once you finished using this workst" & _
            "ation. Commerce I.T. cannot be held liable for lost or stolen property."
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.flashresult)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(12, 128)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(430, 160)
        Me.Panel1.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(144, 16)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Assign New Drive Letter:"
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Green
        Me.Label8.Location = New System.Drawing.Point(144, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(272, 16)
        Me.Label8.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(48, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 16)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Drive Letter:"
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Green
        Me.Label6.Location = New System.Drawing.Point(144, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(272, 16)
        Me.Label6.TabIndex = 6
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(48, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 16)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Caption:"
        '
        'flashresult
        '
        Me.flashresult.ForeColor = System.Drawing.Color.Green
        Me.flashresult.Location = New System.Drawing.Point(144, 8)
        Me.flashresult.Name = "flashresult"
        Me.flashresult.Size = New System.Drawing.Size(128, 16)
        Me.flashresult.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(144, 16)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Searching For Flash Disk:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.DimGray
        Me.Label5.Location = New System.Drawing.Point(12, 296)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(560, 16)
        Me.Label5.TabIndex = 9
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(604, 294)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(18, 18)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 2000
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Enabled = False
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReScanToolStripMenuItem, Me.HelpToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(632, 24)
        Me.MenuStrip1.TabIndex = 57
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(40, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ReScanToolStripMenuItem
        '
        Me.ReScanToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReScanToolStripMenuItem1, Me.AssignDriveToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ReScanToolStripMenuItem.Name = "ReScanToolStripMenuItem"
        Me.ReScanToolStripMenuItem.Size = New System.Drawing.Size(72, 20)
        Me.ReScanToolStripMenuItem.Text = "Operations"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'AssignDriveToolStripMenuItem
        '
        Me.AssignDriveToolStripMenuItem.Enabled = False
        Me.AssignDriveToolStripMenuItem.Name = "AssignDriveToolStripMenuItem"
        Me.AssignDriveToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AssignDriveToolStripMenuItem.Text = "Assign Drive"
        '
        'ReScanToolStripMenuItem1
        '
        Me.ReScanToolStripMenuItem1.Name = "ReScanToolStripMenuItem1"
        Me.ReScanToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ReScanToolStripMenuItem1.Text = "Re-Scan"
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(632, 320)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.FullError)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(638, 352)
        Me.MinimumSize = New System.Drawing.Size(638, 352)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Assign USB Flash Disk Drive Letter"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            workerbusy = False
            button_enable(True)
            SendMessage("Label5", "Error Eoncountered")

            Dim Display_Message1 As New Display_Message()
            Display_Message1.Message_Textbox.Text = message
            Display_Message1.ShowDialog()

        Catch ex As Exception
            MsgBox("An error occurred in Assign USB Flash Disk Drive Letter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub




    Private Sub Error_Handler(ByVal except As Exception, Optional ByVal inmessage As String = "")
        Try
            workerbusy = False
            SendMessage("Label5", "Error Eoncountered")
            button_enable(True)

            Dim message As String
            If FullError.Checked = True Then
                message = "Exception Encountered:" & vbCrLf & vbCrLf & except.ToString
            Else
                message = "Exception Encountered:" & vbCrLf & vbCrLf & except.Message.ToString
            End If

            If Not (except.GetType.Name = "ThreadAbortException") Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = message
                Display_Message1.ShowDialog()
            End If

            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Try
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & inmessage & ": " & except.ToString)
                filewriter.WriteLine()
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            Catch ex As Exception
                SendMessage("Label5", "Error Eoncountered")
            End Try
        Catch ex As Exception
            MsgBox(ex.ToString)
            MsgBox("An error occurred in Assign USB Flash Disk Drive Letter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Protected Sub SendMessage(ByVal labelname As String, ByVal message As String)
        Try
            Dim controllist As ControlCollection = Me.Controls
            Dim cont As Control

            For Each cont In controllist
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
        Catch ex As Exception
            If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                Error_Handler(message & vbCrLf & labelname & vbCrLf & Worker1.WorkerThread.ThreadState.ToString)
                Error_Handler(ex)
            End If
        End Try
    End Sub


    Public Sub WorkerStatusMessageHandler(ByVal message As String, ByVal statustag As Integer)
        Try
            If statustag = 1 Then
                SendMessage("Label5", message)
            Else
                SendMessage("Label5", message)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception)
        Try
            Error_Handler(Message)
            workerbusy = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub




    Protected Sub WorkerCompleteHandler(ByVal found As Integer)
        'found = 0 not found
        'found = 1 found
        Try
            Dim eventhandled As Boolean = False



            If found = 1 Then
                flashresult.Text = "OK"
                flashresult.ForeColor = Color.Green
                Label6.Text = Worker1.drive_caption
                Label6.Refresh()
                Label8.ForeColor = Color.Green
                Label8.Text = Worker1.drive_letter
                If Label8.Text = "" Then
                    Label8.Text = "Unassigned"
                    Label8.ForeColor = Color.OrangeRed
                End If
                Label8.Refresh()
            Else
                flashresult.Text = "FAIL"
                flashresult.ForeColor = Color.Red
            End If
            flashresult.Refresh()
            workerbusy = False
            button_enable(True)

            eventhandled = True


        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub FillTree(ByVal Location As String)
        Try
            workerbusy = True
            button_enable(False)
            SendMessage("Label5", "Loading Current Drive List")
            ComboBox1.Items.Clear()
            ComboBox1.Refresh()
            'images
            '0 - stiffy
            '1 - hard
            '2 - network
            '3 - cd
            '4 - flash

            TreeView1.Nodes.Clear()
            Dim topnode As TreeNode = TreeView1.Nodes.Add("Current Drives")
            topnode.ImageIndex = 5
            topnode.SelectedImageIndex = 5
            'FIND ALL DRIVES

            Dim runner As IEnumerator
            Dim fso As New Scripting.FileSystemObject
            runner = fso.Drives.GetEnumerator()

            While runner.MoveNext() = True
                Dim d As Scripting.Drive
                d = runner.Current()

                Select Case d.DriveType
                    Case Scripting.DriveTypeConst.CDRom
                        Dim name As String
                        If d.IsReady = True Then
                            name = d.DriveLetter & ": " & d.VolumeName
                        Else
                            name = d.DriveLetter & ":"
                        End If
                        Dim tnode As TreeNode = topnode.Nodes.Add(name)
                        tnode.Tag = name
                        tnode.ImageIndex = 3
                        tnode.SelectedImageIndex = 3

                    Case Scripting.DriveTypeConst.Fixed
                        Dim name As String
                        If d.IsReady = True Then
                            name = d.DriveLetter & ": " & d.VolumeName
                        Else
                            name = d.DriveLetter & ":"
                        End If
                        Dim tnode As TreeNode = topnode.Nodes.Add(name)
                        tnode.Tag = name
                        tnode.ImageIndex = 1
                        tnode.SelectedImageIndex = 1

                    Case Scripting.DriveTypeConst.Removable
                        Dim name As String
                        If d.IsReady = True Then
                            name = d.DriveLetter & ": " & d.VolumeName
                        Else
                            name = d.DriveLetter & ":"
                        End If
                        Dim tnode As TreeNode = topnode.Nodes.Add(name)
                        tnode.Tag = name
                        If d.DriveLetter.ToLower = "a" Then
                            tnode.ImageIndex = 0
                            tnode.SelectedImageIndex = 0
                        Else
                            tnode.ImageIndex = 4
                            tnode.SelectedImageIndex = 4
                        End If

                    Case Scripting.DriveTypeConst.Remote
                        Dim name As String
                        If d.IsReady = True Then
                            Dim cinfo As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
                            name = d.DriveLetter & ": " & cinfo.TextInfo.ToTitleCase(d.ShareName.Substring(d.ShareName.LastIndexOf("\")).ToLower)
                        Else
                            name = d.DriveLetter & ":"
                        End If
                        Dim tnode As TreeNode = topnode.Nodes.Add(name)
                        tnode.Tag = name
                        tnode.ImageIndex = 2
                        tnode.SelectedImageIndex = 2
                    Case Else
                        Dim name As String
                        If d.IsReady = True Then
                            name = d.DriveLetter & ": " & d.VolumeName
                        Else
                            name = d.DriveLetter & ":"
                        End If
                        Dim tnode As TreeNode = topnode.Nodes.Add(name)
                        tnode.Tag = name
                        tnode.ImageIndex = 1
                        tnode.SelectedImageIndex = 1

                End Select

            End While
            TreeView1.ExpandAll()

            populate_letterlist()
            remove_existing_letterlist()
            populate_combobox()

            workerbusy = False
            button_enable(True)
            SendMessage("Label5", "Current Drive List Loaded")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub USBSearch()
        Try
            workerbusy = True
            button_enable(False)
            flashresult.Text = ""
            flashresult.Refresh()
            Label6.Text = ""
            Label6.Refresh()
            Label8.Text = ""
            Label8.Refresh()
            Worker1.ChooseThreads(1)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub worksteps()
        Try
            workerbusy = True
            imagecontroller = 7
            PictureBox1.Image = ImageList1.Images(imagecontroller)
            PictureBox1.Refresh()
            flashresult.Text = ""
            flashresult.Refresh()
            Label6.Text = ""
            Label6.Refresh()
            Label8.Text = ""
            Label8.Refresh()
            FillTree("")
            USBSearch()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            initialsteps = False
            Me.Text = My.Application.Info.ProductName & " " & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            Label1.Text = "BUILD " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            SendMessage("Label5", "Initializing Drive Scan...")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If workerbusy = False Then
                imagecontroller = 9
                PictureBox1.Image = ImageList1.Images(imagecontroller)
            Else
                If imagecontroller = 7 Then
                    imagecontroller = 6
                    PictureBox1.Image = ImageList1.Images(imagecontroller)
                Else
                    imagecontroller = 7
                    PictureBox1.Image = ImageList1.Images(imagecontroller)
                End If
            End If
            PictureBox1.Refresh()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub button_enable(ByVal enable As Boolean)
        Try
            Select Case enable
                Case True
                    If flashresult.Text = "OK" Then
                        Button1.Enabled = True
                        ComboBox1.Enabled = True
                        AssignDriveToolStripMenuItem.Enabled = True
                    Else
                        Button1.Enabled = False
                        AssignDriveToolStripMenuItem.Enabled = False
                    End If
                    Button2.Enabled = True
                    MenuStrip1.Enabled = True
                Case False
                    Button1.Enabled = False
                    ComboBox1.Enabled = False
                    Button2.Enabled = False
                    MenuStrip1.Enabled = False
            End Select
            Button1.Refresh()
            Button2.Refresh()
            ComboBox1.Refresh()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub

    Private Sub main_screen_closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        worksteps()
    End Sub

    Private Sub populate_combobox()
        Try
            ComboBox1.Items.Clear()
            Dim i As Integer
            For i = 0 To letterlist.Count - 1
                ComboBox1.Items.Add(letterlist.Item(i).ToString())
            Next
            ComboBox1.SelectedIndex = 0
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub remove_existing_letterlist()
        Try
            Dim runner As IEnumerator
            Dim fso As New Scripting.FileSystemObject
            runner = fso.Drives.GetEnumerator()
            While runner.MoveNext() = True
                Dim d As Scripting.Drive
                d = runner.Current()
                letterlist.RemoveAt((letterlist.IndexOf(d.DriveLetter.ToString)))
            End While
            If letterlist.IndexOf("A") <> -1 Then
                letterlist.RemoveAt(letterlist.IndexOf("A"))
            End If
            If letterlist.IndexOf("B") <> -1 Then
                letterlist.RemoveAt(letterlist.IndexOf("B"))
            End If
            If letterlist.IndexOf("C") <> -1 Then
                letterlist.RemoveAt(letterlist.IndexOf("C"))
            End If
            If letterlist.IndexOf("D") <> -1 Then
                letterlist.RemoveAt(letterlist.IndexOf("D"))
            End If
            If letterlist.IndexOf("E") <> -1 Then
                letterlist.RemoveAt(letterlist.IndexOf("E"))
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub populate_letterlist()
        Try
            letterlist.Clear()
            Dim i As Integer
            For i = 65 To 90
                letterlist.Add(Chr(i).ToString)
            Next
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub USBAssign()
        Try
            workerbusy = True
            button_enable(False)
            Worker1.selectedletter = ComboBox1.SelectedItem.ToString()
            Worker1.volume = 3
            Worker1.ChooseThreads(2)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            USBAssign()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If initialsteps = False Then
            initialsteps = True
            worksteps()
        End If

    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AssignDriveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssignDriveToolStripMenuItem.Click
        Try
            USBAssign()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub ReScanToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReScanToolStripMenuItem1.Click
        worksteps()
    End Sub
End Class
