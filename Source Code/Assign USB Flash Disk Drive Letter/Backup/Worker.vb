Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public drive_caption As String
    Public drive_type As String
    Public drive_letter As String
    Public volume As Integer
    Public selectedletter As String
    
    Private filereader As StreamReader
    Private filewriter As StreamWriter

    Public Event WorkerFileProcessing(ByVal filename As String, ByVal queue As Integer)
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerFileCount(ByVal Result As Long, ByVal count As Integer)
    Public Event WorkerComplete(ByVal found As Integer)




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in Assign USB Flash Disk Drive Letter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf FlashDiskFind_Routine)
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf FlashDiskAssign_Routine)
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



   

    Private Sub FlashDiskFind_Routine()
        RaiseEvent WorkerStatusMessage("Searching for USB Flash Disk", 1)
        Dim apptorun As String
        apptorun = """" & Application.StartupPath & "\cscript.exe"" """ & Application.StartupPath & "\disksearch.vbs"" > """ & Application.StartupPath & "\output.txt"""

        DosShellCommand(apptorun)

        filereader = New StreamReader(Application.StartupPath & "\output.txt", True)

        Dim line As String
        line = filereader.ReadLine
        Dim dcaption As String = ""
        Dim dtype As String = ""
        Dim dletter As String = ""

        volume = 0
        While Not IsNothing(line) = True
            If line.StartsWith("Disk drive Caption:") = True Then
                If dtype = "" And dletter = "" And Not dcaption = "" Then
                    dtype = 2
                    Exit While
                End If
                If dtype.Length > 0 Then
                    If CInt(dtype) = 2 Then
                        Exit While
                    End If
                End If
                volume = volume + 1
                dcaption = ""
                dtype = ""
                dletter = ""
                dcaption = line.Replace("Disk drive Caption: ", "")
            End If
            If line.StartsWith("Drive Type:") = True Then
                dtype = line.Replace("Drive Type: ", "")
            End If
            If line.StartsWith("Associated Drive letter:") = True Then
                dletter = line.Replace("Associated Drive letter: ", "")
            End If

            line = filereader.ReadLine
        End While

        filereader.Close()

        drive_caption = dcaption
        drive_type = dtype
        drive_letter = dletter


        If dtype = "" And dletter = "" And Not dcaption = "" Then
            dtype = 2
        End If

        If dtype.Length > 0 Then
            If CInt(dtype) = 2 Then
                RaiseEvent WorkerComplete(1)
            Else
                RaiseEvent WorkerComplete(0)
            End If
        Else
            RaiseEvent WorkerComplete(0)
        End If
        RaiseEvent WorkerStatusMessage("USB Flash Disk Search Concluded", 1)
    End Sub

    
    Private Sub FlashDiskAssign_Routine()
        RaiseEvent WorkerStatusMessage("Assigning new Drive Letter to USB Flash Disk", 1)


        filewriter = New StreamWriter(Application.StartupPath & "\input.txt", False)

        filewriter.WriteLine("select volume " & volume)
        filewriter.WriteLine("assign letter=" & selectedletter)
        filewriter.Close()
        
        Dim apptorun As String
        apptorun = """" & Application.StartupPath & "\diskpart.exe"" /s """ & Application.StartupPath & "\input.txt"""
        DosShellCommand(apptorun)

        DosShellCommand("explorer " & selectedletter & ":\")

        FlashDiskFind_Routine()



        RaiseEvent WorkerStatusMessage("Assignment of Drive Letter to USB Flash Disk Concluded", 1)
    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function

End Class
