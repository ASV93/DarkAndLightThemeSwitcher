Imports Microsoft.Win32

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NotifyIcon1.Dispose()
        End
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        NotifyIcon1.Dispose()
        End
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MessageBox.Show("Created by ASV93", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Hide()
        Timer1.Interval = 60000
        If Now.Hour >= 19 Or Now.Hour < 9 Then
            Label2.Text = "Dark"
            SetTheme(False)
        Else 'From 9:00 AM to 18:59 PM
            Label2.Text = "Light"
            SetTheme(True)
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim SrcFile As String = Application.ExecutablePath
            Dim DestFile As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\DNLTS.exe"
            If (SrcFile = DestFile) = False Then
                IO.File.Copy(SrcFile, DestFile, True)
            End If
        Catch ex As Exception
            MessageBox.Show("Critical Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetTheme(ByVal Light As Boolean)
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", True)
        If Light = True Then
            regKey.SetValue("SystemUsesLightTheme", 1, RegistryValueKind.DWord)
        Else
            regKey.SetValue("SystemUsesLightTheme", 0, RegistryValueKind.DWord)
        End If
        regKey.Close()
    End Sub
End Class
