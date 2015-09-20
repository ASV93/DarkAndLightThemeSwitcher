'============================================================================
'
'    DarkAndLightThemeSwitcher
'    Copyright (C) 2015 Visual Software Corporation
'
'    Author: ASV93
'    File: Form1.vb
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'============================================================================

Imports Microsoft.Win32

Public Class Form1

    Dim DarkTheme As Boolean

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
            If Not DarkTheme Then
                Label2.Text = "Dark"
                SetTheme(False)
                DarkTheme = True
            End If
        Else 'From 9:00 AM to 18:59 PM
            If DarkTheme Then
                Label2.Text = "Light"
                SetTheme(True)
                DarkTheme = False
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim SrcFile As String = Application.ExecutablePath
            Dim DestFile As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\DNLTS.exe"
            If (SrcFile = DestFile) = False Then
                IO.File.Copy(SrcFile, DestFile, True)
            End If
            Dim regKey As RegistryKey
            regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", True)
            If regKey.GetValue("AppsUseLightTheme", 1) = 0 Then
                DarkTheme = True
            Else
                DarkTheme = False
            End If
        Catch ex As Exception
            MessageBox.Show("Critical Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetTheme(ByVal Light As Boolean)
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", True)
        If Light = True Then
            regKey.SetValue("AppsUseLightTheme", 1, RegistryValueKind.DWord) 'AppsUseLightTheme on > 10000; SystemUsesLightTheme on < 10000
        Else
            regKey.SetValue("AppsUseLightTheme", 0, RegistryValueKind.DWord)
        End If
        regKey.Close()
    End Sub

    Private Sub DarkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkToolStripMenuItem.Click
        Timer1.Enabled = False
        Label2.Text = "Dark"
        SetTheme(False)
        DarkTheme = True
    End Sub

    Private Sub LightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LightToolStripMenuItem.Click
        Timer1.Enabled = False
        Label2.Text = "Light"
        SetTheme(True)
        DarkTheme = False
    End Sub

    Private Sub DisableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisableToolStripMenuItem.Click
        Timer1.Interval = 2500
        Timer1.Enabled = True
    End Sub
End Class
