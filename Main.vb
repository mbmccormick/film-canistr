Imports FlickrNet

Public Class Main

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        e.Cancel = True
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False

        Globals.WorkingDirectory = Application.ExecutablePath.Replace(System.IO.Path.GetFileName(Application.ExecutablePath), "")

        If Not SystemProvider.CheckDependencies() Then
            End
        End If

        SettingsProvider.LoadSettings()
        Me.Button3.Focus()

        Me.Timer1.Interval = 60000
        Me.Timer1.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ServiceProvider.GetPhotos(Me.Button2)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not Me.ComboBox1.Text.Length = 0 Then
            Select Case Me.ComboBox1.SelectedIndex
                Case 0
                    Globals.RefreshTime = Now.AddMinutes(Me.NumericUpDown1.Value)
                Case 1
                    Globals.RefreshTime = Now.AddHours(Me.NumericUpDown1.Value)
                Case 2
                    Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value)
                Case 3
                    Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value * 7)
                Case 4
                    Globals.RefreshTime = Now.AddMonths(Me.NumericUpDown1.Value)
            End Select
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Not Globals.RefreshTime = Nothing Then
            If Now >= Globals.RefreshTime Then
                ServiceProvider.GetPhotos(Me.Timer1)

                If Not Me.ComboBox1.Text.Length = 0 Then
                    Select Case Me.ComboBox1.SelectedIndex
                        Case 0
                            Globals.RefreshTime = Now.AddMinutes(Me.NumericUpDown1.Value)
                        Case 1
                            Globals.RefreshTime = Now.AddHours(Me.NumericUpDown1.Value)
                        Case 2
                            Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value)
                        Case 3
                            Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value * 7)
                        Case 4
                            Globals.RefreshTime = Now.AddMonths(Me.NumericUpDown1.Value)
                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        If Not Me.ComboBox1.Text.Length = 0 Then
            Select Case Me.ComboBox1.SelectedIndex
                Case 0
                    Globals.RefreshTime = Now.AddMinutes(Me.NumericUpDown1.Value)
                Case 1
                    Globals.RefreshTime = Now.AddHours(Me.NumericUpDown1.Value)
                Case 2
                    Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value)
                Case 3
                    Globals.RefreshTime = Now.AddDays(Me.NumericUpDown1.Value * 7)
                Case 4
                    Globals.RefreshTime = Now.AddMonths(Me.NumericUpDown1.Value)
            End Select
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SettingsProvider.SaveSettings()
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        SettingsProvider.SaveSettings()
        End
    End Sub

    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        If Globals.HasPhotoDetails = True Then
            CreateObject("WScript.Shell").Run(Globals.CurrentPhotoURL)
        End If
    End Sub

    Private Sub NotifyIcon1_BalloonTipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClosed
        Globals.HasPhotoDetails = False
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        ServiceProvider.GetPhotos(Me.RefreshToolStripMenuItem)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MessageBox.Show("Film Canistr Version 2.0" & vbNewLine & _
                        "Copyright © mbmccormick software. All rights reserved." & vbNewLine & _
                        "Licensed under the GNU General Public License." & vbNewLine & _
                        vbNewLine & _
                        "Next refresh at " & Globals.RefreshTime & ".", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked Then
            Me.SaveToolStripMenuItem.Enabled = False
        Else
            Me.SaveToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If IO.File.Exists("C:\WINDOWS\Web\Wallpaper\CurrentPhoto.jpg.bmp") Then
            IO.File.Copy("C:\WINDOWS\Web\Wallpaper\CurrentPhoto.jpg.bmp", "C:\WINDOWS\Web\Wallpaper\" & Guid.NewGuid.ToString & ".bmp")
            Me.NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "This photo has been saved to your wallpaper directory.", ToolTipIcon.Info)
        End If
    End Sub

    Private Sub PhotoDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PhotoDetailsToolStripMenuItem.Click

        If Globals.CurrentPhotoDetails = "" Then
            NotifyIcon1.BalloonTipText = "Unable to retrieve details for this photo, refresh the desktop wallpaper and try again."
            Globals.HasPhotoDetails = False
            NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "Unable to retrieve details for this photo, refresh the desktop wallpaper and try again.", ToolTipIcon.Warning)
            Return
        End If

        Dim id As String
        id = ServiceProvider.GetToken(Globals.CurrentPhotoDetails, 1, "_")
        Dim f As Flickr = New Flickr("fe2d6808ec60eba68b35fd7b4f28e129", "0c29c45bf527c2b2")

        Globals.CurrentPhotoURL = f.PhotosGetInfo(id).WebUrl
        Globals.HasPhotoDetails = True

        If f.PhotosGetInfo(id).Owner.RealName = "" Then
            If f.PhotosGetInfo(id).Title = "" Then
                NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "This photo, (unknown), was taken by (unknown) on " & f.PhotosGetInfo(id).Dates.TakenDate.Date & ". Click here to see more information about this photo.", ToolTipIcon.Info)
            Else
                NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "This photo, " & f.PhotosGetInfo(id).Title & ", was taken by (unknown) on " & f.PhotosGetInfo(id).Dates.TakenDate.Date & ". Click here to see more information about this photo.", ToolTipIcon.Info)
            End If
        Else
            If f.PhotosGetInfo(id).Title = "" Then
                NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "This photo, (unknown), was taken by " & f.PhotosGetInfo(id).Owner.RealName & " on " & f.PhotosGetInfo(id).Dates.TakenDate.Date & ". Click here to see more information about this photo.", ToolTipIcon.Info)
            Else
                NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "This photo, " & f.PhotosGetInfo(id).Title & ", was taken by " & f.PhotosGetInfo(id).Owner.RealName & " on " & f.PhotosGetInfo(id).Dates.TakenDate.Date & ". Click here to see more information about this photo.", ToolTipIcon.Info)
            End If
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If Me.RadioButton1.Checked Then
            Me.RadioButton2.Checked = False
        End If

        Me.NumericUpDown1.Enabled = True
        Me.ComboBox1.Enabled = True
        Me.Timer1.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If Me.RadioButton2.Checked Then
            Me.RadioButton1.Checked = False
        End If

        Me.NumericUpDown1.Enabled = False
        Me.ComboBox1.Enabled = False
        Me.Timer1.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked Then
            SystemProvider.AddToStartup()
        Else
            SystemProvider.RemoveFromStartup()
        End If
    End Sub
End Class
