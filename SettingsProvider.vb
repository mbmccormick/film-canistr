Public Class SettingsProvider

    Public Shared Sub LoadSettings()

        Dim settingsMatrix As New DataSet

        If System.IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\settings.xml") Then

            settingsMatrix.ReadXml(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\settings.xml")

            Main.NumericUpDown1.Value = settingsMatrix.Tables("settings").Rows(0).Item("NumericUpDown1")
            Main.ComboBox1.SelectedIndex = settingsMatrix.Tables("settings").Rows(0).Item("ComboBox1")
            Main.TextBox1.Text = settingsMatrix.Tables("settings").Rows(0).Item("TextBox1")
            Main.CheckBox1.Checked = settingsMatrix.Tables("settings").Rows(0).Item("CheckBox1")
            Main.CheckBox2.Checked = settingsMatrix.Tables("settings").Rows(0).Item("CheckBox2")
            Main.CheckBox3.Checked = settingsMatrix.Tables("settings").Rows(0).Item("CheckBox3")
            Main.RadioButton1.Checked = settingsMatrix.Tables("settings").Rows(0).Item("RadioButton1")
            Globals.RefreshTime = Convert.ToDateTime(settingsMatrix.Tables("settings").Rows(0).Item("RefreshTime"))

        End If

    End Sub

    Public Shared Sub SaveSettings()

        Dim settingsMatrix As New DataSet

        settingsMatrix.Tables.Add("settings")
        settingsMatrix.Tables("settings").Columns.Add("NumericUpDown1")
        settingsMatrix.Tables("settings").Columns.Add("ComboBox1")
        settingsMatrix.Tables("settings").Columns.Add("TextBox1")
        settingsMatrix.Tables("settings").Columns.Add("CheckBox1")
        settingsMatrix.Tables("settings").Columns.Add("CheckBox2")
        settingsMatrix.Tables("settings").Columns.Add("CheckBox3")
        settingsMatrix.Tables("settings").Columns.Add("RadioButton1")
        settingsMatrix.Tables("settings").Columns.Add("RefreshTime")

        settingsMatrix.Tables("settings").Rows.Add(New String() {Main.NumericUpDown1.Value, Main.ComboBox1.SelectedIndex, Main.TextBox1.Text, Main.CheckBox1.Checked, Main.CheckBox2.Checked, Main.CheckBox3.Checked, Main.RadioButton1.Checked, Globals.RefreshTime})

        If System.IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\settings.xml") Then
            System.IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\settings.xml")
        End If

        settingsMatrix.WriteXml(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\settings.xml")

    End Sub

End Class
