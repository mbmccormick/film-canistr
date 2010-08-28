Imports Microsoft.Win32
Imports System.Runtime.InteropServices

Public Class SystemProvider

    Const SPI_SETDESKWALLPAPER As Integer = 20
    Const SPIF_UPDATEINIFILE As Integer = &H1&
    Const SPIF_SENDWININICHANGE As Integer = &H2&

    <DllImport("user32")> _
    Public Shared Function SystemParametersInfo(ByVal uAction As Integer, _
        ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer
    End Function

    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer

    Public Shared Function SetWallpaper(ByVal path As String, ByVal save As Boolean) As Boolean

        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", SPIF_UPDATEINIFILE)

        ' Begin Write to registry and set wallpaper stuff...
        Dim key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", True)

        key.SetValue("WallpaperStyle", "2")
        key.SetValue("TileWallpaper", "0")

        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", SPIF_UPDATEINIFILE)
        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE)

        If save = True Then
            If IO.File.Exists("C:\WINDOWS\Web\Wallpaper\CurrentPhoto.jpg.bmp") Then
                IO.File.Copy("C:\WINDOWS\Web\Wallpaper\CurrentPhoto.jpg.bmp", "C:\WINDOWS\Web\Wallpaper\" & Guid.NewGuid.ToString & ".bmp")
            End If
        End If

        Return True

    End Function

    Public Shared Function DownloadFile(ByVal url As String, ByVal fileName As String) As Boolean

        Dim retVal As Integer
        Dim theUrl As String = url
        Dim savePath As String = "C:\WINDOWS\Web\Wallpaper\" & fileName

        retVal = URLDownloadToFile(0, theUrl, savePath, 0, 0)

        If retVal = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function CheckDependencies() As Boolean

        If Not System.IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\history.dat") Then
            System.IO.File.Create(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\history.dat")
        End If

        If Not System.IO.File.Exists(Globals.WorkingDirectory & "FlickrNet.dll") Then
            MessageBox.Show("One or more of Film Canistr's dependencies are missing. Please reinstall Film Canistr to resolve this issue.", "Dependency Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Return True

    End Function

    Public Shared Sub AddToStartup()
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.SetValue("Film Canistr", Application.ExecutablePath.ToString())
        regKey.Close()
    End Sub

    Public Shared Sub RemoveFromStartup()
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
        regKey.DeleteValue("Film Canistr", False)
        regKey.Close()
    End Sub

End Class
