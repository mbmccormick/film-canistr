Imports FlickrNet
Imports System.Drawing
Imports System.Windows.Forms.PictureBox

Public Class ServiceProvider

#Region " Flickr Methods "

    Public Shared Sub GetPhotos(ByVal sender As Object)

        Main.NotifyIcon1.Text = "Film Canistr - Working..."

        Main.Cursor = Cursors.WaitCursor

        Globals.HasPhotoDetails = False

        If Not Main.TextBox1.Text.Length > 0 Then

            If sender Is Main.Button2 Then
                MessageBox.Show("You have not specified any photo preference tags! Please modify your settings.", "Film Canistr", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Main.NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "You have not specified any photo preference tags! Please modify your settings.", ToolTipIcon.Warning)
            End If

            Main.Cursor = Cursors.Default
            Exit Sub

        End If

        Dim randomGenerator As New Random
        Dim tagList As String = ""
        Dim resultImage As String = Nothing

        For Each item As String In Main.TextBox1.Text.Split(" ,")
            tagList = tagList & item & " "
        Next

        Try

            Dim f As Flickr = New Flickr("fe2d6808ec60eba68b35fd7b4f28e129", "0c29c45bf527c2b2")

            Dim options As PhotoSearchOptions = New PhotoSearchOptions()
            options.TagMode() = TagMode.AnyTag
            options.PerPage = 100

            Dim allPhotos As PhotoCollection

            options.Tags() = tagList.Trim

            allPhotos = f.PhotosSearch(options).PhotoCollection

            Dim tempPhoto As Photo
            Do
                tempPhoto = allPhotos(randomGenerator.Next(0, allPhotos.Count()))
            Loop While IsInHistory(tempPhoto.PhotoId)

            resultImage = tempPhoto.LargeUrl

        Catch ex As Exception

            If sender Is Main.Button2 Then
                MessageBox.Show("A connection to Flickr could not be established. Please check your Internet connection.", "Flickr Film Canistr", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Main.NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "A connection to Flickr could not be established. Please check your Internet connection.", ToolTipIcon.Warning)
            End If

            Main.Cursor = Cursors.Default
            Exit Sub

        End Try

        SystemProvider.DownloadFile(resultImage, "CurrentPhoto.jpg")

        Globals.CurrentPhotoDetails = GetToken(resultImage, 5, "/")

        Dim imagePath As String = ConvertImage("C:\WINDOWS\Web\Wallpaper\CurrentPhoto.jpg")

        If imagePath Is Nothing Then

            If sender Is Main.Button2 Then
                MessageBox.Show("Your search returned no results, please modify your photo preference settings!", "Flickr Film Canistr", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Main.NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "Your search returned no results, please modify your photo preference settings!", ToolTipIcon.Warning)
            End If

            Main.Cursor = Cursors.Default
            Exit Sub

        End If

        If IsErrorImage(imagePath) Then
            GetPhotos(sender)
            Exit Sub
        End If

        If Main.CheckBox3.Checked Then
            If Not ValidateSize(imagePath) Then
                GetPhotos(sender)
                Exit Sub
            End If
        End If

        SystemProvider.SetWallpaper(imagePath, Main.CheckBox2.Checked)
        SaveHistory(Globals.CurrentPhotoDetails)

        Main.Cursor = Cursors.Default

        Main.NotifyIcon1.ShowBalloonTip(10, "Film Canistr", "Your desktop wallpaper has been updated.", ToolTipIcon.Info)
        Main.NotifyIcon1.Text = "Film Canistr"

    End Sub

#End Region

#Region " History Methods "

    Public Shared Function IsInHistory(ByVal photoID As String) As Boolean
        Dim tempLine As String
        Dim fs As IO.StreamReader
        fs = IO.File.OpenText(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\history.dat")
        tempLine = fs.ReadLine()
        While Not tempLine Is Nothing AndAlso tempLine.Length > 0
            If photoID = tempLine Then
                Return True
            End If
            tempLine = fs.ReadLine()
        End While
        fs.Close()
        fs.Dispose()
        Return False
    End Function

    Public Shared Sub SaveHistory(ByVal photoID As String)
        Dim fs As IO.StreamWriter
        Try
            fs = IO.File.AppendText(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\history.dat")
            fs.WriteLine(photoID)
            fs.Close()
            fs.Dispose()
        Catch ex As Exception
            ' Do nothing
        End Try
    End Sub

#End Region

#Region " Utility Methods "

    Private Shared Function ConvertImage(ByVal jpgImagePath As String) As String

        Dim image As Bitmap
        Dim result As String

        Try
            image = New Bitmap(jpgImagePath)
            image.Save(jpgImagePath & ".bmp", System.Drawing.Imaging.ImageFormat.Bmp)
            image.Dispose()
        Catch ex As Exception
            image.Dispose()
            Return Nothing
        End Try

        image.Dispose()

        result = jpgImagePath & ".bmp"

        Try
            System.IO.File.Delete(jpgImagePath)
        Catch ex As Exception
            Return False
        End Try

        Return result

    End Function

    Public Shared Function GetToken(ByVal strVal As String, ByVal intIndex As Integer, ByVal strDelimiter As String) As String
        Dim strSubString() As String
        Dim intIndex2 As Integer
        Dim i As Integer
        Dim intDelimitLen As Integer

        intIndex2 = 1
        i = 0
        intDelimitLen = Len(strDelimiter)

        Do While intIndex2 > 0

            ReDim Preserve strSubString(i + 1)

            intIndex2 = InStr(1, strVal, strDelimiter)

            If intIndex2 > 0 Then
                strSubString(i) = Mid(strVal, 1, (intIndex2 - 1))
                strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
            Else
                strSubString(i) = strVal
            End If

            i = i + 1

        Loop

        If intIndex > (i + 1) Or intIndex < 1 Then
            GetToken = ""
        Else
            GetToken = strSubString(intIndex - 1)
        End If
    End Function

#End Region

#Region " Validation Methods "

    Private Shared Function IsErrorImage(ByVal imagePath As String) As Boolean

        Dim image As New Bitmap(imagePath)

        Dim height = image.Height
        Dim width = image.Width

        image.Dispose()

        If height = 375 And width = 500 Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Shared Function ValidateSize(ByVal imagePath As String) As Boolean

        Dim image As New Bitmap(imagePath)

        Dim screenResolution = My.Computer.Screen.Bounds.Height / My.Computer.Screen.Bounds.Width
        Dim imageResolution = image.Height / image.Width

        image.Dispose()

        'MessageBox.Show("screenResolution = " & screenResolution & vbNewLine & _
        '                "imageResolution = " & imageResolution & vbNewLine & _
        '                vbNewLine & _
        '                "upper bound = " & screenResolution + (screenResolution * 0.25) & vbNewLine & _
        '                "lower bound = " & screenResolution - (screenResolution * 0.25))


        If (imageResolution > (screenResolution - (screenResolution * 0.25))) AndAlso _
            (imageResolution < (screenResolution + (screenResolution * 0.25))) Then
            Return True
        Else
            Return False
        End If

    End Function

#End Region

End Class
