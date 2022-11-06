Imports System.Drawing.Imaging
Imports System.Text
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Drawing
Imports System.IO
Imports AnimatedGif
Imports System.Net.Security

Public Class Form1

    'SpriteManager (JASON Halloween Version)
    '30-10-2022
    'lord.marte@gmail.com

    'https://social.msdn.microsoft.com/Forums/en-US/a6eb8a7f-6e4b-4ad3-986d-8f43d3bf48a6/gif-image?forum=vbgeneral
    'AnimatedGIf
    '// 33ms delay (~30fps)
    'Using gif = AnimatedGif.Create("gif.gif", 33)
    '   Dim img = Image.FromFile("img.png")
    '   gif.AddFrame(img, delay:=-1, quality:=GifQuality.Bit8)
    'End Using

    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Int32, ByVal yPoint As Int32) As Int32
    Private Declare Function GetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As Int32, ByVal x As Int32, ByVal y As Int32) As Int32

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowDC(ByVal hwnd As IntPtr) As IntPtr
        'Do not try to name this method "GetDC" it will say that user32 doesnt have GetDC !!!
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ReleaseDC(ByVal hwnd As IntPtr, ByVal hdc As IntPtr) As Int32

    End Function

    <DllImport("gdi32.dll", SetLastError:=True)>
    Public Shared Function GetPixel(ByVal hdc As IntPtr, ByVal nXPos As Integer, ByVal nYPos As Integer) As UInteger

    End Function

    Public cartel As String = Application.StartupPath.ToString    'dir save
    Public Flybox As Integer = 0    'rect on mouse pos when right button pressed
    Public JASONarray As Array
    Public win32colorArray As Array
    Public CSVstring As String = ""
    Public LibraryPos As Integer = 0

    'GIFs
    Public IntervalGIF As Integer = 500     'milliseconds


    Enum PropertyID
        PropertyTagDateTime = 306
        PropertyTagExifDTOrig = 36867
        PropertyTagEquipMake = 271
        PropertyTagEquipModel = 272
    End Enum

    'Dim oCurrentPictureBox As PictureBox = Nothing
    'Public Overridable Property Palette As ColorPalette





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'apri
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "All files images|*.jpg;*.bmp;*.png;*.gif"
        OpenFileDialog1.Title = "open image file..."
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
            PictureBox1.Size = New Size(1920, 1080)

            CreatePalette()

            'Mouse e Anteprima
            Timer1.Start()
            Timer2.Start()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub






    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'mouse pos
        Dim win As Int32
        Dim rgb As Int32
        Dim hdc As Int32

        win = WindowFromPoint(Windows.Forms.Cursor.Position.X, Windows.Forms.Cursor.Position.Y)
        Dim g As Graphics = Graphics.FromHwnd(New IntPtr(win))
        hdc = g.GetHdc.ToInt32

        rgb = GetPixel(hdc, Windows.Forms.Cursor.Position.X, Windows.Forms.Cursor.Position.Y)

        If MouseButtons.HasFlag(MouseButtons.Left) Then
            TextBox5.Text = Windows.Forms.Cursor.Position.X
            TextBox6.Text = Windows.Forms.Cursor.Position.Y
        Else
            TextBox3.Text = Windows.Forms.Cursor.Position.X
            TextBox4.Text = Windows.Forms.Cursor.Position.Y
        End If

        'capture
        If Flybox = 1 And MouseButtons.HasFlag(MouseButtons.Left) And MouseButtons.HasFlag(MouseButtons.Right) Then
            Dim x As Integer = CInt(TextBox3.Text)
            Dim y As Integer = CInt(TextBox4.Text)
            Dim boxSize = CInt(Windows.Forms.Cursor.Position.X) - CInt(x)
            Dim boxSize2 = CInt(Windows.Forms.Cursor.Position.Y) - CInt(y)

            Dim SC As New ScreenShot.ScreenCapture
            Dim MyBitMap As Bitmap = SC.CaptureDeskTopRectangle(New Rectangle(x, y, boxSize, boxSize2), boxSize, boxSize2)
            MyBitMap.Save(cartel & "\CaptureRegion.jpg", Imaging.ImageFormat.Jpeg)
            If PictureBox4.Image Is Nothing Then
                PictureBox4.Image = MyBitMap
                PictureBox7.Image = Nothing
            Else
                If PictureBox5.Image Is Nothing Then
                    PictureBox5.Image = MyBitMap
                Else
                    If PictureBox6.Image Is Nothing Then
                        PictureBox6.Image = MyBitMap
                    Else
                        If PictureBox7.Image Is Nothing Then
                            PictureBox7.Image = MyBitMap
                        Else
                            PictureBox4.Image = Nothing
                            PictureBox5.Image = Nothing
                            PictureBox6.Image = Nothing
                        End If
                    End If
                End If
            End If
        End If
        PictureBox1.Refresh()

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Anteprima

        Dim x1 As Integer = TextBox3.Text
        Dim y1 As Integer = TextBox4.Text
        Dim boxSize3 = CInt(Windows.Forms.Cursor.Position.X) - CInt(x1)
        Dim boxSize4 = CInt(Windows.Forms.Cursor.Position.Y) - CInt(y1)

        If MouseButtons.HasFlag(MouseButtons.Left) Then
            Dim g1 = Graphics.FromHwnd(IntPtr.Zero)

            Dim screenWidth As Integer = Screen.GetWorkingArea(Me).Width
            Dim screenHeight As Integer = Screen.GetWorkingArea(Me).Height


            Dim brsh As Brush = New SolidBrush(Color.FromArgb(64, 255, 0, 0))  '128 trasparenza
            Dim pen2 As Pen = New Pen(brsh)
            'g1.DrawRectangle(pen2, x1, y1, boxSize3, boxSize4)
            g1.DrawRectangle(pen2, x1 - 1, y1 - 1, boxSize3 + 1, boxSize4 + 1)
            'g1.FillRectangle(brsh, x1, y1, boxSize3, boxSize4)
            Label8.Text = "Dimensions (px) : " & (CInt(Windows.Forms.Cursor.Position.X) - x1).ToString & " x " & (CInt(Windows.Forms.Cursor.Position.Y) - y1).ToString
            brsh.Dispose()
            g1.Dispose()
        End If

        If Not MouseButtons.HasFlag(MouseButtons.Left) Then
            boxSize3 = x1 + 2
            boxSize4 = y1 + 1
        End If
        'Fine Anteprima
    End Sub





    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Save colors to JASON
        SaveFileDialog1.Filter = "Files JASON|*.json"
        SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments
        SaveFileDialog1.FileName = "SpriteColors.json"
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, convertToJason(win32colorArray, JASONarray).ToString, False)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'To colors CSV
        SaveFileDialog1.Filter = "Files CSV|*.csv"
        SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments
        SaveFileDialog1.FileName = "SpriteColors.csv"
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, "Colors," & vbCrLf & CSVstring, False)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        'FlyBox
        If Flybox = 0 Then
            Flybox = 1
        ElseIf Flybox = 1 Then
            Flybox = 0
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'save1
        Dim latx As Integer = PictureBox4.Image.Width
        Dim laty As Integer = PictureBox4.Image.Height
        PictureBox8.Image = PictureBox4.Image
        Label7.Text = "Dimensions : " & latx.ToString & "x" & laty.ToString
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'save2
        Dim latx As Integer = PictureBox5.Image.Width
        Dim laty As Integer = PictureBox5.Image.Height
        PictureBox8.Image = PictureBox5.Image
        Label7.Text = "Dimensions : " & latx.ToString & "x" & laty.ToString
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'save3
        Dim latx As Integer = PictureBox6.Image.Width
        Dim laty As Integer = PictureBox6.Image.Height
        PictureBox8.Image = PictureBox6.Image
        Label7.Text = "Dimensions : " & latx.ToString & "x" & laty.ToString
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'save4
        Dim latx As Integer = PictureBox7.Image.Width
        Dim laty As Integer = PictureBox7.Image.Height
        PictureBox8.Image = PictureBox7.Image
        Label7.Text = "Dimensions : " & latx.ToString & "x" & laty.ToString
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Save Image Sprite
        SaveFileDialog1.Filter = "Files images PNG|*.png|Files images JPEG|*.jpg|Files images Bitmap|*.bmp|Files images GIF|*.gif|Files images TIFF|*.tif|Files images WMF|*.wmf"
        SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments
        SaveFileDialog1.FileName = "SpriteImage.png"
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Select Case SaveFileDialog1.FileName.Split(".").Last.ToLower
                Case "png"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Png)
                Case "jpg"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Jpeg)
                Case "bmp"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Bmp)
                Case "tif"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Tiff)
                Case "wmf"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Wmf)
                Case "gif"
                    PictureBox8.Image.Save(SaveFileDialog1.FileName, Imaging.ImageFormat.Gif)
            End Select
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'To Clipboard
        Clipboard.SetImage(PictureBox8.Image)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Delete Picturebox8 image
        PictureBox8.Image = Nothing
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Flip Image 90
        Dim Bitmap = New Bitmap(PictureBox8.Image)
        Bitmap.RotateFlip(RotateFlipType.Rotate90FlipXY)
        PictureBox8.Image = Bitmap
        PictureBox8.Refresh()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'To grayscale
        Dim grayscale As New Imaging.ColorMatrix(New Single()() _
        {
            New Single() {0.299, 0.299, 0.299, 0, 0},
            New Single() {0.587, 0.587, 0.587, 0, 0},
            New Single() {0.114, 0.114, 0.114, 0, 0},
            New Single() {0, 0, 0, 1, 0},
            New Single() {0, 0, 0, 0, 1}
        })

        Dim bmp As New Bitmap(PictureBox8.Image)
        Dim imgattr As New Imaging.ImageAttributes()
        imgattr.SetColorMatrix(grayscale)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.DrawImage(bmp, New Rectangle(0, 0, bmp.Width, bmp.Height),
                        0, 0, bmp.Width, bmp.Height,
                        GraphicsUnit.Pixel, imgattr)
        End Using
        PictureBox8.Image = bmp
        PictureBox8.Refresh()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'Invert Picturebox8
        PictureBox8.Image = InvertColorMatrix(PictureBox8.Image)
        PictureBox8.Refresh()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'To Library
        ToLibrary(PictureBox4.Image)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'To Library
        ToLibrary(PictureBox5.Image)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        'To Library
        ToLibrary(PictureBox6.Image)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        'To Library
        ToLibrary(PictureBox7.Image)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        'Save GIFanim
        'Save Image Sprite
        SaveFileDialog1.Filter = "Files images Animated GIF|*.gif"
        SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.MyDocuments
        SaveFileDialog1.FileName = "LibraryGIF.gif"
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then


            IntervalGIF = CInt(TextBox7.Text)
            Dim repeat As Integer = 0
            'AnimatedGIf

            Using gif = AnimatedGif.AnimatedGif.Create(SaveFileDialog1.FileName, IntervalGIF, repeat)
                For i = 0 To ImageList1.Images.Count - 1
                    Dim img = ImageList1.Images(i)

                    gif.AddFrame(img, delay:=-1, quality:=GifQuality.Default)
                Next
            End Using

            MsgBox("Animated GIF saved in " & SaveFileDialog1.FileName)

            PictureBox61.Image = Image.FromFile(SaveFileDialog1.FileName)
            PictureBox61.Refresh()
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        'To PNGs
        FolderBrowserDialog1.Description = "Select th folder to save all Library PNGs ..."
        FolderBrowserDialog1.SelectedPath = Environment.CurrentDirectory.ToString

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim currentdirectory As String = FolderBrowserDialog1.SelectedPath
            Dim pngdest As Bitmap = New Bitmap(ImageList1.ImageSize.Width, ImageList1.ImageSize.Height)
            Dim g As Graphics = Graphics.FromImage(pngdest)

            For i = 0 To ImageList1.Images.Count - 1
                g.DrawImage(ImageList1.Images(i), 0, 0)
                pngdest.Save(currentdirectory & "\PNG_" & i & ".png", Imaging.ImageFormat.Png)
            Next

            MsgBox("Saved in " & currentdirectory)
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        'To BMPs
        FolderBrowserDialog1.Description = "Select th folder to save all Library BMPs ..."
        FolderBrowserDialog1.SelectedPath = Environment.CurrentDirectory.ToString

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim currentdirectory As String = FolderBrowserDialog1.SelectedPath
            Dim pngdest As Bitmap = New Bitmap(ImageList1.ImageSize.Width, ImageList1.ImageSize.Height)
            Dim g As Graphics = Graphics.FromImage(pngdest)

            For i = 0 To ImageList1.Images.Count - 1
                g.DrawImage(ImageList1.Images(i), 0, 0)
                pngdest.Save(currentdirectory & "\PNG_" & i & ".bmp", Imaging.ImageFormat.Bmp)
            Next

            MsgBox("Saved in " & currentdirectory)
        End If
    End Sub












    'Delete Images ------------------
    Private Sub PictureBox4_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox4.DoubleClick
        'Delete img
        PictureBox4.Image = Nothing
        PictureBox4.Refresh()
    End Sub

    Private Sub PictureBox5_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox5.DoubleClick
        'Delete img
        PictureBox5.Image = Nothing
        PictureBox5.Refresh()
    End Sub

    Private Sub PictureBox6_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox6.DoubleClick
        'Delete img
        PictureBox6.Image = Nothing
        PictureBox6.Refresh()
    End Sub

    Private Sub PictureBox7_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox7.DoubleClick
        'Delete img
        PictureBox7.Image = Nothing
        PictureBox7.Refresh()
    End Sub

    Private Sub PictureBox8_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox8.DoubleClick
        'Delete img
        PictureBox8.Image = Nothing
        PictureBox8.Refresh()
    End Sub
    '--------------------------------------




















    Public Function ToLibrary(Pic As Image)

        Dim w As Integer = Pic.Width
        Dim h As Integer = Pic.Height

        Select Case LibraryPos
            Case 0
                PictureBox11.Image = Pic
            Case 1
                PictureBox12.Image = Pic
            Case 2
                PictureBox13.Image = Pic
            Case 3
                PictureBox14.Image = Pic
            Case 4
                PictureBox15.Image = Pic
            Case 5
                PictureBox16.Image = Pic
            Case 6
                PictureBox17.Image = Pic
            Case 7
                PictureBox18.Image = Pic
            Case 8
                PictureBox19.Image = Pic
            Case 9
                PictureBox20.Image = Pic
            Case 10
                PictureBox21.Image = Pic
            Case 11
                PictureBox22.Image = Pic
            Case 12
                PictureBox23.Image = Pic
            Case 13
                PictureBox24.Image = Pic
            Case 14
                PictureBox25.Image = Pic
            Case 15
                PictureBox26.Image = Pic
            Case 16
                PictureBox27.Image = Pic
            Case 17
                PictureBox28.Image = Pic
            Case 18
                PictureBox29.Image = Pic
            Case 19
                PictureBox30.Image = Pic
            Case 20
                PictureBox31.Image = Pic
            Case 21
                PictureBox32.Image = Pic
            Case 22
                PictureBox33.Image = Pic
            Case 23
                PictureBox34.Image = Pic
            Case 24
                PictureBox35.Image = Pic
            Case 25
                PictureBox36.Image = Pic
            Case 26
                PictureBox37.Image = Pic
            Case 27
                PictureBox38.Image = Pic
            Case 28
                PictureBox39.Image = Pic
            Case 29
                PictureBox40.Image = Pic
            Case 30
                PictureBox41.Image = Pic
            Case 31
                PictureBox42.Image = Pic
            Case 32
                PictureBox43.Image = Pic
            Case 33
                PictureBox44.Image = Pic
            Case 34
                PictureBox45.Image = Pic
            Case 35
                PictureBox46.Image = Pic
            Case 36
                PictureBox47.Image = Pic
            Case 37
                PictureBox48.Image = Pic
            Case 38
                PictureBox49.Image = Pic
            Case 39
                PictureBox50.Image = Pic
            Case 40
                PictureBox51.Image = Pic
            Case 41
                PictureBox52.Image = Pic
            Case 42
                PictureBox53.Image = Pic
            Case 43
                PictureBox54.Image = Pic
            Case 44
                PictureBox55.Image = Pic
            Case 45
                PictureBox56.Image = Pic
            Case 46
                PictureBox57.Image = Pic
            Case 47
                PictureBox58.Image = Pic
            Case 48
                PictureBox59.Image = Pic
            Case 49
                PictureBox60.Image = Pic
        End Select

        If LibraryPos = 0 Then
            ImageList1.ImageSize = New Size(Pic.Width, Pic.Height)
        Else
            If w > ImageList1.ImageSize.Width Then
                If h > ImageList1.ImageSize.Height Then
                    ImageList1.ImageSize = New Size(Pic.Width, Pic.Height)
                Else
                    ImageList1.ImageSize = New Size(Pic.Width, ImageList1.ImageSize.Height)
                End If
            Else
                If h > ImageList1.ImageSize.Height Then
                    ImageList1.ImageSize = New Size(ImageList1.ImageSize.Width, Pic.Height)
                Else
                    'null
                End If
            End If
        End If


        ImageList1.Images.Add(Pic)

        If LibraryPos >= 49 Then
            LibraryPos = 0
        Else
            LibraryPos = LibraryPos + 1
        End If

        'Image counter
        Label14.Text = LibraryPos.ToString & " / 50"

    End Function





    Private Sub PreviewCapture(ByVal sender As Object, ByVal e As EventArgs)
        Dim g = Graphics.FromHwnd(IntPtr.Zero)

        Dim screenWidth As Integer = Screen.GetWorkingArea(Me).Width
        Dim screenHeight As Integer = Screen.GetWorkingArea(Me).Height
        Dim boxSize = 100

        Dim brsh As Brush = New SolidBrush(Color.FromArgb(128, 255, 0, 0))

        g.FillRectangle(brsh, screenWidth - boxSize, screenHeight - boxSize, boxSize, boxSize)

        brsh.Dispose()
        g.Dispose()
    End Sub









    Public Function CreatePalette()

        ' Try creating a new image with a custom palette.
        Dim colors As New List(Of Color)
        Dim colorHTML As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml("#000000")

        ' Creates a new empty image with the pre-defined palette
        Dim g2 As System.Drawing.Graphics
        Dim img As New Bitmap(4000, 32, Drawing.Imaging.PixelFormat.Format24bppRgb)
        g2 = Graphics.FromImage(img)

        'Paint the canvas white
        g2.Clear(Color.White)
        'Set various modes to higher quality
        g2.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g2.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g2.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias

        'numero di immagini
        Dim numpbs As Integer = 0
        Dim posX As Integer = 0

        'bitmap immagine
        Dim b As Bitmap = New Bitmap(PictureBox1.Image)
        Dim stringcolors As String = ""

        '=======================================================
        '   COLORS USED
        '=======================================================
        Dim colorPaletteLength As Integer = b.Palette.Entries.Length
        Dim colorsArray(colorPaletteLength) As String

        If b.Palette IsNot Nothing And (colorPaletteLength > 0 And colorPaletteLength < 257) Then

            For i = 0 To colorPaletteLength - 1
                colorsArray(i) = ColorTranslator.ToHtml(b.Palette.Entries(i))
            Next

            Dim removeDuplicateColors As New List(Of String)()

            For Each item In colorsArray
                If removeDuplicateColors.Contains(item) = False Then
                    removeDuplicateColors.Add(item)
                End If
            Next

            Dim hexCount As Integer = 0

            For Each hexVal In removeDuplicateColors
                stringcolors &= "#" + hexVal.ToString & ","
                hexCount += 1
            Next

            For Each hexVal In removeDuplicateColors
                numpbs = numpbs + 1
                posX = numpbs * 32

                'Paint the rect
                Dim col As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml("#" & hexVal.ToString)
                Dim brush As New SolidBrush(col)
                Dim rect As New Rectangle(posX, 0, 4000, 32)
                g2.FillRectangle(brush, rect)
            Next



        End If

        PictureBox2.Image = img
        PictureBox2.Size = New Size(4000, 32)
        'TextBox2.Text &= "PALETTE:" & stringcolors & ","


        'Metodo 2
        '-------------------------------------------

        ' Creates a new empty image with the pre-defined palette
        Dim g3 As System.Drawing.Graphics
        Dim img3 As New Bitmap(4000, 32, Drawing.Imaging.PixelFormat.Format24bppRgb)
        g3 = Graphics.FromImage(img3)   'img

        Dim curPixColor As Color
        numpbs = 0
        posX = 0
        Dim LineY As Integer = 0
        Dim HTMLstring As String = ""
        Dim win32color As String = ""

        For x = 0 To b.Width - 1
            For y = 0 To b.Height - 1
                curPixColor = b.GetPixel(x, y)

                If x = 0 And y = 0 Then
                    colors.Add(curPixColor)
                    numpbs = numpbs + 1
                    posX = numpbs * 32
                    Dim brush3 As New SolidBrush(curPixColor)
                    Dim rect As New Rectangle(LineY, 0, LineY, 32)
                    g3.FillRectangle(brush3, rect)
                    LineY = LineY + 32
                    HTMLstring = HTMLstring & ColorTranslator.ToHtml(curPixColor).ToString & ","
                    win32color = win32color & ColorTranslator.ToWin32(curPixColor).ToString & ","
                Else
                    If colors.Contains(curPixColor) Then
                        'null
                    Else
                        colors.Add(curPixColor)
                        numpbs = numpbs + 1
                        posX = numpbs * 32
                        Dim brush3 As New SolidBrush(curPixColor)
                        Dim rect As New Rectangle(LineY, 0, LineY, 32)
                        g3.FillRectangle(brush3, rect)
                        LineY = LineY + 32
                        HTMLstring = HTMLstring & ColorTranslator.ToHtml(curPixColor).ToString & ","
                        win32color = win32color & ColorTranslator.ToWin32(curPixColor).ToString & ","
                    End If
                End If
            Next
        Next
        PictureBox3.Image = img3
        TextBox2.Text &= "NUM_COLORS:" & numpbs.ToString & ","
        TextBox2.Text &= "COLORS:" & HTMLstring
        PictureBox3.Size = New Size(numpbs * 32, 32)
        CSVstring = HTMLstring

        JASONarray = Split(HTMLstring, ",")
        win32colorArray = Split(win32color, ",")

    End Function

    Public Function ConvertToRbg(ByVal HexColor As String) As Color
        Dim Red As String
        Dim Green As String
        Dim Blue As String
        HexColor = Replace(HexColor, "#", "")
        Red = Val("&H" & Mid(HexColor, 1, 2))
        Green = Val("&H" & Mid(HexColor, 3, 2))
        Blue = Val("&H" & Mid(HexColor, 5, 2))
        Return Color.FromArgb(Red, Green, Blue)
    End Function




    Private Function convertToJason(ByVal firstLine As Array, ByVal values As Array)
        ' vbCrLf & vbNewLine
        Dim localString As String = ""
        Dim localStringBuilder As String

        localStringBuilder = ("{" & vbCrLf)
        localStringBuilder = localStringBuilder & "ColorsWin32HTML:" & firstLine(0) & ";" & values(0) & "," & vbCrLf
        For index As Integer = 0 To values.Length - 2
            localStringBuilder = localStringBuilder & firstLine(index) & ";" & values(index) & "," & vbCrLf
        Next ' literal.Substring(0, 3) 
        localStringBuilder = localStringBuilder.Substring(0, localStringBuilder.Length - 3)
        localStringBuilder = localStringBuilder & vbCrLf & "},"

        Debug.Print(localStringBuilder)


        Return localStringBuilder

    End Function

    Private Shared Function InvertColorMatrix(ByVal imgSource As Image) As Image
        Dim bmpDest As Bitmap = New Bitmap(imgSource.Width, imgSource.Height)
        Dim clrMatrix As ColorMatrix = New ColorMatrix(New Single() _
         () {New Single() {-1, 0, 0, 0, 0}, New Single() {0, -1,
         0, 0, 0}, New Single() {0, 0, -1, 0, 0}, New Single() _
         {0, 0, 0, 1, 0}, New Single() {1, 1, 1, 0, 1}})

        Using attrImage As ImageAttributes = New ImageAttributes()
            attrImage.SetColorMatrix(clrMatrix)

            Using g As Graphics = Graphics.FromImage(bmpDest)
                g.DrawImage(imgSource, New Rectangle(0, 0,
               imgSource.Width, imgSource.Height), 0, 0,
               imgSource.Width, imgSource.Height,
               GraphicsUnit.Pixel, attrImage)
            End Using
        End Using

        Return bmpDest
    End Function


End Class
