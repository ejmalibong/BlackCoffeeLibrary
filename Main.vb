Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Windows
Imports System.Windows.Forms
Imports System.Drawing

Namespace BlackCoffee

    ''' <summary>
    ''' Compiled and extended by: Efraim Joshua B. Malibong
    ''' Email address: ej.malibong15@gmail.com
    ''' Date: March 14, 2021
    ''' </summary>
    Public Class Main

#Region "DataGridView Optimization"
        ''' <summary>
        ''' Enables DataGridView double buffer to optimize loading of large data.
        ''' </summary>
        ''' <param name="dgv">Bound or unbound DataGridView control</param>
        ''' <remarks></remarks>
        Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)
            Dim dgvType As Type = dgv.[GetType]()
            Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            pi.SetValue(dgv, True, Nothing)
        End Sub
#End Region

#Region "Image Manipulation"
        ''' <summary>
        ''' Gets the encoder used for the image. 
        ''' </summary>
        ''' <param name="format"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo
            Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
            Dim codec As ImageCodecInfo
            For Each codec In codecs
                If codec.FormatID = format.Guid Then
                    Return codec
                End If
            Next codec

            Return Nothing
        End Function

        ''' <summary>
        ''' Resizes the image to reduce file size.
        ''' </summary>
        ''' <param name="image">Image to be resize.</param>
        ''' <param name="size">Proposed size for the image.</param>
        ''' <param name="preserveAspectRatio">Boolean if image resizing preserves aspect ratio.</param>
        ''' <remarks> from http://mrbigglesworth79.blogspot.com/2011/04/resizing-image-on-fly-using-net.html </remarks>
        Public Function ResizeImage(ByVal image As Image, ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image
            Dim newWidth As Integer
            Dim newHeight As Integer

            If preserveAspectRatio Then
                Dim originalWidth As Integer = image.Width
                Dim originalHeight As Integer = image.Height
                Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
                Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
                Dim percent As Single = If(percentHeight < percentWidth, percentHeight, percentWidth)
                newWidth = CInt(originalWidth * percent)
                newHeight = CInt(originalHeight * percent)
            Else
                newWidth = size.Width
                newHeight = size.Height
            End If

            Dim newImage As Image = New Bitmap(newWidth, newHeight)

            Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
                graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
            End Using

            Return newImage
        End Function
#End Region

#Region "Conversion & Formatting"
        ''' <summary>
        ''' Converts bytes to a more readable format with round up.
        ''' </summary>
        ''' <param name="fileLength">Bytes to be converted.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReadableFileSize(ByVal fileLength As Long) As String
            Dim suffix As String = String.Empty

            If fileLength < 1000 Then
                suffix = " Bytes"
                GoTo showIt
            End If
            If fileLength > 1000000 Then
                fileLength = Int(fileLength / 1000000)
                suffix = " Mb"
                GoTo showIt
            Else
                fileLength = Int(fileLength / 1000)
                suffix = " Kb"
            End If

showIt:     Return fileLength & suffix
        End Function

        'format date to capture am to pm

        ''' <summary>
        ''' Formats date/datetime to capture whole shifting hours in case of day shift or night shift variables.
        ''' </summary>
        ''' <param name="dateParam">Date to be formatted.</param>
        ''' <param name="isStartDate">Boolean if dateParam capture 12 am to 12 pm.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FormatDate(ByVal dateParam As DateTime, ByVal isStartDate As Boolean) As String
            Dim year As String = "" & dateParam.Year
            Dim month As String = If((dateParam.Month < 10), "0" & dateParam.Month, "" & dateParam.Month)
            Dim day As String = If((dateParam.Day < 10), "0" & dateParam.Day, "" & dateParam.Day)

            If isStartDate = True Then
                Return year & "-" & month & "-" & day & " 00:00:00"
            Else
                Return year & "-" & month & "-" & day & " 23:59:59"
            End If
        End Function
#End Region

#Region "MDI Support"
        ''' <summary>
        ''' Allows single instance only of a child form.
        ''' </summary>
        ''' <param name="frmParent">Parent form. Set IsMdiContainer to True.</param>
        ''' <param name="frmChild">Child form to be call.</param>
        ''' <remarks>If IsFill is set to True, change the child form's FormBorderStyle to None.</remarks>
        Public Sub FormLoader(ByVal frmParent As Form, ByVal frmChild As Form, Optional ByVal isFill As Boolean = False)
            Try
                For Each mdiChild As Form In frmParent.MdiChildren
                    If mdiChild.Name = frmChild.Name Then
                        mdiChild.Activate()
                        mdiChild.KeyPreview = True
                        Exit Sub
                    End If
                Next

                frmChild.MdiParent = frmParent
                frmChild.Show()

                If isFill = True Then
                    frmChild.FormBorderStyle = FormBorderStyle.None
                    frmChild.Dock = DockStyle.Fill
                End If

                frmChild.StartPosition = FormStartPosition.CenterParent
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.SetExcpTitle(ex), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Traps child form inside the parent form or mdi container.
        ''' </summary>
        ''' <param name="form">Child form.</param>
        ''' <remarks></remarks>
        Public Sub FormTrap(ByVal form As Form)
            If form.Left <= 0 Then
                form.Location = New Point(0, form.Location.Y)
                If form.Top < 0 Then
                    form.Top = 0
                ElseIf form.Top > ((form.MdiParent.ClientRectangle.Height - 50) - form.Height) Then
                    form.Top = (form.MdiParent.ClientRectangle.Height - 30) - form.Height
                End If
            ElseIf form.Right >= form.MdiParent.ClientRectangle.Width Then
                form.Left = (form.MdiParent.Width - 20) - form.Width
            ElseIf form.Top < 0 Then
                form.Top = 0
            ElseIf form.Top >= ((form.MdiParent.ClientRectangle.Height - 50) - form.Height) Then
                form.Top = (form.MdiParent.ClientRectangle.Height - 50) - form.Height
            End If
        End Sub

        ''' <summary>
        ''' Reset all controls of a form usually the date entry form.
        ''' </summary>
        ''' <param name="form">Data entry form.</param>
        ''' <remarks></remarks>
        Public Sub FormReset(ByVal form As Control)
            Try
                For Each ctrl As Control In form.Controls
                    FormReset(ctrl)

                    If TypeOf ctrl Is TextBox Then
                        CType(ctrl, TextBox).Text = String.Empty
                    End If

                    If TypeOf ctrl Is ComboBox Then
                        CType(ctrl, ComboBox).SelectedValue = 0
                    End If

                    If TypeOf ctrl Is DateTimePicker Then
                        CType(ctrl, DateTimePicker).Value = Date.Now
                    End If
                Next ctrl
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.SetExcpTitle(ex), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
#End Region

#Region "Exception Handling"
        ''' <summary>
        ''' Set exception message with line number for easy debugging.
        ''' </summary>
        ''' <param name="ex">Exception</param>
        ''' <param name="format">Boolean if class involve show or only the line number.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetExcpTitle(ByVal ex As Exception, Optional ByVal format As Integer = 1) As String
            Dim sTrace = New System.Diagnostics.StackTrace(ex, True)
            Dim frames() As StackFrame = sTrace.GetFrames()
            Dim title As String = String.Empty

            For Each frame As StackFrame In frames
                If format = 1 Then
                    title = frame.GetFileLineNumber().ToString
                Else
                    title = frame.GetFileName & " " & frame.GetFileLineNumber().ToString
                    title += frame.GetFileLineNumber().ToString
                End If
            Next

            Return title
        End Function
#End Region

#Region "Clear Data Entry Form"
        ''' <summary>
        ''' Iterates to all controls of a form and set each to its default value.
        ''' </summary>
        ''' <param name="form">Owner form.</param>
        ''' <remarks></remarks>
        Public Sub ClearForm(ByVal form As Control)
            Try
                For Each ctrl As Control In form.Controls
                    ClearForm(ctrl)

                    If TypeOf ctrl Is TextBox Then
                        CType(ctrl, TextBox).Text = String.Empty
                    End If

                    If TypeOf ctrl Is ComboBox Then
                        CType(ctrl, ComboBox).SelectedValue = 0
                    End If

                    If TypeOf ctrl Is DateTimePicker Then
                        CType(ctrl, DateTimePicker).Value = Date.Now
                    End If
                Next ctrl
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.SetExcpTitle(ex), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
#End Region

    End Class

End Namespace