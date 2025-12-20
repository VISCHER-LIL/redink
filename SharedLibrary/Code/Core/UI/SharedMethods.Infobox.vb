' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods


        Public Class InfoBox

            Inherits Form

            Private Shared InfoBox As InfoBox
            Private timer As System.Windows.Forms.Timer
            Private label As System.Windows.Forms.Label

            Private Sub New(ByVal text As String, ByVal duration As Integer)
                ' Set form properties
                Me.Text = ""
                Me.FormBorderStyle = FormBorderStyle.None
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.BackColor = ColorTranslator.FromWin32(&H8000000F)
                Me.TopMost = True

                ' Create and add the App logo PictureBox
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                Dim iconPictureBox As New PictureBox()
                iconPictureBox.Image = bmp
                iconPictureBox.SizeMode = PictureBoxSizeMode.Zoom
                iconPictureBox.Size = New Size(32, 32) ' Icon size
                iconPictureBox.Location = New System.Drawing.Point(10, 10) ' Top-left corner
                Me.Controls.Add(iconPictureBox)

                ' Initialize label
                label = New System.Windows.Forms.Label()
                label.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
                label.TextAlign = ContentAlignment.MiddleLeft
                label.MaximumSize = New Size(450, 240)
                label.Width = 450
                label.Height = 240
                label.Text = text
                label.AutoSize = True
                label.AutoEllipsis = True
                'SetWrappedText(label, text)  ' not necessary, if autoellipsis is set

                ' Adjust form size dynamically to accommodate PictureBox and label
                Dim contentRight As Integer = iconPictureBox.Right + 10
                Me.ClientSize = New Size(Math.Max(contentRight + label.Width + 10, iconPictureBox.Width + 20), Math.Max(label.Height + 20, iconPictureBox.Height + 20))

                ' Position label below the icon
                label.Location = New System.Drawing.Point(contentRight, 10)
                Me.Controls.Add(label)


                ' Initialize and start timer if duration > 0
                If duration > 0 Then
                    timer = New System.Windows.Forms.Timer()
                    timer.Interval = duration * 1000
                    AddHandler timer.Tick, AddressOf Timer_Tick
                    timer.Start()
                End If
            End Sub

            Private Sub SetWrappedText(lbl As System.Windows.Forms.Label, text As String)
                ' Set the wrapped text in the label
                lbl.Text = text

                Using g As Graphics = lbl.CreateGraphics()
                    ' Measure the size of the text
                    Dim size As SizeF = g.MeasureString(text, lbl.Font, lbl.Width)

                    ' Check if the text exceeds the maximum label height
                    Dim lineHeight As Single = lbl.Font.GetHeight(g)
                    Dim maxLines As Integer = CInt(System.Math.Floor(lbl.MaximumSize.Height / lineHeight))
                    Dim textLines As Integer = CInt(System.Math.Ceiling(size.Height / lineHeight))

                    If textLines > maxLines Then
                        ' Truncate and add ellipsis if exceeding the maximum visible lines
                        Dim visibleText As String = text.Substring(0, CInt(System.Math.Min(text.Length, lbl.Width * maxLines \ CLng(lbl.Font.Size)))) & " (...)"
                        lbl.Text = visibleText
                    End If
                End Using
            End Sub


            Private Sub Timer_Tick(ByVal sender As Object, ByVal e As EventArgs)
                Me.Close()
            End Sub

            Public Shared Sub ShowInfoBox(ByVal text As String, Optional ByVal duration As Integer = 0)
                ' Close current InfoBox if open
                If InfoBox IsNot Nothing Then
                    InfoBox.Close()
                End If

                ' If text is empty, return without creating a new form
                If String.IsNullOrEmpty(text) Then
                    Return
                End If

                ' Create a new InfoBox instance and display it
                InfoBox = New InfoBox(text, duration)
                InfoBox.Show()
                InfoBox.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End Sub

        End Class

    End Class
End Namespace