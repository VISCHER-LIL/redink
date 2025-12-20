' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class SplashScreenCountDown
        Inherits System.Windows.Forms.Form

        ' ─── Controls & state ───────────────────────────────────────
        Private lblMessage As System.Windows.Forms.Label
        Private picLogo As System.Windows.Forms.PictureBox
        Private remainingSeconds As Integer
        Private baseText As String
        Private countdownCts As System.Threading.CancellationTokenSource

        ' Used to wait until the form is loaded before returning from Show()
        Private loadedEvent As System.Threading.ManualResetEventSlim
        Private splashThread As System.Threading.Thread

        ''' <summary>
        ''' Fires when the user presses Esc.
        ''' </summary>
        Public Event CancelRequested As System.EventHandler

        ' ─── WinAPI for dragging ─────────────────────────────────────
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function ReleaseCapture() As Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SendMessage(
    ByVal hWnd As IntPtr,
    ByVal wMsg As Integer,
    ByVal wParam As IntPtr,
    ByVal lParam As IntPtr
) As IntPtr
        End Function

        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const HTCAPTION As Integer = 2

        ''' <summary>
        ''' customText: text prefix  
        ''' formWidth/Height: if >0, override autosize  
        ''' countdownSeconds: initial countdown length (0=no countdown)  
        ''' </summary>
        Public Sub New(
    Optional ByVal customText As String = "Please wait …",
    Optional ByVal formWidth As Integer = 0,
    Optional ByVal formHeight As Integer = 0,
    Optional ByVal countdownSeconds As Integer = 0)

            MyBase.New()

            ' ─── Form basics ──────────────────────────────────────────
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.BackColor = System.Drawing.ColorTranslator.FromWin32(&H8000000F)
            Me.KeyPreview = True
            Me.TopMost = True

            ' ─── Logo ──────────────────────────────────────────────────
            picLogo = New System.Windows.Forms.PictureBox() With {
        .Image = New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo),
        .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
    }
            Me.Controls.Add(picLogo)

            ' ─── Label ────────────────────────────────────────────────
            Dim stdFont As System.Drawing.Font =
        New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            lblMessage = New System.Windows.Forms.Label() With {
        .Font = stdFont,
        .AutoSize = True,
        .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    }
            Me.Controls.Add(lblMessage)

            ' ─── Layout & initial text ────────────────────────────────
            baseText = customText
            remainingSeconds = countdownSeconds
            Dim initialText As String = If(countdownSeconds > 0,
                                   $"{customText} {countdownSeconds}s",
                                   customText)
            lblMessage.Text = initialText

            Dim padding As Integer = 10
            Dim textSize As System.Drawing.Size =
        System.Windows.Forms.TextRenderer.MeasureText(initialText, stdFont)
            lblMessage.Size = textSize

            ' logo height == text height (equal vertical padding)
            Dim logoSize As Integer = textSize.Height
            picLogo.SetBounds(padding, padding, logoSize, logoSize)

            ' center label vertically next to logo
            Dim labelX As Integer = picLogo.Right + padding
            Dim labelY As Integer = padding + (logoSize - textSize.Height) \ 2
            lblMessage.SetBounds(labelX, labelY, textSize.Width, textSize.Height)

            ' auto-size form (unless overridden)
            Dim clientW As Integer = lblMessage.Right + padding
            Dim clientH As Integer = logoSize + padding * 2
            If formWidth > 0 Then clientW = formWidth
            If formHeight > 0 Then clientH = formHeight
            Me.ClientSize = New System.Drawing.Size(clientW, clientH)

            ' ESC cancels
            AddHandler Me.KeyDown, AddressOf OnKeyDown

            ' kick off countdown if requested
            If countdownSeconds > 0 Then
                StartCountdown()
            End If
        End Sub

        ''' <summary>
        ''' Instance-based Show: spins up its own STA thread & message loop.
        ''' </summary>
        Public Shadows Sub Show()
            ' prevent multiple shows
            If splashThread IsNot Nothing Then Return

            loadedEvent = New System.Threading.ManualResetEventSlim(False)

            ' start a new STA thread for this form
            splashThread = New System.Threading.Thread(Sub()
                                                           ' signal when the form is loaded
                                                           AddHandler Me.Load, Sub(s, e) loadedEvent.Set()
                                                           System.Windows.Forms.Application.Run(Me)
                                                       End Sub)

            splashThread.SetApartmentState(System.Threading.ApartmentState.STA)
            splashThread.IsBackground = True
            splashThread.Start()

            ' wait until the Load event has fired
            loadedEvent.Wait()
        End Sub

        ''' <summary>
        ''' Instance-based Close: marshals back to the form's thread.
        ''' </summary>
        Public Shadows Sub Close()
            If Me.InvokeRequired Then
                Me.Invoke(New System.Action(Sub() MyBase.Close()))
            Else
                MyBase.Close()
            End If
        End Sub

        ''' <summary>
        ''' Update the label text without affecting the countdown.
        ''' </summary>
        Public Sub UpdateMessage(ByVal newMessage As String)
            If Me.InvokeRequired Then
                Me.Invoke(New System.Action(Sub() UpdateMessage(newMessage)))
            Else
                lblMessage.Text = newMessage
                Dim newSize As System.Drawing.Size =
            System.Windows.Forms.TextRenderer.MeasureText(newMessage, lblMessage.Font)
                lblMessage.Size = newSize
                lblMessage.Refresh()
            End If
        End Sub

        ''' <summary>
        ''' Stop any running countdown and start a fresh one.
        ''' </summary>
        Public Sub RestartCountdown(
    ByVal seconds As Integer,
    Optional ByVal newBaseText As String = Nothing)

            If newBaseText IsNot Nothing Then
                baseText = newBaseText
            End If

            remainingSeconds = seconds
            UpdateMessage($"{baseText} {remainingSeconds}s")
            StartCountdown()
        End Sub

        ''' <summary>
        ''' Runs on a background Task, delays 1s between ticks, marshals updates via Invoke.
        ''' </summary>
        Private Sub StartCountdown()
            ' cancel prior if any
            countdownCts?.Cancel()
            countdownCts = New System.Threading.CancellationTokenSource()
            Dim ct = countdownCts.Token

            System.Threading.Tasks.Task.Run(Async Function()
                                                While remainingSeconds > 0 AndAlso Not ct.IsCancellationRequested
                                                    Try
                                                        Await System.Threading.Tasks.Task.Delay(1000, ct)
                                                    Catch ex As System.Threading.Tasks.TaskCanceledException
                                                        Exit While
                                                    End Try

                                                    remainingSeconds -= 1
                                                    If remainingSeconds < 0 Then remainingSeconds = 0

                                                    ' marshal update to UI thread
                                                    If Not Me.IsDisposed Then
                                                        If Me.InvokeRequired Then
                                                            Me.Invoke(New System.Action(Sub()
                                                                                            lblMessage.Text = $"{baseText} {remainingSeconds}s"
                                                                                            lblMessage.Size = System.Windows.Forms.TextRenderer.MeasureText(lblMessage.Text, lblMessage.Font)
                                                                                        End Sub))
                                                        Else
                                                            lblMessage.Text = $"{baseText} {remainingSeconds}s"
                                                            lblMessage.Size = System.Windows.Forms.TextRenderer.MeasureText(lblMessage.Text, lblMessage.Font)
                                                        End If
                                                    End If
                                                End While
                                            End Function)
        End Sub

        ''' <summary>
        ''' Closes + raises CancelRequested when Esc is pressed.
        ''' </summary>
        Private Sub OnKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                countdownCts?.Cancel()
                RaiseEvent CancelRequested(Me, System.EventArgs.Empty)
                Close()
            End If
        End Sub

        ''' <summary>
        ''' Allow dragging the borderless form.
        ''' </summary>
        Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseDown(e)
            If e.Button = System.Windows.Forms.MouseButtons.Left Then
                ReleaseCapture()
                SendMessage(Me.Handle, WM_NCLBUTTONDOWN, CType(HTCAPTION, IntPtr), IntPtr.Zero)
            End If
        End Sub

    End Class
End Namespace