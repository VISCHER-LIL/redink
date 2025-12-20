' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public Class SplashScreenWorks
        Inherits System.Windows.Forms.Form

        ' ─── Controls & state ────────────────────────────────────────
        Private lblMessage As System.Windows.Forms.Label
        Private picLogo As System.Windows.Forms.PictureBox
        Private remainingSeconds As Integer
        Private baseText As String
        Private countdownCts As System.Threading.CancellationTokenSource

        ''' <summary>
        ''' Fires when the user presses Esc.
        ''' </summary>
        Public Event CancelRequested As System.EventHandler

        ' ─── WinAPI for borderless dragging ───────────────────────────
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

            ' ─── Form setup ───────────────────────────────────────────
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.BackColor = System.Drawing.ColorTranslator.FromWin32(&H8000000F)
            Me.KeyPreview = True

            ' ─── Logo ─────────────────────────────────────────────────
            picLogo = New System.Windows.Forms.PictureBox() With {
        .Image = New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo),
        .SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
    }
            Me.Controls.Add(picLogo)

            ' ─── Label ───────────────────────────────────────────────
            Dim stdFont As System.Drawing.Font =
        New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            lblMessage = New System.Windows.Forms.Label() With {
        .Font = stdFont,
        .AutoSize = True,
        .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    }
            Me.Controls.Add(lblMessage)

            ' ─── Layout & initial text ───────────────────────────────
            baseText = customText
            remainingSeconds = countdownSeconds
            Dim initialText As String = If(countdownSeconds > 0,
                                   $"{customText} {countdownSeconds}s",
                                   customText)
            lblMessage.Text = initialText

            Dim padding As Integer = 10
            Dim textSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(initialText, stdFont)
            lblMessage.Size = textSize

            ' Logo height matches text height for equal top/bottom padding
            Dim logoSize As Integer = textSize.Height
            picLogo.SetBounds(padding, padding, logoSize, logoSize)

            ' Center label vertically next to logo
            Dim labelX As Integer = picLogo.Right + padding
            Dim labelY As Integer = padding + (logoSize - textSize.Height) \ 2
            lblMessage.SetBounds(labelX, labelY, textSize.Width, textSize.Height)

            ' Auto‐size form to content (unless overridden)
            Dim clientW As Integer = lblMessage.Right + padding
            Dim clientH As Integer = logoSize + padding * 2
            If formWidth > 0 Then clientW = formWidth
            If formHeight > 0 Then clientH = formHeight
            Me.ClientSize = New System.Drawing.Size(clientW, clientH)

            ' ─── ESC cancels ──────────────────────────────────────────
            AddHandler Me.KeyDown, AddressOf OnKeyDown

            ' ─── Start countdown if needed ───────────────────────────
            If countdownSeconds > 0 Then
                StartCountdown()
            End If
        End Sub

        ''' <summary>
        ''' Updates the label instantly without affecting the countdown.
        ''' </summary>
        Public Sub UpdateMessage(ByVal newMessage As String)
            lblMessage.Text = newMessage
            Dim newSize As System.Drawing.Size = System.Windows.Forms.TextRenderer.MeasureText(newMessage, lblMessage.Font)
            lblMessage.Size = newSize
            lblMessage.Refresh()
        End Sub

        ''' <summary>
        ''' Stops any running countdown and starts a new one.
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
        ''' Fires every second on a background Task and updates the UI via Invoke.
        ''' </summary>
        Private Sub StartCountdown()
            ' Cancel previous if running
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

                                                    ' Update on UI thread
                                                    If Not Me.IsDisposed Then
                                                        If Me.InvokeRequired Then
                                                            Me.Invoke(Sub() UpdateMessage($"{baseText} {remainingSeconds}s"))
                                                        Else
                                                            UpdateMessage($"{baseText} {remainingSeconds}s")
                                                        End If
                                                    End If
                                                End While
                                            End Function)
        End Sub

        ''' <summary>
        ''' ESC closes + raises CancelRequested.
        ''' </summary>
        Private Sub OnKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                countdownCts?.Cancel()
                RaiseEvent CancelRequested(Me, System.EventArgs.Empty)
                Me.Close()
            End If
        End Sub

        ''' <summary>
        ''' Allow dragging borderless form.
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