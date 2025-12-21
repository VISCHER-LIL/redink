' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class DPIProgressForm
        Inherits System.Windows.Forms.Form

        Private WithEvents progressBar As System.Windows.Forms.ProgressBar
        Private WithEvents lblHeader As System.Windows.Forms.Label
        Private WithEvents lblStatus As System.Windows.Forms.Label
        Private WithEvents btnCancel As System.Windows.Forms.Button
        Private WithEvents uiTimer As System.Windows.Forms.Timer

        ' Constructor receives the header text and the initial status text.
        Public Sub New(headerText As String, initialLabel As String)
            ' --- Auto-Scale für DPI und Font ---
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

            ' --- Form-Eigenschaften ---
            Me.ClientSize = New System.Drawing.Size(400, 220)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Text = headerText

            ' Icon setzen
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Standard-Font
            Dim standardFont As New System.Drawing.Font(
    "Segoe UI",
    9.0F,
    System.Drawing.FontStyle.Regular,
    System.Drawing.GraphicsUnit.Point)

            ' --- Header Label ---
            lblHeader = New System.Windows.Forms.Label()
            lblHeader.Text = "Progress ..."
            lblHeader.AutoSize = True
            lblHeader.Font = standardFont
            lblHeader.Location = New System.Drawing.Point(10, 10)
            lblHeader.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left
            Me.Controls.Add(lblHeader)

            ' --- ProgressBar ---
            progressBar = New System.Windows.Forms.ProgressBar()
            progressBar.Minimum = 0
            progressBar.Maximum = ProgressBarModule.GlobalProgressMax
            progressBar.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, 25)
            progressBar.Location = New System.Drawing.Point(10, 40)
            progressBar.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                     System.Windows.Forms.AnchorStyles.Left Or
                     System.Windows.Forms.AnchorStyles.Right
            Me.Controls.Add(progressBar)

            ' --- Status Label ---
            lblStatus = New System.Windows.Forms.Label()
            lblStatus.Text = initialLabel
            lblStatus.AutoSize = False
            lblStatus.Font = standardFont
            lblStatus.Location = New System.Drawing.Point(10, 75)
            lblStatus.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, 20)
            lblStatus.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                 System.Windows.Forms.AnchorStyles.Left Or
                 System.Windows.Forms.AnchorStyles.Right
            Me.Controls.Add(lblStatus)

            ' --- Cancel Button ---
            btnCancel = New System.Windows.Forms.Button()
            btnCancel.Text = "Cancel"
            btnCancel.Font = standardFont
            btnCancel.AutoSize = True
            btnCancel.Location = New System.Drawing.Point(10, 120)
            btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            ' --- Resize-Event für dynamische Anpassung ---
            AddHandler Me.ClientSizeChanged, AddressOf Form_Resize

            ' --- UI-Timer für periodische Updates ---
            uiTimer = New System.Windows.Forms.Timer()
            uiTimer.Interval = 250 ' Update every 250 ms
            AddHandler uiTimer.Tick, AddressOf Timer_Tick
            uiTimer.Start()
        End Sub

        ' Dynamisches Anpassen der Steuerelemente bei Größenänderung
        Private Sub Form_Resize(sender As Object, e As EventArgs)
            progressBar.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, progressBar.Height)
            lblStatus.Size = New System.Drawing.Size(Me.ClientSize.Width - 20, lblStatus.Height)
        End Sub

        ' Timer tick event updates the progress bar and status label.
        Private Sub Timer_Tick(sender As Object, e As EventArgs)
            Try
                ' Update the progress bar maximum and value.
                progressBar.Maximum = ProgressBarModule.GlobalProgressMax
                progressBar.Value = Math.Min(ProgressBarModule.GlobalProgressValue, progressBar.Maximum)

                ' Update the status text.
                lblStatus.Text = ProgressBarModule.GlobalProgressLabel

                ' If the cancel flag is set, close the form with a Cancel result.
                If ProgressBarModule.CancelOperation Then
                    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Me.Close()
                End If
            Catch ex As System.Exception
                ' It is possible to get an exception if the form is closing.
                System.Diagnostics.Debug.WriteLine("Timer error: " & ex.Message)
            End Try
        End Sub

        ' When the Cancel button is clicked, set the global cancel flag.
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ProgressBarModule.CancelOperation = True
        End Sub

        ' Stop the timer when the form is closed.
        Protected Overrides Sub OnFormClosed(e As System.Windows.Forms.FormClosedEventArgs)
            uiTimer.Stop()
            ProgressBarModule.CancelOperation = True
            MyBase.OnFormClosed(e)
        End Sub
    End Class

End Namespace