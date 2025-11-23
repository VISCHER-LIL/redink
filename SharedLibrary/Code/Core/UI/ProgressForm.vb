' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public Class ProgressForm
        Inherits System.Windows.Forms.Form

        Private WithEvents progressBar As System.Windows.Forms.ProgressBar
        Private WithEvents lblHeader As System.Windows.Forms.Label
        Private WithEvents lblStatus As System.Windows.Forms.Label
        Private WithEvents btnCancel As System.Windows.Forms.Button
        Private WithEvents uiTimer As System.Windows.Forms.Timer

        ' Constructor: receives the header text and the initial status text.
        Public Sub New(headerText As String, initialLabel As String)
            ' --- Use Font scaling ---
            Dim standardFont As New System.Drawing.Font(
    "Segoe UI",
    9.0F,
    System.Drawing.FontStyle.Regular,
    System.Drawing.GraphicsUnit.Point)

            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Font = standardFont
            Me.AutoSize = True
            Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.Text = SharedMethods.AN ' headerText

            ' --- Icon setzen ---
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' --- Header Label ---
            lblHeader = New System.Windows.Forms.Label() With {
    .Text = headerText,
    .AutoSize = True
}

            ' --- ProgressBar ---
            progressBar = New System.Windows.Forms.ProgressBar() With {
    .Minimum = 0,
    .Maximum = ProgressBarModule.GlobalProgressMax,
    .Dock = System.Windows.Forms.DockStyle.Fill
}

            ' --- Status Label ---
            lblStatus = New System.Windows.Forms.Label() With {
    .Text = initialLabel,
    .AutoSize = True,
    .Dock = System.Windows.Forms.DockStyle.Fill
}

            ' --- Cancel Button ---
            btnCancel = New System.Windows.Forms.Button() With {
    .Text = "Cancel",
    .AutoSize = True
}
            AddHandler btnCancel.Click, AddressOf btnCancel_Click

            ' --- Layout in TableLayoutPanel ---
            Dim layout As New System.Windows.Forms.TableLayoutPanel() With {
    .AutoSize = True,
    .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
    .Dock = System.Windows.Forms.DockStyle.Fill,
    .Padding = New System.Windows.Forms.Padding(10),
    .ColumnCount = 1,
    .RowCount = 4
}
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

            layout.Controls.Add(lblHeader, 0, 0)
            layout.Controls.Add(progressBar, 0, 1)
            layout.Controls.Add(lblStatus, 0, 2)
            layout.Controls.Add(btnCancel, 0, 3)

            Me.Controls.Add(layout)

            ' --- UI-Timer für periodische Updates ---
            uiTimer = New System.Windows.Forms.Timer() With {
    .Interval = 250 ' Update every 250 ms
}
            AddHandler uiTimer.Tick, AddressOf Timer_Tick
            uiTimer.Start()
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
                ' Possible exception if the form is closing.
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