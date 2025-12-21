' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Threading

Namespace SharedLibrary
    Public Module ProgressBarModule
        ' Global variables to control the progress form.
        Public GlobalProgressValue As Integer = 0
        Public GlobalProgressMax As Integer = 100
        Public GlobalProgressLabel As String = "Initializing..."
        Public CancelOperation As Boolean = False

        ' Call this procedure to launch the progress form.
        Public Sub ShowProgressBarInSeparateThread(headerText As String, initialLabel As String)
            Dim t As New Thread(Sub()
                                    ' Create and show the progress form modally.
                                    Dim progressForm As New ProgressForm(headerText, initialLabel)
                                    progressForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        End Sub
    End Module

End Namespace