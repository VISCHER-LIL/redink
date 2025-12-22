' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary

    Partial Public Class SharedMethods

        Private Shared clipboardData As Object = Nothing ' Variable to store clipboard content

        Public Shared Sub StoreClipboard()
            If Clipboard.ContainsText() Then
                clipboardData = Clipboard.GetText()
            ElseIf Clipboard.ContainsImage() Then
                clipboardData = Clipboard.GetImage()
            ElseIf Clipboard.ContainsData(DataFormats.Serializable) Then
                clipboardData = Clipboard.GetData(DataFormats.Serializable)
            ElseIf Clipboard.ContainsData(DataFormats.FileDrop) Then
                clipboardData = Clipboard.GetData(DataFormats.FileDrop)
            Else
                clipboardData = Nothing ' No supported data format found
            End If
        End Sub

        Public Shared Sub RestoreClipboard()
            If clipboardData Is Nothing Then Return

            If TypeOf clipboardData Is String Then
                Clipboard.SetText(CStr(clipboardData))
            ElseIf TypeOf clipboardData Is Image Then
                Clipboard.SetImage(CType(clipboardData, Image))
            ElseIf TypeOf clipboardData Is Object Then
                Clipboard.SetData(DataFormats.Serializable, clipboardData)
            End If
        End Sub

    End Class
End Namespace