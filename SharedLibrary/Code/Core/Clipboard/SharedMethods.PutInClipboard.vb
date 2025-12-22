' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods


        Public Shared Sub PutInClipboard(text As String)
            Dim thread As New Threading.Thread(Sub()
                                                   ' Check if the text is RTF formatted
                                                   If text.StartsWith("{\rtf") Then
                                                       ' Set RTF content to the clipboard
                                                       'Clipboard.SetData(DataFormats.Rtf, text)

                                                       Dim plainText As String

                                                       ' Convert RTF to plain text using RichTextBox
                                                       Using rtb As New RichTextBox()
                                                           rtb.Rtf = text
                                                           plainText = rtb.Text
                                                       End Using

                                                       ' Set both RTF and plain text in the clipboard
                                                       Dim dataObj As New DataObject()
                                                       dataObj.SetData(DataFormats.Rtf, text)
                                                       dataObj.SetData(DataFormats.Text, plainText)
                                                       Clipboard.SetDataObject(dataObj, True)

                                                   Else
                                                       ' Set plain text to the clipboard
                                                       Clipboard.SetText(text)
                                                   End If
                                               End Sub)

            ' Ensure the thread is STA (Single Thread Apartment), as required by the clipboard
            thread.SetApartmentState(Threading.ApartmentState.STA)
            thread.Start()
            thread.Join()

        End Sub

    End Class
End Namespace
