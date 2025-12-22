' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FileHelpers.vb
' Purpose: Supplies helper routines for acquiring a file path and reading supported
'          document types into plain text for downstream processing.
'
' Architecture:
'  - File Selection: Accepts an optional path (with environment expansion) or invokes DragDropForm for user input.
'  - Normalization & Validation: Trims CR characters, resolves absolute paths, and confirms file existence.
'  - Content Extraction: Dispatches to SharedMethods readers based on extension (.txt/.rtf/.doc/.pdf/.pptx, etc.),
'                        optionally enabling OCR for PDF input.
'  - Error Reporting: Surfaces validation or parsing issues through ShowCustomMessageBox unless Silent mode is enabled.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Retrieves textual content from a supported file by using the provided path or prompting the user when no path is supplied.
    ''' </summary>
    ''' <param name="optionalFilePath">File path to load; environment variables are expanded when provided.</param>
    ''' <param name="Silent">Suppresses UI error/notification messages when set to True.</param>
    ''' <param name="DoOCR">Enables OCR while reading PDF files when True.</param>
    ''' <param name="AskUser">Indicates whether PDF processing may prompt the user (passed through to PDF reader).</param>
    ''' <returns>The file content as plain text, or an empty string when the operation is cancelled or fails.</returns>
    Public Async Function GetFileContent(Optional ByVal optionalFilePath As String = Nothing, Optional Silent As Boolean = False, Optional DoOCR As Boolean = False, Optional AskUser As Boolean = True) As Task(Of String)
        Dim filePath As String = ""
        Try

            If optionalFilePath IsNot Nothing Then
                filePath = ExpandEnvironmentVariables(optionalFilePath)
            End If

            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = System.IO.Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                If Not Silent Then ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If

            If Not String.IsNullOrWhiteSpace(filePath) AndAlso IO.File.Exists(filePath) Then
                Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                Dim FromFile As String
                Select Case ext
                    Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                        FromFile = ReadTextFile(filePath)
                    Case ".rtf"
                        FromFile = ReadRtfAsText(filePath)
                    Case ".doc", ".docx"
                        FromFile = ReadWordDocument(filePath)
                    Case ".pdf"
                        FromFile = Await ReadPdfAsText(filePath, True, DoOCR, AskUser, _context)
                    Case ".pptx"
                        FromFile = GetPresentationJson(filePath)
                    Case Else
                        FromFile = "Error: File type not supported."
                End Select
                If FromFile.StartsWith("Error") And Len(FromFile) < 100 And Not Silent Then
                    ShowCustomMessageBox(FromFile)
                    Return ""
                Else
                    Return FromFile
                End If
            End If
        Catch ex As System.Exception
            If Not Silent Then ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Prompts the user for a file via DragDropForm, validates the selection, and returns the absolute path.
    ''' </summary>
    ''' <returns>The normalized file path, or an empty string when the user cancels or validation fails.</returns>
    Public Function GetFileName() As String
        Dim filePath As String = ""
        Try
            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = System.IO.Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If
            Return filePath

        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function
End Class
