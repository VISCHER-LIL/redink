' =============================================================================
' File: ThisAddIn.FileHelpers.vb
' Part of: Red Ink for Excel
' Purpose: Provides file path acquisition via drag-and-drop form and content
'          loading utilities for supported document types with optional OCR.
'
' Copyright: David Rosenthal, david.rosenthal@vischer.com
' License: May only be used with an appropriate license (see redink.ai)
'
' Architecture:
' - This file contributes to the partial class ThisAddIn.
' - Relies on external helper methods (e.g., RemoveCR, ShowCustomMessageBox,
'   ExpandEnvironmentVariables, ReadTextFile, ReadRtfAsText, ReadWordDocument,
'   ReadPdfAsText) from SharedLibrary.SharedLibrary.SharedMethods.
' - Uses a modal DragDropForm UI to capture a user-selected file path.
' - Validates file existence before returning or processing.
' - Asynchronous content loading for PDF to allow OCR and user interaction flags.
' - File type dispatch via Select Case on extension for text extraction.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Opens a drag-and-drop selection form, obtains a file path, normalizes it,
    ''' validates existence, and returns the absolute path or empty string on failure.
    ''' </summary>
    ''' <returns>Full file path if found; otherwise empty string.</returns>
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
            filePath = Path.GetFullPath(filePath)
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

    ''' <summary>
    ''' Retrieves file content for a supported file type. Optionally accepts a preset path,
    ''' performs environment variable expansion, invokes a drag-and-drop UI if needed,
    ''' validates existence, and dispatches to appropriate reader. PDF reading supports
    ''' asynchronous OCR and user interaction flags. Returns empty string on errors.
    ''' </summary>
    ''' <param name="optionalFilePath">Optional preset file path before user selection.</param>
    ''' <param name="Silent">If True, suppresses message box error notifications.</param>
    ''' <param name="DoOCR">If True, enables OCR when reading PDF files.</param>
    ''' <param name="AskUser">If True, allows user interaction during PDF processing.</param>
    ''' <returns>Task producing file content as String; empty string on failure.</returns>
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
            filePath = Path.GetFullPath(filePath)
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
End Class