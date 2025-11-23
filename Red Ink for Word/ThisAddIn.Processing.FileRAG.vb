' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Public Async Function ConsultLibrary(DoMarkup As Boolean) As Task(Of Boolean)

        Try

            Dim SysPromptTemp As String

            ' Load the library text

            Dim LibFilePath As String = ExpandEnvironmentVariables(INI_Lib_File)

            InfoBox.ShowInfoBox("Loading the library from " & LibFilePath & " ...")

            LibraryText = ReadTextFile(LibFilePath)

            If String.IsNullOrWhiteSpace(LibraryText) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The library file '" & LibFilePath & "' is empty or could not be read.")
                Return False
            End If

            InfoBox.ShowInfoBox("Asking the LLM to search the library based on the instruction ....")

            SysPromptTemp = InterpolateAtRuntime(INI_Lib_Find_SP)

            LibResult = Await LLM(SysPromptTemp, SelectedText, "", "", INI_Lib_Timeout)

            If String.IsNullOrWhiteSpace(LibResult) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to retrieve relevant content from the library. Will abort.")
                Return False
            End If

            InfoBox.ShowInfoBox("Having the LLM apply the result from the library search: " & LibResult, 6)

            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_Lib_Apply_SP)
            End If

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Error in ConsultLibrary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

End Class
