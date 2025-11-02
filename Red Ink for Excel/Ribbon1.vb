' Red Ink Ribbon Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 31.10.2025

Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32

Public Class Ribbon1

    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Private Sub ApplyThemeAwareMenuIcon()
        Try
            Dim theme = DetectOfficeTheme()
            Select Case theme
                Case OfficeTheme.Light
                    Menu1.Image = My.Resources.Red_Ink_Logo_Medium
                Case Else
                    Menu1.Image = My.Resources.Red_Ink_Logo
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = My.Resources.Red_Ink_Logo
            Menu1.ShowImage = True
        End Try
    End Sub

    Private Function DetectOfficeTheme() As OfficeTheme
        Const registryPath As String = "Software\Microsoft\Office\16.0\Common"
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then Return OfficeTheme.Unknown

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then Return OfficeTheme.Unknown

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0 ' Colorful
                            Return OfficeTheme.Light
                        Case 1, 2 ' Dark Gray, Black
                            Return OfficeTheme.Dark
                        Case 3 ' White
                            Return OfficeTheme.Light
                        Case 4 ' Use system setting -> resolve via Windows app theme
                            Return If(IsWindowsAppsLightTheme(), OfficeTheme.Light, OfficeTheme.Dark)
                    End Select
                End If
            End Using
        Catch
            ' fall through
        End Try

        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True ' default to light if unknown
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0 ' 1=Light, 0=Dark
                End If
            End Using
        Catch
            ' default to light on error
        End Try
        Return True
    End Function

    Public Async Function RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Correct.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Shorten.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Primlang.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_PrimLang2.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SecLang.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function
    Public Async Function RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Improve.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_FreestyleNM2_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleNM2.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Anonymize.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Sub RI_AdjustHeight_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustHeight.Click
        Globals.ThisAddIn.AdjustHeight()
    End Sub

    Public Sub RI_AdjustLegacyNotes_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AdjustLegacyNotes.Click
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Private Async Function RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_Translate.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InOther()
    End Function

    Private Async Function RI_TranslateF_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_TranslateF.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Async Function RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_FreestyleAM.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Private Async Function RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) As Threading.Tasks.Task Handles RI_SwitchParty.Click
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Regex.Click
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Undo_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Undo.Click
        Globals.ThisAddIn.UndoAction()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat.Click
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Chat2.Click
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Private Sub RI_CSVAnalyze_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CSVAnalyze.Click
        Globals.ThisAddIn.AnalyzeCsvWithLLM()
    End Sub
End Class