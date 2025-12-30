' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32
Imports SharedLibrary.SharedLibrary
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

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_Extractor_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Extractor.Click
        Globals.ThisAddIn.FactExtraction()
    End Sub

    Private Sub RI_Renamer_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Renamer.Click
        Globals.ThisAddIn.RenameDocumentsWithAi()
    End Sub

    Private Sub RI_Model1_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model1.Click
        Globals.ThisAddIn.SelectModel(1)
    End Sub

    Private Sub RI_Model2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model2.Click
        Globals.ThisAddIn.SelectModel(2)
    End Sub

    Private Sub RI_Model3_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model3.Click
        Globals.ThisAddIn.SelectModel(3)
    End Sub

    Private Sub RI_Model4_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model4.Click
        Globals.ThisAddIn.SelectModel(4)
    End Sub

    Private Sub RI_Model5_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model5.Click
        Globals.ThisAddIn.SelectModel(5)
    End Sub

    Private Sub RI_Model6_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model6.Click
        Globals.ThisAddIn.SelectModel(6)
    End Sub

    Private Sub RI_Model7_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model7.Click
        Globals.ThisAddIn.SelectModel(7)
    End Sub

    Private Sub RI_Model8_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model8.Click
        Globals.ThisAddIn.SelectModel(8)
    End Sub

    Private Sub RI_Model9_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model9.Click
        Globals.ThisAddIn.SelectModel(9)
    End Sub

    Private Sub RI_Model10_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Model10.Click
        Globals.ThisAddIn.SelectModel(10)
    End Sub

    Public Sub UpdateModelsMenu()
        Try
            Dim available = ModelConfigManager.GetAvailableModels()
            Dim current = ModelConfigManager.GetCurrentModelNumber()

            For i = 1 To 10
                Dim btn = GetModelButton(i)
                If btn Is Nothing Then Continue For

                If available.Contains(i) Then
                    btn.Visible = True
                    Dim label = ModelConfigManager.GetModelDisplayName(i)
                    btn.Label = If(i = current, $"{label} ??????", label)
                Else
                    btn.Visible = False
                End If
            Next
        Catch
            ' non-critical
        End Try
    End Sub

    Private Function GetModelButton(i As Integer) As RibbonButton
        Select Case i
            Case 1 : Return RI_Model1
            Case 2 : Return RI_Model2
            Case 3 : Return RI_Model3
            Case 4 : Return RI_Model4
            Case 5 : Return RI_Model5
            Case 6 : Return RI_Model6
            Case 7 : Return RI_Model7
            Case 8 : Return RI_Model8
            Case 9 : Return RI_Model9
            Case 10 : Return RI_Model10
        End Select
        Return Nothing
    End Function
End Class