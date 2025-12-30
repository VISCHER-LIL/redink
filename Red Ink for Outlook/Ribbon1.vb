' Part of "Red Ink for Outlook"
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
                Case OfficeTheme.Dark
                    Menu1.Image = My.Resources.Red_Ink_Logo
                Case Else
                    Menu1.Image = My.Resources.Red_Ink_Logo_Medium
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = My.Resources.Red_Ink_Logo
            Menu1.ShowImage = True
        End Try
    End Sub

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
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

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub


    Private Sub Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup.Click
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup2.Click
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub
    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub
    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        Globals.ThisAddIn.DefineMyStyle()
    End Sub
    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        Globals.ThisAddIn.HelpMeInky()
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
                    btn.Label = If(i = current, label & " *", label)
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

Public Class Ribbon2


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

    Private Sub Ribbon2_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ApplyThemeAwareMenuIcon()
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
                If key Is Nothing Then Return True
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0
                End If
            End Using
        Catch
        End Try
        Return True
    End Function

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct.Click
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Correct2.Click
        Globals.ThisAddIn.MainMenu("Correct")
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Summarize.Click
        Globals.ThisAddIn.MainMenu("Summarize")
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Shorten.Click
        Globals.ThisAddIn.MainMenu("Shorten")
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Primlang.Click
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrimLang2.Click
        Globals.ThisAddIn.MainMenu("PrimLang")
    End Sub

    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Improve.Click
        Globals.ThisAddIn.MainMenu("Improve")
    End Sub

    Public Sub RI_Freestyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Freestyle.Click
        Globals.ThisAddIn.MainMenu("Freestyle")
    End Sub

    Public Sub RI_Answers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Answers.Click
        Globals.ThisAddIn.MainMenu("Answers")
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Translate.Click
        Globals.ThisAddIn.MainMenu("Translate")
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) Handles Settings.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub


    Private Sub Sumup_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup.Click
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub Sumup2_Click(sender As Object, e As RibbonControlEventArgs) Handles Sumup2.Click
        Globals.ThisAddIn.MainMenu("Sumup")
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_NoFillers.Click
        Globals.ThisAddIn.MainMenu("NoFillers")
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Friendly.Click
        Globals.ThisAddIn.MainMenu("Friendly")
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Convincing.Click
        Globals.ThisAddIn.MainMenu("Convincing")
    End Sub

    Private Sub RI_Clipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Clipboard.Click
        Globals.ThisAddIn.MainMenu("InsertClipboard")
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        Globals.ThisAddIn.MainMenu("ApplyMyStyle")
    End Sub

    Private Sub RI_DefineMyStyle_Click_1(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        Globals.ThisAddIn.HelpMeInky()
    End Sub

End Class