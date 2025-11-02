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

    'Private ReadOnly RedDragonCodeInstance As New RedDragonCode()

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

    'Private ReadOnly RedDragonCodeInstance As New RedDragonCode()

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


End Class