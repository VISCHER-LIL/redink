' Red Ink Ribbon Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 8.9.2025

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
                    Menu1.Image = My.Resources.Red_Ink_Logo_Dark
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
' 2025-10-27-RK: Software\Microsoft\Office\16.0\Common is the standard registry hive for Office 2016/Office 365 (all modern “16.x” builds, including current Microsoft 365). If you ever need to support older Office releases (15.0 = Office 2013, 14.0 = Office 2010, etc.), you’d have to probe the other hives as well. For today’s target (O365/Office 2016+), 16.0 is the right—and necessary—value.
		
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then
                    Return OfficeTheme.Unknown
                End If

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then
                    Return OfficeTheme.Unknown
                End If

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0, 4
                            Return OfficeTheme.Light
                        Case 1, 2
                            Return OfficeTheme.Dark
                    End Select
                End If
            End Using
        Catch
            Return OfficeTheme.Unknown
        End Try

        Return OfficeTheme.Unknown
    End Function

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct.Click
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct2.Click
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Summarize.Click
        Globals.ThisAddIn.Summarize()
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Shorten.Click
        Globals.ThisAddIn.Shorten()
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Primlang.Click
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_PrimLang2.Click
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SecLang.Click
        Globals.ThisAddIn.InLanguage2()
    End Sub
    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Improve.Click
        Globals.ThisAddIn.Improve()
    End Sub

    Public Sub RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleNM.Click
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    Public Sub RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Anonymize.Click
        Globals.ThisAddIn.Anonymize()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat.Click
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat2.Click
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_TimeSpan_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_TimeSpan.Click
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub
    Public Sub RI_AcceptFormat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_AcceptFormat.Click
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Translate.Click
        Globals.ThisAddIn.InOther()
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) 'Handles Settings.Click
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleAM.Click
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    Private Sub RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SwitchParty.Click
        Globals.ThisAddIn.SwitchParty()
    End Sub

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Regex.Click
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Import_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        Globals.ThisAddIn.ImportTextFile()
    End Sub

    Private Sub RI_Halves_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Halves.Click
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    Private Sub RI_Search_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        Globals.ThisAddIn.ContextSearch()
    End Sub

    Private Sub Easteregg_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.EasterEgg()
    End Sub

    Private Sub RI_Transcriptor_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.Transcriptor()
    End Sub

    Private Sub RI_Explain_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.Explain()
    End Sub

    Private Sub RI_SuggestTitles_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.SuggestTitles()
    End Sub

    Private Sub RI_CreatePodcast_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.CreatePodcast()
    End Sub

    Private Sub RI_CreateAudio_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.CreateAudio()
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.NoFillers()
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.Friendly()
    End Sub
    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.Convincing()
    End Sub
    Private Sub RI_SpecialModel_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.SpecialModel()
    End Sub

    Private Sub RI_Anonymization_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.AnonymizeSelection()
    End Sub

    Private Sub RI_InsertClipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_InsertClipboard.Click
        Globals.ThisAddIn.InsertClipboard()
    End Sub

    Private Sub RI_BallooMergePart_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePart.Click
        Globals.ThisAddIn.BalloonMerge(False, True)
    End Sub

    Private Sub RI_BalloonMergeFull_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFull.Click
        Globals.ThisAddIn.BalloonMerge(True, True)
    End Sub

    Private Sub RI_BalloonMergePartPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePartPrompt.Click
        Globals.ThisAddIn.BalloonMerge(False, False)
    End Sub

    Private Sub RI_BalloonMergeFullPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFullPrompt.Click
        Globals.ThisAddIn.BalloonMerge(True, False)
    End Sub

    Private Sub RI_FreestyleRepeat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FreestyleRepeat.Click
        Globals.ThisAddIn.FreeStyleRepeat()
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        Globals.ThisAddIn.ApplyMyStyle()
    End Sub

    Private Sub RI_DefineMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_DocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DocCheck.Click
        Globals.ThisAddIn.RunDocCheck()
    End Sub

    Private Sub RI_FindClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FindClause.Click
        Globals.ThisAddIn.FindClause()
    End Sub

    Private Sub RI_AddClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AddClause.Click
        Globals.ThisAddIn.AddClause()
    End Sub

    Private Sub RI_WebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_WebAgent.Click
        Globals.ThisAddIn.WebAgent()
    End Sub

    Private Sub RI_EditWebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_EditWebAgent.Click
        Globals.ThisAddIn.CreateModifyWebAgentScript()
    End Sub

    Private Sub RI_Markdown_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Markdown.Click
        Globals.ThisAddIn.ConvertMarkdownToWord()
    End Sub
End Class