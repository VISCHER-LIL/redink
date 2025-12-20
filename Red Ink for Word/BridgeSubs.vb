' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

<ComVisible(True)>
Public Class BridgeSubs
    Public Sub DoInLanguage1()
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub DoInLanguage2()
        Globals.ThisAddIn.InLanguage2()
    End Sub

    Public Sub DoInOther()
        Globals.ThisAddIn.InOther()
    End Sub

    Public Sub DoCorrect()
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub DoImprove()
        Globals.ThisAddIn.Improve()
    End Sub

    Public Sub DoNoFillers()
        Globals.ThisAddIn.NoFillers()
    End Sub

    Public Sub DoConvincing()
        Globals.ThisAddIn.Convincing()
    End Sub

    Public Sub DoFriendly()
        Globals.ThisAddIn.Friendly()
    End Sub

    Public Sub DoShorten()
        Globals.ThisAddIn.Shorten()
    End Sub

    Public Sub DoAnonymize()
        Globals.ThisAddIn.Anonymize()
    End Sub

    Public Sub DoSwitchParty()
        Globals.ThisAddIn.SwitchParty()
    End Sub

    Public Sub DoSummarize()
        Globals.ThisAddIn.Summarize()
    End Sub

    Public Sub DoFreestyleNM()
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    Public Sub DoFreestyleAM()
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    Public Sub DoContextSearch()
        Globals.ThisAddIn.ContextSearch()
    End Sub

    Public Sub DoCompareSelectionHalves()
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    Public Sub DoAcceptFormatting()
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    Public Sub DoCalculateUserMarkupTimeSpan()
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Public Sub DoImportTextFile()
        Globals.ThisAddIn.ImportTextFile()
    End Sub
    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

End Class