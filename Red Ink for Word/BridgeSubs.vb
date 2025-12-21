' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: BridgeSubs.vb
' Purpose: Exposes COM-visible entry points that forward automation calls into
'          the corresponding members of Globals.ThisAddIn.
'
' Architecture:
'  - COM Visibility: BridgeSubs is marked <ComVisible(True)> so external callers
'    (e.g., VBA macros or Office UI bindings) can access the add-in commands.
'  - Command Delegation: Every public method synchronously invokes the matching
'    Globals.ThisAddIn member; no additional logic is applied in this layer.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices

''' <summary>
''' COM-visible bridge class exposing Globals.ThisAddIn commands to external callers.
''' </summary>
<ComVisible(True)>
Public Class BridgeSubs
    ''' <summary>
    ''' Invokes Globals.ThisAddIn.InLanguage1().
    ''' </summary>
    Public Sub DoInLanguage1()
        Globals.ThisAddIn.InLanguage1()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.InLanguage2().
    ''' </summary>
    Public Sub DoInLanguage2()
        Globals.ThisAddIn.InLanguage2()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.InOther().
    ''' </summary>
    Public Sub DoInOther()
        Globals.ThisAddIn.InOther()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Correct().
    ''' </summary>
    Public Sub DoCorrect()
        Globals.ThisAddIn.Correct()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Improve().
    ''' </summary>
    Public Sub DoImprove()
        Globals.ThisAddIn.Improve()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.NoFillers().
    ''' </summary>
    Public Sub DoNoFillers()
        Globals.ThisAddIn.NoFillers()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Convincing().
    ''' </summary>
    Public Sub DoConvincing()
        Globals.ThisAddIn.Convincing()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Friendly().
    ''' </summary>
    Public Sub DoFriendly()
        Globals.ThisAddIn.Friendly()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Shorten().
    ''' </summary>
    Public Sub DoShorten()
        Globals.ThisAddIn.Shorten()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Anonymize().
    ''' </summary>
    Public Sub DoAnonymize()
        Globals.ThisAddIn.Anonymize()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.SwitchParty().
    ''' </summary>
    Public Sub DoSwitchParty()
        Globals.ThisAddIn.SwitchParty()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.Summarize().
    ''' </summary>
    Public Sub DoSummarize()
        Globals.ThisAddIn.Summarize()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.FreeStyleNM().
    ''' </summary>
    Public Sub DoFreestyleNM()
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.FreeStyleAM().
    ''' </summary>
    Public Sub DoFreestyleAM()
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.ContextSearch().
    ''' </summary>
    Public Sub DoContextSearch()
        Globals.ThisAddIn.ContextSearch()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.CompareSelectionHalves().
    ''' </summary>
    Public Sub DoCompareSelectionHalves()
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.AcceptFormatting().
    ''' </summary>
    Public Sub DoAcceptFormatting()
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.CalculateUserMarkupTimeSpan().
    ''' </summary>
    Public Sub DoCalculateUserMarkupTimeSpan()
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.RegexSearchReplace().
    ''' </summary>
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.ImportTextFile().
    ''' </summary>
    Public Sub DoImportTextFile()
        Globals.ThisAddIn.ImportTextFile()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.AddContextMenu().
    ''' </summary>
    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

End Class