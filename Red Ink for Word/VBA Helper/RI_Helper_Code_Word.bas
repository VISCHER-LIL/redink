Attribute VB_Name = "RI_Helper_Code"

' Helper Code for "Red Ink for Word"
'
' 2.2.2025
'
' These procedures are used if you configure "Red Ink for Word" to assign a key to a particular
' functionality. If that happens, the key is assigned to these macros, which then call the relevant
' procedure within the VSTO Add-in of Red Ink for Word. If you are not allowed to run them, the
' add-in will still work, but you can't use the key shortcuts defined.
'
' All Rights Reserved. david.rosenthal@vischer.com  https://vischer.com/redink

Option Explicit

Const CurrentVersion As Integer = 2
Const AddinName As String = "Red Ink for Word"

Public ModuleRunning As Integer

Sub Autoexec()
    ModuleRunning = CurrentVersion
    CallAddContextMenu
End Sub

' Loopback Function

Public Function CheckAppHelper() As Integer
    CheckAppHelper = ModuleRunning
End Function

Sub CallAddContextMenu()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoAddContextMenu
    End If
End Sub

Sub CallInLanguage1()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoInLanguage1
    End If
End Sub

Sub CallInLanguage2()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoInLanguage2
    End If
End Sub

Sub CallInOther()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoInOther
    End If
End Sub

Sub CallCorrect()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoCorrect
    End If
End Sub

Sub CallImprove()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoImprove
    End If
End Sub

Sub CallNoFillers()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoNoFillers
    End If
End Sub
Sub CallFriendly()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoFriendly
    End If
End Sub
Sub CallConvincing()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoConvincing
    End If
End Sub

Sub CallShorten()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoShorten
    End If
End Sub

Sub CallAnonymize()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoAnonymize
    End If
End Sub

Sub CallSwitchParty()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoSwitchParty
    End If
End Sub

Sub CallSummarize()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoSummarize
    End If
End Sub

Sub CallFreestyleNM()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoFreestyleNM
    End If
End Sub

Sub CallFreestyleAM()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoFreestyleAM
    End If
End Sub

Sub CallContextSearch()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoContextSearch
    End If
End Sub

Sub CallCompareSelectionHalves()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoCompareSelectionHalves
    End If
End Sub

Sub CallAcceptFormatting()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoAcceptFormatting
    End If
End Sub

Sub CallCalculateUserMarkupTimeSpan()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoCalculateUserMarkupTimeSpan
    End If
End Sub
Sub CallRegexSearchReplace()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoRegexSearchReplace
    End If
End Sub
Sub CallImportTextFile()
    Dim addIn As Object
    On Error Resume Next
    Set addIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not addIn Is Nothing Then
        addIn.DoImportTextFile
    End If
End Sub
