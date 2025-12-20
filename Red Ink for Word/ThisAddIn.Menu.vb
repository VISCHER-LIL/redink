' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn
    Public Sub AddContextMenu()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

        If MenusAdded Then Return

        ' Remove existing context menus from relevant context menus
        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Return

        If Not VBAModuleWorking() Then Return

        If INIloaded = False Then Return

        MenusAdded = True

        ' List of relevant context menus
        Dim contextMenus As String() = {
        "Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}
        Dim application As Word.Application = Globals.ThisAddIn.Application

        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Check if the menu already exists
                    If Not ContextMenuExists(cb, RIMenu) Then
                        Dim myControl As CommandBarPopup = Nothing
                        Try
                            myControl = CType(cb.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
                        Catch ex As System.Exception
                            ' Handle potential errors
                        End Try

                        If myControl IsNot Nothing Then
                            myControl.Caption = RIMenu
                            myControl.Visible = True
                            myControl.Enabled = True

                            ' Add submenu items
                            AddSubMenuItems(myControl)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Function ContextMenuExists(cb As CommandBar, menuName As String) As Boolean
        For Each ctrl As CommandBarControl In cb.Controls
            If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = menuName Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub AddSubMenuItems(myControl As CommandBarPopup)

        Try
            Dim subControl As CommandBarButton
            Dim wordHelpersMenu As CommandBarPopup
            Dim improveMenu As CommandBarPopup
            Dim subSubControl As CommandBarButton
            Dim shortcutsArray() As String
            Dim shortcutPair() As String
            Dim shortcutDict As New Dictionary(Of String, String) ' Use native .NET Dictionary
            Dim i As Integer

            ' Parse the shortcuts from INI_ShortcutsWordExcel
            shortcutsArray = INI_ShortcutsWordExcel.Split(";"c)

            ' Populate the dictionary
            For i = 0 To shortcutsArray.Length - 1
                If shortcutsArray(i).Contains("=") Then
                    shortcutPair = shortcutsArray(i).Split("="c)
                    shortcutDict(shortcutPair(0).Trim()) = shortcutPair(1).Trim()
                End If
            Next
            myControl.Visible = True

            ' Add menu items and assign shortcuts
            ' The OnAction refers to a Word Macro that has to be loaded as a helper for this to work; it will call up the BridgeSubs methods

            If Not String.IsNullOrWhiteSpace(INI_Language1) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language1
                subControl.FaceId = 6112
                subControl.Visible = True
                subControl.OnAction = "CallInLanguage1"
                If shortcutDict.ContainsKey(subControl.Caption) Then ' Check for key existence
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption) ' Access the value
                End If
            End If

            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112

                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create new submenu "Improve"
            improveMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            improveMenu.Caption = "Improve"

            ' Add submenu items to "Improve"
            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Improve" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallImprove"
            subSubControl.FaceId = 329
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "No Filler Words" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallNoFillers"
            subSubControl.FaceId = 4242
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Friendly" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallFriendly"
            subSubControl.FaceId = 59
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Convincing" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallConvincing"
            subSubControl.FaceId = 343
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If


            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Summarize"
            subControl.OnAction = "CallSummarize"
            subControl.FaceId = 602

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (2nd)"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If


            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Context Search"
            subControl.OnAction = "CallContextSearch"
            subControl.FaceId = 46
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create new submenu "Word helpers"
            wordHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            wordHelpersMenu.Caption = "Word helpers"

            ' Add submenu items to "Word helpers"
            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Self-Compare Selection"
            subSubControl.OnAction = "CallCompareSelectionHalves"
            subSubControl.FaceId = 304
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Accept Format Changes"
            subSubControl.OnAction = "CallAcceptFormatting"
            subSubControl.FaceId = 161
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Markup Time Span"
            subSubControl.OnAction = "CallCalculateUserMarkupTimeSpan"
            subSubControl.FaceId = 33
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Import Text File"
            subSubControl.OnAction = "CallImportTextFile"
            subSubControl.FaceId = 2311
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then

                ' Assign shortcuts using the dictionary
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other", "CallInOther", shortcutDict)
                AssignShortcut("Correct (Markup)", "CallCorrect", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Improve (Markup)", "CallImprove", shortcutDict)
                AssignShortcut("Improve", "CallImprove", shortcutDict)
                AssignShortcut("No Filler Words (Markup)", "CallNoFillers", shortcutDict)
                AssignShortcut("No Filler Words", "CallNoFillers", shortcutDict)
                AssignShortcut("More Friendly (Markup)", "CallFriendly", shortcutDict)
                AssignShortcut("More Friendly", "CallFriendly", shortcutDict)
                AssignShortcut("More Convincing (Markup)", "CallConvincing", shortcutDict)
                AssignShortcut("More Convincing", "CallConvincing", shortcutDict)
                AssignShortcut("Shorten (Markup)", "CallShorten", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize (Markup)", "CallAnonymize", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party (Markup)", "CallSwitchParty", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Summarize", "CallSummarize", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)
                AssignShortcut("Context Search", "CallContextSearch", shortcutDict)

                ' Assign shortcuts for second API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (2nd)", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for submenu "Word helpers"
                AssignShortcut("Self-Compare Selection", "CallCompareSelectionHalves", shortcutDict)
                AssignShortcut("Accept Format Changes", "CallAcceptFormatting", shortcutDict)
                AssignShortcut("Markup Time Span", "CallCalculateUserMarkupTimeSpan", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Import Text File", "CallImportTextFile", shortcutDict)

            End If
        Catch ex As System.Exception

        End Try
    End Sub
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCode As Long
        Try
            ' Check if there is a shortcut assigned for this menu item
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Return ' No shortcut assigned
            End If

            ' Build KeyCode from shortcutKey text
            keyCode = BuildKeyCodeFromText(shortcutKey)

            If keyCode > 0 Then
                Globals.ThisAddIn.Application.CustomizationContext = Globals.ThisAddIn.Application.NormalTemplate
                Globals.ThisAddIn.Application.KeyBindings.Add(KeyCode:=CInt(keyCode), KeyCategory:=WdKeyCategory.wdKeyCategoryMacro, Command:=macro)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully
            ' Debug.WriteLine("Error in AssignShortcut " & ex.Message)
        End Try
    End Sub

    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As Long
        Dim parts() As String
        Dim keysCollection As New List(Of Integer)()
        Dim keyCode As Long = 0

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    Case "CTRL"
                        keysCollection.Add(WdKey.wdKeyControl)
                    Case "SHIFT"
                        keysCollection.Add(WdKey.wdKeyShift)
                    Case "ALT"
                        keysCollection.Add(WdKey.wdKeyAlt)

                ' Map digits directly
                    Case "0"
                        keysCollection.Add(WdKey.wdKey0)
                    Case "1"
                        keysCollection.Add(WdKey.wdKey1)
                    Case "2"
                        keysCollection.Add(WdKey.wdKey2)
                    Case "3"
                        keysCollection.Add(WdKey.wdKey3)
                    Case "4"
                        keysCollection.Add(WdKey.wdKey4)
                    Case "5"
                        keysCollection.Add(WdKey.wdKey5)
                    Case "6"
                        keysCollection.Add(WdKey.wdKey6)
                    Case "7"
                        keysCollection.Add(WdKey.wdKey7)
                    Case "8"
                        keysCollection.Add(WdKey.wdKey8)
                    Case "9"
                        keysCollection.Add(WdKey.wdKey9)

                ' Map function keys directly
                    Case "F1"
                        keysCollection.Add(WdKey.wdKeyF1)
                    Case "F2"
                        keysCollection.Add(WdKey.wdKeyF2)
                    Case "F3"
                        keysCollection.Add(WdKey.wdKeyF3)
                    Case "F4"
                        keysCollection.Add(WdKey.wdKeyF4)
                    Case "F5"
                        keysCollection.Add(WdKey.wdKeyF5)
                    Case "F6"
                        keysCollection.Add(WdKey.wdKeyF6)
                    Case "F7"
                        keysCollection.Add(WdKey.wdKeyF7)
                    Case "F8"
                        keysCollection.Add(WdKey.wdKeyF8)
                    Case "F9"
                        keysCollection.Add(WdKey.wdKeyF9)
                    Case "F10"
                        keysCollection.Add(WdKey.wdKeyF10)
                    Case "F11"
                        keysCollection.Add(WdKey.wdKeyF11)
                    Case "F12"
                        keysCollection.Add(WdKey.wdKeyF12)

                ' Navigation and special keys
                    Case "LEFT"
                        keysCollection.Add(CustomWdKey.wdKeyLeft)
                    Case "RIGHT"
                        keysCollection.Add(CustomWdKey.wdKeyRight)
                    Case "UP"
                        keysCollection.Add(CustomWdKey.wdKeyUp)
                    Case "DOWN"
                        keysCollection.Add(CustomWdKey.wdKeyDown)
                    Case "HOME"
                        keysCollection.Add(WdKey.wdKeyHome)
                    Case "END"
                        keysCollection.Add(WdKey.wdKeyEnd)
                    Case "PAGEUP"
                        keysCollection.Add(WdKey.wdKeyPageUp)
                    Case "PAGEDOWN"
                        keysCollection.Add(WdKey.wdKeyPageDown)
                    Case "ESC"
                        keysCollection.Add(WdKey.wdKeyEsc)
                    Case "TAB"
                        keysCollection.Add(WdKey.wdKeyTab)
                    Case "BACKSPACE"
                        keysCollection.Add(WdKey.wdKeyBackspace)
                    Case "DELETE"
                        keysCollection.Add(WdKey.wdKeyDelete)
                    Case "INSERT"
                        keysCollection.Add(WdKey.wdKeyInsert)
                    Case "SPACE"
                        keysCollection.Add(CustomWdKey.wdKeySpace)

                ' Letters mapped directly
                    Case "A" To "Z"
                        keysCollection.Add([Enum].Parse(GetType(WdKey), "wdKey" & part.Trim().ToUpper()))
                    Case Else
                        ' Unknown key
                        Return 0
                End Select
            Next

            ' Build the KeyCode using Application.BuildKeyCode
            Select Case keysCollection.Count
                Case 1
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0))
                Case 2
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1))
                Case 3
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2))
                Case 4
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2), keysCollection(3))
                Case Else
                    ' Unknown key
                    Return 0
            End Select

            'Debug.WriteLine("Shortcutkey " & shortcutKey & "  Keycode: " & keyCode)

            Return keyCode

        Catch ex As System.Exception
            ' Handle errors gracefully
            Return 0
        End Try
    End Function

    Public Sub RemoveOldContextMenu()
        Dim application As Word.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {
"Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = RIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Handle errors if needed, e.g., logging
                                'Debug.WriteLine($"Error removing control {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next
    End Sub

    Public Sub RemoveVeryOldContextMenu()
        Dim application As Word.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {
"Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
"Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
"Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
"Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
"Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
"Business Card", "Equation Popup", "WordArt Context Menu",
"Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
"Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
}

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = OldRIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Handle errors if needed, e.g., logging
                                'Debug.WriteLine($"Error removing control {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next
    End Sub
End Class
