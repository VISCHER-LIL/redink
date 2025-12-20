' =============================================================================
' File: ThisAddIn.MenuContext.vb
' Part of: Red Ink for Excel
' Purpose: Adds a custom popup menu (RIMenu) to relevant Excel context menus; populates
'          submenu items and assigns optional keyboard shortcuts derived from INI settings.
'          This requires the VBA helper macros to be loaded and functional.
'
' Copyright: David Rosenthal, david.rosenthal@vischer.com
' License: May only be used with an appropriate license (see redink.ai)
'
' Architecture:
'   - Partial class ThisAddIn manages lifecycle of a context popup named RIMenu.
'   - AddContextMenu: Validates state/config flags; optionally removes old menu; adds new popup
'     to each relevant CommandBar if absent; delegates population to AddSubMenuItems.
'   - ContextMenuExists: Linear scan of controls on a CommandBar to detect existing popup caption.
'   - AddSubMenuItems: Creates buttons plus nested "Excel Helpers" submenu; parses
'     INI_ShortcutsWordExcel into a Dictionary for tooltips and shortcut assignment.
'   - AssignShortcut: Builds Excel OnKey string via BuildKeyCodeFromText and registers macro shortcut.
'   - BuildKeyCodeFromText: Maps tokens (CTRL, SHIFT, ALT, digits, letters A–Z, F1–F12) to Excel
'     OnKey syntax; returns empty string for unknown token.
'   - RemoveOldContextMenu: Deletes prior instances of RIMenu from all relevant context menus.
'   - Error handling: Try/Catch blocks swallow exceptions; comments indicate where errors are handled.
'   - Uses global/state variables: MenusAdded, RemoveMenu, INI_ContextMenu, VBAModuleWorking(),
'     INIloaded, INI_Language1, INI_Language2, INI_SecondAPI, INI_Model_2, INI_ShortcutsWordExcel, RIMenu.
' =============================================================================

Option Strict On
Option Explicit On

Imports Microsoft.Office.Core

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Adds the RIMenu popup to each relevant Excel context menu if configuration permits and menu not already present; sets MenusAdded flag.
    ''' </summary>
    ''' <remarks>Uses MenusAdded, RemoveMenu, INI_ContextMenu, VBAModuleWorking(), INIloaded, RIMenu.</remarks>
    Public Sub AddContextMenu()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

        If MenusAdded Then Exit Sub

        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Exit Sub

        If Not VBAModuleWorking() Then Exit Sub

        If INIloaded = False Then Exit Sub

        MenusAdded = True

        ' List of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}
        Dim application As Excel.Application = Globals.ThisAddIn.Application

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

    ''' <summary>
    ''' Determines whether a popup control with the specified caption already exists on a CommandBar.
    ''' </summary>
    ''' <param name="cb">Target CommandBar.</param>
    ''' <param name="menuName">Caption to match.</param>
    ''' <returns>True if found; otherwise False.</returns>
    Private Function ContextMenuExists(cb As CommandBar, menuName As String) As Boolean
        For Each ctrl As CommandBarControl In cb.Controls
            If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = menuName Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Populates the supplied popup with submenu buttons and an "Excel Helpers" nested submenu; assigns tooltips and shortcuts based on INI_ShortcutsWordExcel.
    ''' </summary>
    ''' <param name="myControl">Parent CommandBarPopup to receive controls.</param>
    ''' <remarks>Uses INI_Language1, INI_Language2, INI_SecondAPI, INI_Model_2, INI_ShortcutsWordExcel.</remarks>
    Public Sub AddSubMenuItems(myControl As CommandBarPopup)

        Try
            Dim subControl As CommandBarButton
            Dim excelHelpersMenu As CommandBarPopup
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
                    shortcutDict(Trim(shortcutPair(0))) = Trim(shortcutPair(1))
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
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption) ' Access the value
                End If
            End If
            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112

                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (text)"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (cells)"
            subControl.OnAction = "CallInOtherFormulas"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct"
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Write Neatly"
            subControl.OnAction = "CallNeatly"
            subControl.FaceId = 162

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten"
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize"
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party"
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (" & INI_Model_2 & ")"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If

            End If

            ' Create new submenu "Excel helpers"
            excelHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            excelHelpersMenu.Caption = "Excel Helpers"

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Cell Height"
            subSubControl.OnAction = "CallAdjustHeight"
            subSubControl.FaceId = 1647

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If


            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Size of Notes"
            subSubControl.OnAction = "CallAdjustLegacyNotes"
            subSubControl.FaceId = 1996

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then
                ' Assign shortcuts using the dictionary
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other (text)", "CallInOther", shortcutDict)
                AssignShortcut("To Other (cells)", "CallInOther", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Write Neatly", "CallImprove", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)

                ' Assign shortcuts for second API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (" & INI_Model_2 & ")", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for submenu "Excel helpers"
                AssignShortcut("Adjust Cell Height", "CallAdjustheight", shortcutDict)
                AssignShortcut("Adjust Size of Notes", "CallAdjustLegacyNotes", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
            End If
        Catch ex As System.Exception

        End Try
    End Sub

    ''' <summary>
    ''' Registers a keyboard shortcut for a macro using Application.OnKey when a mapping exists in shortcutDict.
    ''' </summary>
    ''' <param name="controlName">Caption used as lookup key.</param>
    ''' <param name="macro">Macro procedure name.</param>
    ''' <param name="shortcutDict">Dictionary mapping captions to shortcut text.</param>
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCombination As String
        Try
            ' Check if there is a shortcut assigned for this control
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Exit Sub ' No shortcut assigned
            End If

            ' Build the key combination string from the shortcutKey text
            keyCombination = BuildKeyCodeFromText(shortcutKey)

            If Not String.IsNullOrEmpty(keyCombination) Then
                ' Assign the shortcut key to the macro in Excel using Application.OnKey
                Globals.ThisAddIn.Application.OnKey(keyCombination, macro)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully
            ' Debug.WriteLine("Error in AssignShortcut: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Converts a textual shortcut descriptor (tokens separated by '-') into an Excel OnKey sequence.
    ''' </summary>
    ''' <param name="shortcutKey">Descriptor containing CTRL, SHIFT, ALT, digits, letters, or function keys.</param>
    ''' <returns>Excel OnKey sequence or empty string if any token unknown.</returns>
    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As String
        Dim parts() As String
        Dim keysCollection As New List(Of String)()
        Dim keyCombination As String = ""

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    Case "CTRL"
                        keysCollection.Add("^") ' Control key representation in Excel
                    Case "SHIFT"
                        keysCollection.Add("+") ' Shift key representation in Excel
                    Case "ALT"
                        keysCollection.Add("%") ' Alt key representation in Excel

                ' Map digits directly
                    Case "0" To "9"
                        keysCollection.Add(part.Trim())

                ' Map function keys directly
                    Case "F1" To "F12"
                        keysCollection.Add(part.Trim())

                ' Letters mapped directly
                    Case "A" To "Z"
                        keysCollection.Add(part.Trim().ToUpper())

                    Case Else
                        ' Unknown key
                        Return ""
                End Select
            Next

            ' Combine the keys into a single shortcut string for VBA
            keyCombination = String.Join("", keysCollection)

            Return keyCombination

        Catch ex As System.Exception
            ' Handle errors gracefully
            ' Debug.WriteLine("Error in BuildKeyCodeFromText: " & ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Removes existing instances of the RIMenu popup from all relevant Excel context menus.
    ''' </summary>
    ''' <remarks>Iterates contextMenus array; deletes matching popup controls.</remarks>
    Public Sub RemoveOldContextMenu()
        Dim application As Excel.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}

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
                                'Debug.WriteLine($"Error removing control: {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next

    End Sub

End Class