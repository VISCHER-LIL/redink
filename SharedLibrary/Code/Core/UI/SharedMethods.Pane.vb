' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Public Shared Property TaskPanes As CustomTaskPaneCollection

        Public Shared Sub Initialize(panes As CustomTaskPaneCollection)
            TaskPanes = panes
        End Sub


        Public Class PaneManager

            Private Shared CurrentCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

            Public Shared Async Function ShowMyPane(
        introLine As String,
        bodyText As String,
        finalRemark As String,
        header As String,
        Optional noRTF As Boolean = False,
        Optional insertMarkdown As Boolean = False,
        Optional mergeCallback As IntelligentMergeCallback = Nothing,
        Optional PreserveLiterals As Boolean = False
    ) As Task(Of String)

                If TaskPanes Is Nothing Then
                    Return String.Empty
                End If

                Dim result = Await PaneManager.ShowCustomPane(
            TaskPanes,
            introLine,
            bodyText,
            finalRemark,
            header,
            noRTF,
            insertMarkdown,
            mergeCallback,
            PreserveLiterals
        )

                Return result
            End Function

            Public Shared Function ShowCustomPane(
        XtaskPanes As Microsoft.Office.Tools.CustomTaskPaneCollection,
        introLine As String,
        bodyText As String,
        finalRemark As String,
        header As String,
        Optional noRTF As Boolean = False,
        Optional insertMarkdown As Boolean = False,
        Optional mergeCallback As IntelligentMergeCallback = Nothing,
        Optional PreserveLiterals As Boolean = False
    ) As Task(Of String)

                If CurrentCustomTaskPane IsNot Nothing Then
                    Try
                        CurrentCustomTaskPane.Visible = False
                        XtaskPanes.Remove(CurrentCustomTaskPane)
                    Catch
                    End Try
                    CurrentCustomTaskPane = Nothing
                End If

                Dim ctrl As New CustomPaneControl() With {.MergeCallback = mergeCallback}
                Dim pane = XtaskPanes.Add(ctrl, header)
                pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                pane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone
                pane.Width = If(My.Settings.PaneWidth > 0, My.Settings.PaneWidth, Default_PaneWidth)
                pane.Visible = True
                ctrl.ParentPane = pane
                CurrentCustomTaskPane = pane

                Return ctrl.ShowPane(introLine, bodyText, finalRemark, header, noRTF, insertMarkdown, PreserveLiterals)
            End Function

        End Class


        Public Class CustomPaneControl
            Inherits System.Windows.Forms.UserControl

            Private Const EM_AUTOURLDETECT As Integer = &H45A

            Private tcs As System.Threading.Tasks.TaskCompletionSource(Of String)
            Private originalText As String
            Private NoRTF As Boolean
            Private InsertMarkdown As Boolean
            Private NoMerge As Boolean

            Private introLabel As System.Windows.Forms.Label
            Private toolStrip As System.Windows.Forms.ToolStrip
            Private bodyTextBox As System.Windows.Forms.RichTextBox
            Private finalRemarkLabel As System.Windows.Forms.Label
            Private btnTable As System.Windows.Forms.TableLayoutPanel
            Private btnMerge As System.Windows.Forms.Button
            Private btnSelected As System.Windows.Forms.Button
            Private btnMark As System.Windows.Forms.Button
            Private btnCancel As System.Windows.Forms.Button
            Private toolTip As System.Windows.Forms.ToolTip

            Private ReadOnly urlRegex As System.Text.RegularExpressions.Regex =
            New System.Text.RegularExpressions.Regex("https?://[^\s<>()]+",
                System.Text.RegularExpressions.RegexOptions.Compiled Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            Public Property MergeCallback As IntelligentMergeCallback
            Public Property ParentPane As Microsoft.Office.Tools.CustomTaskPane

            <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
            Private Shared Function SendMessage(hWnd As System.IntPtr, msg As Integer, wParam As System.IntPtr, lParam As System.IntPtr) As System.IntPtr
            End Function

            Public Sub New()
                MyBase.New()
                InitializeComponent()
            End Sub

            ' -------------------- Initialization --------------------
            Private Sub InitializeComponent()
                Const padding As Integer = 10
                NoMerge = String.IsNullOrEmpty(SharedMethods.SP_MergePrompt_Cached)

                Dim stdFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                Me.Font = stdFont
                Me.Dock = System.Windows.Forms.DockStyle.Fill

                toolTip = New System.Windows.Forms.ToolTip() With {.ShowAlways = True}

                Dim tbl As New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New System.Windows.Forms.Padding(padding)
            }
                tbl.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
                tbl.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                tbl.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                tbl.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
                tbl.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                tbl.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
                Me.Controls.Add(tbl)

                introLabel = New System.Windows.Forms.Label() With {
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Font = stdFont,
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }
                tbl.Controls.Add(introLabel, 0, 0)

                toolStrip = New System.Windows.Forms.ToolStrip() With {
                .GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Padding = New System.Windows.Forms.Padding(0),
                .AutoSize = False,
                .Height = 26,
                .RenderMode = System.Windows.Forms.ToolStripRenderMode.System
            }

                Dim buttonFont As New System.Drawing.Font(stdFont.FontFamily, 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                For Each sym In New String() {"B", "I", "U", "•"}
                    Dim style As System.Drawing.FontStyle =
                    If(sym = "B", System.Drawing.FontStyle.Bold,
                    If(sym = "I", System.Drawing.FontStyle.Italic,
                    If(sym = "U", System.Drawing.FontStyle.Underline, System.Drawing.FontStyle.Regular)))
                    Dim tsb As New System.Windows.Forms.ToolStripButton(sym) With {
                    .Font = New System.Drawing.Font(buttonFont, style),
                    .Name = "tsb" & sym,
                    .DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text,
                    .AutoSize = False,
                    .Width = 28,
                    .Height = 24,
                    .Margin = New System.Windows.Forms.Padding(0),
                    .Padding = New System.Windows.Forms.Padding(0),
                    .TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay
                }
                    AddHandler tsb.Click, AddressOf ToolStripButton_Click
                    toolStrip.Items.Add(tsb)
                Next
                tbl.Controls.Add(toolStrip, 0, 1)

                bodyTextBox = New System.Windows.Forms.RichTextBox() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .DetectUrls = True,
                .HideSelection = False,
                .WordWrap = True
            }
                bodyTextBox.Font = New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

                AddHandler bodyTextBox.LinkClicked, AddressOf BodyTextBox_LinkClicked
                AddHandler bodyTextBox.KeyDown, AddressOf BodyTextBox_KeyDown
                tbl.Controls.Add(bodyTextBox, 0, 2)

                finalRemarkLabel = New System.Windows.Forms.Label() With {
                .AutoSize = True,
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .Font = New System.Drawing.Font(stdFont.FontFamily, stdFont.Size - 1.0F, stdFont.Style),
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }
                tbl.Controls.Add(finalRemarkLabel, 0, 3)

                btnTable = New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 1,
                .Margin = New System.Windows.Forms.Padding(0)
            }
                For i As Integer = 1 To 4
                    btnTable.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
                Next

                If NoMerge Then
                    btnMerge = New System.Windows.Forms.Button() With {.Text = "Apply selection"}
                    btnSelected = New System.Windows.Forms.Button() With {.Text = "Copy selection"}
                    btnMark = New System.Windows.Forms.Button() With {.Text = "Apply all", .Visible = False}
                    btnCancel = New System.Windows.Forms.Button() With {.Text = "Close"}
                Else
                    btnMerge = New System.Windows.Forms.Button() With {.Text = "Merge selection"}
                    btnSelected = New System.Windows.Forms.Button() With {.Text = "Copy selection"}
                    btnMark = New System.Windows.Forms.Button() With {.Text = "Insert && close", .Visible = False}
                    btnCancel = New System.Windows.Forms.Button() With {.Text = "Close"}
                End If

                Dim addBtn =
                Sub(b As System.Windows.Forms.Button, col As Integer, tip As String)
                    b.Dock = System.Windows.Forms.DockStyle.Fill
                    b.AutoEllipsis = True
                    b.Margin = New System.Windows.Forms.Padding(2)
                    AddHandler b.Click, AddressOf Button_Click
                    toolTip.SetToolTip(b, tip)
                    btnTable.Controls.Add(b, col, 0)
                End Sub

                If NoMerge Then
                    addBtn(btnMerge, 0, "")
                    addBtn(btnSelected, 1, "")
                    addBtn(btnMark, 2, "")
                    addBtn(btnCancel, 3, "")
                Else
                    addBtn(btnMerge, 0, "")
                    addBtn(btnSelected, 1, "")
                    addBtn(btnMark, 2, "Insert original text and close")
                    addBtn(btnCancel, 3, "")
                End If

                tbl.Controls.Add(btnTable, 0, 4)
                AddHandler Me.Resize, AddressOf OnControlResize
            End Sub


            Private Sub BodyTextBox_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                ' Intercept Copy (Ctrl+C) and Copy (Ctrl+Insert) to exclude trailing NBSPs when entire text is selected.
                If e.Control AndAlso (e.KeyCode = Keys.C) OrElse (e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.Insert) Then
                    Try
                        If Not NoRTF Then
                            SharedMethods.CopySelectionExcludingTrailingNbsp(bodyTextBox)
                        Else
                            ' Plain text mode: default behavior is fine; use PutInClipboard to preserve existing behavior.
                            If bodyTextBox.SelectionLength > 0 Then
                                SharedMethods.PutInClipboard(bodyTextBox.SelectedText)
                            Else
                                SharedMethods.PutInClipboard(bodyTextBox.Text)
                            End If
                        End If
                        e.Handled = True
                    Catch
                        ' fallback to default behavior if anything goes wrong
                    End Try
                End If

                ' Intercept Select All (Ctrl+A) only to avoid changing default selection behavior.
                ' We don't alter Ctrl+A here; copy-time trimming handles the clipboard contents.
            End Sub


            ' -------------------- Public API --------------------
            Public Function ShowPane(introLine As String,
                                 bodyText As String,
                                 finalRemark As String,
                                 header As String,
                                 Optional noRTF As Boolean = False,
                                 Optional insertMarkdown As Boolean = False,
                                 Optional PreserveLiterals As Boolean = False) As System.Threading.Tasks.Task(Of String)

                tcs = New System.Threading.Tasks.TaskCompletionSource(Of String)()
                Me.NoRTF = noRTF
                Me.InsertMarkdown = insertMarkdown
                originalText = bodyText

                ' Make pane visible first
                If ParentPane IsNot Nothing Then
                    ParentPane.Width = SharedMethods.Default_PaneWidth
                    ParentPane.Visible = True
                End If

                ' Ensure control is created
                If Not bodyTextBox.IsHandleCreated Then
                    bodyTextBox.CreateControl()
                End If

                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(50)

                introLabel.Text = introLine
                finalRemarkLabel.Text = finalRemark
                btnMark.Visible = insertMarkdown

                Try
                    If noRTF Then
                        bodyTextBox.Text = bodyText
                    Else
                        Dim rtf As String = MarkdownToRtfConverter.Convert(bodyText, PreserveLiterals)
                        bodyTextBox.Rtf = If(rtf, String.Empty)

                        ' append non-breaking spaces equal to total length of hyperlinks found in the RTF
                        AppendNbspForHyperlinks(bodyTextBox, rtf)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[CustomPaneControl] Error setting content: {ex.Message}")
                    bodyTextBox.Text = bodyText
                End Try

                ' Ensure URL detection is enabled by the control
                bodyTextBox.DetectUrls = True

                System.Windows.Forms.Application.DoEvents()

                bodyTextBox.Select(0, 0)

                Try
                    bodyTextBox.Focus()
                Catch
                End Try

                Return tcs.Task
            End Function

            Public ReadOnly Property BodyBox As System.Windows.Forms.RichTextBox
                Get
                    Return bodyTextBox
                End Get
            End Property

            ' -------------------- Handlers --------------------
            Private Sub ToolStripButton_Click(sender As Object, e As System.EventArgs)
                Dim tsb = DirectCast(sender, System.Windows.Forms.ToolStripButton)
                If bodyTextBox.SelectionLength > 0 AndAlso bodyTextBox.SelectionFont IsNot Nothing Then
                    Select Case tsb.Name
                        Case "tsbB"
                            bodyTextBox.SelectionFont =
                            New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor System.Drawing.FontStyle.Bold)
                        Case "tsbI"
                            bodyTextBox.SelectionFont =
                            New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor System.Drawing.FontStyle.Italic)
                        Case "tsbU"
                            bodyTextBox.SelectionFont =
                            New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor System.Drawing.FontStyle.Underline)
                        Case "tsb•"
                            bodyTextBox.SelectionIndent = If(bodyTextBox.SelectionIndent = 20, 0, 20)
                            bodyTextBox.SelectionBullet = Not bodyTextBox.SelectionBullet
                            bodyTextBox.BulletIndent = If(bodyTextBox.BulletIndent = 15, 0, 15)
                    End Select
                End If
            End Sub

            Private Sub BodyTextBox_LinkClicked(sender As Object, e As System.Windows.Forms.LinkClickedEventArgs)
                System.Diagnostics.Debug.WriteLine($"CustomPaneControl: Received LinkClicked event with URL: {e.LinkText}")
                SafeOpenLink(e.LinkText)
            End Sub

            Private Sub Button_Click(sender As Object, e As System.EventArgs)
                Dim btn = DirectCast(sender, System.Windows.Forms.Button)
                Dim result As String = String.Empty

                If btn Is btnSelected Then
                    If NoRTF Then
                        SharedMethods.PutInClipboard(bodyTextBox.SelectedText)
                    Else
                        ' Use helper that will perform a safe copy (preserving RTF) but exclude trailing nbsp when entire text is selected
                        SharedMethods.CopySelectionExcludingTrailingNbsp(bodyTextBox)
                    End If
                    Return
                End If

                If btn Is btnMerge Then
                    Dim cb = MergeCallback
                    If cb IsNot Nothing Then cb.Invoke(bodyTextBox.SelectedText)
                    Return
                End If

                If btn Is btnMark Then
                    If NoMerge Then
                        Dim cb = MergeCallback
                        If cb IsNot Nothing Then cb.Invoke(bodyTextBox.Text)
                        Return
                    Else
                        result = "Markdown"
                    End If
                End If

                tcs.TrySetResult(result)
                HidePane()
            End Sub

            ' -------------------- Helpers --------------------
            Private Sub OnControlResize(sender As Object, e As System.EventArgs)
                If ParentPane IsNot Nothing Then
                    My.Settings.PaneWidth = ParentPane.Width
                    My.Settings.Save()
                End If
            End Sub

            Private Sub SafeOpenLink(url As String)
                Try
                    System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(url) With {.UseShellExecute = True})
                Catch
                End Try
            End Sub

            Private Sub HidePane()
                Try
                    If ParentPane IsNot Nothing Then ParentPane.Visible = False
                Catch
                End Try
            End Sub

            ' -------------------- Dispose --------------------
            Protected Overrides Sub Dispose(disposing As Boolean)
                Try
                    If disposing Then
                        toolTip?.Dispose()
                        toolStrip?.Dispose()
                    End If
                Finally
                    MyBase.Dispose(disposing)
                End Try
            End Sub

        End Class

        Public Shared Sub AppendNbspForHyperlinks(targetBox As RichTextBox, rtf As String)
            If targetBox Is Nothing OrElse String.IsNullOrEmpty(rtf) Then Return

            Dim totalLength As Integer = 0

            ' 1) Find HYPERLINK field URLs in RTF (fldinst HYPERLINK "url")
            Dim fldRegex As New Regex("HYPERLINK\s+""([^""]+)""", RegexOptions.IgnoreCase)
            For Each m As Match In fldRegex.Matches(rtf)
                totalLength += m.Groups(1).Value.Length
            Next

            ' 2) Also catch plain http/https occurrences that may appear in the RTF
            Dim urlRegex As New Regex("https?://[^\s\\\}\)\]""<>]+", RegexOptions.IgnoreCase)
            For Each m As Match In urlRegex.Matches(rtf)
                totalLength += m.Value.Length
            Next

            If totalLength <= 0 Then Return

            ' Insert one paragraph and then neutral NBSPs so they are not clickable / styled as links.
            Try
                ' Move caret to end
                targetBox.SelectionStart = targetBox.TextLength

                ' Insert single paragraph break (one CRLF)
                ' Using SelectedText to safely append without manual RTF manipulation.
                ' Temporarily disable URL auto-detection so insertion cannot become an auto-link.
                Dim prevDetect As Boolean = targetBox.DetectUrls
                targetBox.DetectUrls = False

                targetBox.SelectedText = vbCrLf

                ' Ensure insertion formatting is neutral: set selection font/color to control defaults
                Try
                    If targetBox.SelectionFont IsNot Nothing Then
                        targetBox.SelectionFont = New System.Drawing.Font(targetBox.Font, System.Drawing.FontStyle.Regular)
                    Else
                        targetBox.SelectionFont = targetBox.Font
                    End If
                Catch
                    ' ignore font-setting errors
                End Try
                Try
                    targetBox.SelectionColor = targetBox.ForeColor
                Catch
                End Try

                ' Insert NBSP characters (U+00A0)
                targetBox.SelectionStart = targetBox.TextLength
                targetBox.SelectedText = New String(ChrW(&HA0), totalLength)

                ' Restore URL detection and scroll to caret
                targetBox.DetectUrls = prevDetect
                targetBox.ScrollToCaret()
            Catch ex As Exception
                ' swallow - resilient helper
            End Try
        End Sub

        Public Shared Sub CopySelectionExcludingTrailingNbsp(rtb As RichTextBox)
            If rtb Is Nothing Then Return

            Try
                Dim hasSel As Boolean = rtb.SelectionLength > 0
                Dim rtfSource As String = If(hasSel, rtb.SelectedRtf, rtb.Rtf)
                Dim plainSource As String = If(hasSel, rtb.SelectedText, rtb.Text)

                ' Count trailing NBSP (U+00A0) in the chosen plain text
                Dim trailing As Integer = 0
                For i As Integer = plainSource.Length - 1 To 0 Step -1
                    If AscW(plainSource(i)) = &HA0 Then
                        trailing += 1
                    Else
                        Exit For
                    End If
                Next

                ' Trim plain text by removing trailing NBSP only
                Dim plainTrimmed As String = If(trailing > 0 AndAlso plainSource.Length >= trailing,
                                                plainSource.Substring(0, plainSource.Length - trailing),
                                                plainSource)

                ' Trim RTF by loading into a temporary RichTextBox and deleting trailing NBSP chars
                Dim rtfTrimmed As String = rtfSource
                If trailing > 0 AndAlso Not String.IsNullOrEmpty(rtfSource) Then
                    Using tmp As New RichTextBox()
                        tmp.Rtf = rtfSource
                        If tmp.TextLength >= trailing Then
                            tmp.Select(tmp.TextLength - trailing, trailing)
                            tmp.SelectedText = "" ' delete trailing NBSPs
                            rtfTrimmed = tmp.Rtf
                        End If
                    End Using
                End If

                ' Put both formats in the clipboard
                Dim dataObj As New DataObject()
                dataObj.SetData(DataFormats.Rtf, rtfTrimmed)
                dataObj.SetData(DataFormats.UnicodeText, plainTrimmed)
                dataObj.SetData(DataFormats.Text, plainTrimmed)
                Clipboard.SetDataObject(dataObj, True)
            Catch
                ' Fallback: default copy behavior
                Try
                    If rtb.SelectionLength > 0 Then
                        rtb.Copy()
                    Else
                        Clipboard.SetText(rtb.Text)
                    End If
                Catch
                End Try
            End Try
        End Sub


    End Class

End Namespace
