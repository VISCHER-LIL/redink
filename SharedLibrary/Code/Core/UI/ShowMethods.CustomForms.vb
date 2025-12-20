' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms

Namespace SharedLibrary
    Partial Public Class SharedMethods


        <DllImport("user32.dll", CharSet:=CharSet.Auto)>
        Private Shared Function SendMessage(
                    ByVal hWnd As IntPtr,
                    ByVal msg As Integer,
                    ByVal wParam As IntPtr,
                    ByVal lParam As IntPtr
                ) As IntPtr
        End Function

        ' Add this import at the class level (with the other imports)
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function

        ' helper method to detect the calling application
        Private Shared Function GetOfficeApplicationHwnd() As IntPtr
            ' Try Word first
            Dim hwnd As IntPtr = FindWindow("OpusApp", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

            ' Try Excel
            hwnd = FindWindow("XLMAIN", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

            ' Try Outlook
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
            If hwnd <> IntPtr.Zero Then Return hwnd

            Return IntPtr.Zero
        End Function


        Public Shared Function ShowSelectionForm(
                                            prompt As String,
                                            title As String,
                                            options As IEnumerable(Of String)
                                        ) As String

            Dim selectedOption As String = "ESC"

            ' Form konfigurieren und DPI‑Unterstützung
            Dim inputForm As New System.Windows.Forms.Form() With {
        .Text = title,
        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
        .StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
        .MinimizeBox = False,
        .MaximizeBox = False,
        .ShowInTaskbar = False,
        .KeyPreview = True,
        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
        .ClientSize = New System.Drawing.Size(450, 320),
        .MinimumSize = New System.Drawing.Size(450, 240)
    }
            inputForm.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            ' Logo als Icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            inputForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Haupt-Layout: Prompt, ListBox, Buttons
            Dim layout As New System.Windows.Forms.TableLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3
    }
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100))
            layout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            inputForm.Controls.Add(layout)

            ' Prompt-Label mit automatischem Zeilenumbruch
            Dim labelPrompt As New System.Windows.Forms.Label() With {
        .Text = prompt,
        .AutoSize = True,
        .MaximumSize = New System.Drawing.Size(inputForm.ClientSize.Width - 40, 0),
        .Margin = New System.Windows.Forms.Padding(20, 20, 20, 10),
        .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    }
            layout.Controls.Add(labelPrompt, 0, 0)

            ' ListBox mit Padding
            Dim listPanel As New System.Windows.Forms.Panel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .Padding = New System.Windows.Forms.Padding(20)
    }
            layout.Controls.Add(listPanel, 0, 1)

            Dim listBoxOptions As New System.Windows.Forms.ListBox() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .SelectionMode = System.Windows.Forms.SelectionMode.One
    }
            listBoxOptions.Items.AddRange(options.ToArray())
            listPanel.Controls.Add(listBoxOptions)

            ' Buttons linksbündig mit Abstand
            Dim panelButtons As New System.Windows.Forms.FlowLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
        .Padding = New System.Windows.Forms.Padding(20, 10, 20, 20),
        .AutoSize = True,
        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
        .WrapContents = False
    }
            layout.Controls.Add(panelButtons, 0, 2)

            ' OK-Button
            Dim buttonOK As New System.Windows.Forms.Button() With {
        .Text = "OK",
        .DialogResult = System.Windows.Forms.DialogResult.OK,
        .Enabled = False,
        .AutoSize = True,
        .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4),
        .Margin = New System.Windows.Forms.Padding(0, 0, 20, 0)
    }
            AddHandler buttonOK.Click, Sub()
                                           selectedOption = CStr(listBoxOptions.SelectedItem)
                                       End Sub

            ' Cancel-Button (jetzt gleiche Margin‑Top wie OK)
            Dim buttonCancel As New System.Windows.Forms.Button() With {
        .Text = "Cancel",
        .DialogResult = System.Windows.Forms.DialogResult.Cancel,
        .AutoSize = True,
        .Padding = New System.Windows.Forms.Padding(8, 4, 8, 4),
        .Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
    }
            AddHandler buttonCancel.Click, Sub()
                                               selectedOption = "ESC"
                                               inputForm.Close()
                                           End Sub

            panelButtons.Controls.Add(buttonOK)
            panelButtons.Controls.Add(buttonCancel)

            ' Sicherstellen, dass beide Buttons dieselbe Höhe haben
            Dim btnHeight As Integer = Math.Max(buttonOK.Height, buttonCancel.Height)
            buttonOK.Height = btnHeight
            buttonCancel.Height = btnHeight

            ' Ereignisse für ListBox
            AddHandler listBoxOptions.SelectedIndexChanged, Sub()
                                                                buttonOK.Enabled = (listBoxOptions.SelectedItem IsNot Nothing)
                                                            End Sub
            AddHandler listBoxOptions.DoubleClick, Sub()
                                                       If listBoxOptions.SelectedItem IsNot Nothing Then
                                                           selectedOption = CStr(listBoxOptions.SelectedItem)
                                                           inputForm.DialogResult = System.Windows.Forms.DialogResult.OK
                                                           inputForm.Close()
                                                       End If
                                                   End Sub
            If listBoxOptions.Items.Count > 0 Then listBoxOptions.SelectedIndex = 0

            ' Tastenkürzel
            inputForm.AcceptButton = buttonOK
            inputForm.CancelButton = buttonCancel
            AddHandler inputForm.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                              If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                  selectedOption = "ESC"
                                                  inputForm.Close()
                                                  e.Handled = True
                                              End If
                                          End Sub

            ' Dialog anzeigen
            inputForm.TopMost = True
            inputForm.ShowDialog()
            Return selectedOption
        End Function


        Public Shared Function ShowCustomInputBox(
                                                    prompt As String,
                                                    title As String,
                                                    SimpleInput As Boolean,
                                                    Optional DefaultValue As String = "",
                                                    Optional CtrlP As String = "",
                                                    Optional OptionalButtons As System.Tuple(Of System.String, System.String, System.String)() = Nothing
                                                ) As String

            ' Screen working area (accounts for taskbar, etc.)
            Dim wa As System.Drawing.Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea

            ' Multi-line sizing rule: height = 1/6 of screen; width = 2 x
            Dim desiredInputHeight As Integer = 0
            Dim desiredInputWidth As Integer = 0
            If Not SimpleInput Then
                desiredInputHeight = Math.Max(150, CInt(wa.Height / 6.0))
                desiredInputWidth = CInt(desiredInputHeight * 3)
                desiredInputWidth = Math.Min(desiredInputWidth, wa.Width - 60) ' margin to fit in screen
            End If

            ' Create and configure the form (resizable in both modes)
            Dim inputForm As New Form() With {
                .Opacity = 0,
                .Text = title,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .StartPosition = FormStartPosition.Manual, ' we will center within working area after layout
                .MaximizeBox = False,
                .MinimizeBox = False,
                .ShowInTaskbar = False,
                .TopMost = True,
                .AutoScaleMode = AutoScaleMode.Font,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }

            ' Set the icon
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
            inputForm.Icon = Icon.FromHandle(bmp.GetHicon())

            ' Standard font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            inputForm.Font = standardFont

            ' Main layout for dynamic resizing
            Dim mainLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(20),
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            If SimpleInput Then
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' label
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' single-line textbox
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' buttons
            Else
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' label
                mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))  ' multi-line textbox grows/shrinks
                mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))         ' buttons
            End If

            ' Prompt label (wrap to initial target width; updated on resize)
            Dim initialLabelWrap As Integer = If(SimpleInput,
                                                 Math.Min(wa.Width - 120, 700),
                                                 Math.Max(400, desiredInputWidth))
            Dim promptLabel As New System.Windows.Forms.Label() With {
                .Text = prompt,
                .Font = standardFont,
                .AutoSize = True,
                .MaximumSize = New Size(initialLabelWrap, 0)
            }
            promptLabel.Dock = DockStyle.Top
            mainLayout.Controls.Add(promptLabel, 0, 0)

            ' Input TextBox
            Dim inputTextBox As New TextBox() With {
                .Font = standardFont,
                .Multiline = Not SimpleInput,
                .WordWrap = True,
                .ScrollBars = If(SimpleInput, ScrollBars.None, ScrollBars.Vertical),
                .Text = DefaultValue
            }
            If SimpleInput Then
                ' Single-line: compute height, stretch horizontally with the form
                inputTextBox.Height = TextRenderer.MeasureText("Wy", standardFont).Height + 6
                inputTextBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
                inputTextBox.Width = initialLabelWrap
            Else
                ' Multi-line: initial size by rule; allow growing with the form
                inputTextBox.MinimumSize = New Size(desiredInputWidth, desiredInputHeight)
                inputTextBox.Dock = DockStyle.Fill
            End If
            mainLayout.Controls.Add(inputTextBox, 0, 1)

            ' OK and Cancel buttons
            Dim okButton As New Button() With {.Text = "OK", .AutoSize = True, .Font = standardFont}
            Dim cancelButton As New Button() With {.Text = "Cancel", .AutoSize = True, .Font = standardFont}

            AddHandler okButton.Click, Sub()
                                           inputForm.DialogResult = DialogResult.OK
                                           inputForm.Close()
                                       End Sub
            AddHandler cancelButton.Click, Sub()
                                               inputForm.DialogResult = DialogResult.Cancel
                                               inputForm.Close()
                                           End Sub

            ' Bottom flow with wrapping to keep all buttons visible if space narrows
            Dim bottomFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Margin = New Padding(0, 20, 0, 0),
                .Dock = DockStyle.Top,
                .WrapContents = True
            }
            bottomFlow.Controls.Add(okButton)
            bottomFlow.Controls.Add(cancelButton)

            ' Optional extra buttons (max 5): label, tooltip, and prefix
            Dim selectedPrefix As String = Nothing
            If OptionalButtons IsNot Nothing AndAlso OptionalButtons.Length > 0 Then
                Dim tip As New System.Windows.Forms.ToolTip()
                Dim count As Integer = Math.Min(5, OptionalButtons.Length)
                For i As Integer = 0 To count - 1
                    Dim item = OptionalButtons(i)
                    Dim extraBtn As New System.Windows.Forms.Button() With {
                        .Text = item.Item1,
                        .AutoSize = True,
                        .Font = standardFont
                    }
                    tip.SetToolTip(extraBtn, item.Item2)
                    If i = 0 Then
                        extraBtn.Margin = New Padding(cancelButton.Margin.Left * 2, cancelButton.Margin.Top, cancelButton.Margin.Right, cancelButton.Margin.Bottom)
                    End If
                    AddHandler extraBtn.Click,
                        Sub()
                            selectedPrefix = item.Item3
                            inputForm.DialogResult = DialogResult.OK
                            inputForm.Close()
                        End Sub
                    bottomFlow.Controls.Add(extraBtn)
                Next
            End If

            mainLayout.Controls.Add(bottomFlow, 0, 2)
            inputForm.Controls.Add(mainLayout)

            ' Resize handler to keep label wrapping sensible when the user resizes the form
            AddHandler inputForm.Resize, Sub()
                                             ' Available width for content inside padding
                                             Dim available As Integer = Math.Max(300, mainLayout.ClientSize.Width)
                                             promptLabel.MaximumSize = New Size(available, 0)
                                             promptLabel.Invalidate()
                                         End Sub

            ' KeyDown handlers for Enter/Escape
            If SimpleInput Then
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            Else
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.Enter AndAlso e.Modifiers = Keys.Control Then
                                                         inputForm.DialogResult = DialogResult.OK
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     ElseIf e.KeyCode = Keys.Escape Then
                                                         inputForm.DialogResult = DialogResult.Cancel
                                                         inputForm.Close()
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            End If

            ' Ctrl+P insertion, if provided
            If Not String.IsNullOrEmpty(CtrlP) Then
                AddHandler inputTextBox.KeyDown, Sub(sender, e)
                                                     If e.KeyCode = Keys.P AndAlso e.Modifiers = Keys.Control Then
                                                         Dim selPos = inputTextBox.SelectionStart
                                                         inputTextBox.Text = inputTextBox.Text.Insert(selPos, CtrlP)
                                                         inputTextBox.SelectionStart = selPos + CtrlP.Length
                                                         e.SuppressKeyPress = True
                                                     End If
                                                 End Sub
            End If

            ' After AutoSize computed, clamp to screen, set MinimumSize (so buttons stay visible),
            ' disable AutoSize to allow user resizing, and center within the working area.
            AddHandler inputForm.Shown, Sub()
                                            ' Let AutoSize produce the preferred size first
                                            inputForm.PerformLayout()

                                            Dim maxW As Integer = wa.Width - 40
                                            Dim maxH As Integer = wa.Height - 40

                                            ' Compute space used by non-textbox rows and window chrome
                                            Dim chromeH As Integer = inputForm.Height - inputForm.ClientSize.Height
                                            Dim labelH As Integer = promptLabel.PreferredSize.Height
                                            Dim buttonsH As Integer = bottomFlow.PreferredSize.Height
                                            Dim paddingV As Integer = mainLayout.Padding.Vertical
                                            Dim gaps As Integer = bottomFlow.Margin.Top ' vertical gap above buttons

                                            Dim fixedRowsH As Integer = paddingV + labelH + gaps + buttonsH
                                            Dim maxClientH As Integer = maxH - chromeH

                                            If Not SimpleInput Then
                                                ' Allocate remaining height to the textbox, but stay within working area
                                                Dim textH As Integer = Math.Max(100, Math.Min(desiredInputHeight, maxClientH - fixedRowsH))

                                                ' Set client size so all rows are visible
                                                Dim newClientH As Integer = Math.Min(fixedRowsH + textH, maxClientH)

                                                ' Keep current width (already autosized) but clamp to screen
                                                Dim newClientW As Integer = Math.Min(inputForm.ClientSize.Width, maxW)

                                                inputForm.ClientSize = New Size(newClientW, newClientH)
                                            Else
                                                ' SimpleInput: just clamp to screen
                                                If inputForm.Width > maxW Then inputForm.Width = maxW
                                                If inputForm.Height > maxH Then inputForm.Height = maxH
                                            End If

                                            ' Minimum cannot be smaller than the current fully-visible content
                                            inputForm.MinimumSize = inputForm.Size

                                            ' Now allow resizing (keep MinimumSize so content/buttons never get clipped)
                                            inputForm.AutoSize = False

                                            ' Center within working area
                                            inputForm.Location = New System.Drawing.Point(
                                                wa.X + (wa.Width - inputForm.Width) \ 2,
                                                wa.Y + (wa.Height - inputForm.Height) \ 2
                                            )
                                        End Sub

            ' Ensure focus/topmost
            inputForm.TopMost = True
            inputForm.BringToFront()
            inputForm.Focus()

            ' Show the dialog, optionally owned by Outlook
            Dim Result As DialogResult
            If title.Contains("Browser") Then
                Dim outlookApp As Object = CreateObject("Outlook.Application")
                If outlookApp IsNot Nothing Then
                    Dim explorer As Object = outlookApp.GetType().InvokeMember("ActiveExplorer", BindingFlags.GetProperty, Nothing, outlookApp, Nothing)
                    If explorer IsNot Nothing Then
                        explorer.GetType().InvokeMember("WindowState", BindingFlags.SetProperty, Nothing, explorer, New Object() {1})
                        explorer.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, Nothing, explorer, Nothing)
                    End If
                End If
                inputForm.Opacity = 1
                Dim outlookHwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
                Result = inputForm.ShowDialog(New WindowWrapper(outlookHwnd))
            Else
                inputForm.Opacity = 1
                Result = inputForm.ShowDialog()
            End If

            ' Return the entered text or appropriate default
            If Result = DialogResult.OK Then
                Dim finalText As String = inputTextBox.Text
                If Not String.IsNullOrEmpty(selectedPrefix) AndAlso Not finalText.StartsWith(selectedPrefix, StringComparison.OrdinalIgnoreCase) Then
                    finalText = selectedPrefix & " " & finalText
                End If
                Debug.WriteLine("Final text: " & finalText)
                Return finalText
            Else
                Return If(Not SimpleInput, "ESC", "")
            End If
        End Function


        Public Shared Function ShowCustomYesNoBox(
                        ByVal bodyText As String,
                        ByVal button1Text As String,
                        ByVal button2Text As String,
                        Optional header As String = AN,
                        Optional autoCloseSeconds As Integer? = Nothing,
                        Optional Defaulttext As String = "",
                        Optional extraButtonText As String = Nothing,
                        Optional extraButtonAction As System.Action = Nothing,
                        Optional CloseAfterExtra As Boolean = False
                    ) As Integer

            ' Truncate if too long
            Dim isTruncated As Boolean = False
            If bodyText.Length > 10000 Then
                bodyText = bodyText.Substring(0, 10000)
                isTruncated = True
            End If

            ' Create and configure form
            Dim messageForm As New Form() With {
            .Opacity = 0,
            .Text = header,
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .StartPosition = FormStartPosition.CenterScreen,
            .MaximizeBox = False,
            .MinimizeBox = False,
            .ShowInTaskbar = False,
            .TopMost = True,
            .AutoScaleMode = AutoScaleMode.Font,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
        }

            ' Icon
            Dim bmpIcon As New Bitmap(My.Resources.Red_Ink_Logo)
            messageForm.Icon = Icon.FromHandle(bmpIcon.GetHicon())

            ' Font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            messageForm.Font = standardFont

            ' Layout containers
            Dim maxLabelWidth = 480
            Dim maxScreenHeight = Screen.PrimaryScreen.WorkingArea.Height - 100

            Dim mainFlow As New FlowLayoutPanel() With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(20),
            .MaximumSize = New Size(maxLabelWidth + 40, 0)
        }

            ' Body label
            Dim bodyLabel As New System.Windows.Forms.Label() With {
            .Text = bodyText,
            .Font = standardFont,
            .AutoSize = True,
            .MaximumSize = New Size(maxLabelWidth, maxScreenHeight \ 2)
        }
            mainFlow.Controls.Add(bodyLabel)

            ' “Text truncated” label, if needed
            If isTruncated Then
                Dim truncatedLabel As New System.Windows.Forms.Label() With {
                .Text = "(text has been truncated)",
                .Font = standardFont,
                .AutoSize = True
            }
                mainFlow.Controls.Add(truncatedLabel)
            End If

            ' Countdown label (for auto-close)
            Dim countdownLabel As New System.Windows.Forms.Label() With {
            .Font = standardFont,
            .AutoSize = True
        }

            ' Yes/No buttons
            Dim button1 As New Button() With {
            .Text = button1Text,
            .AutoSize = True,
            .Font = standardFont
        }
            Dim button2 As New Button() With {
            .Text = button2Text,
            .AutoSize = True,
            .Font = standardFont
        }

            ' Result variable
            Dim result As Integer = 0

            AddHandler button1.Click, Sub()
                                          result = 1
                                          messageForm.Close()
                                      End Sub
            AddHandler button2.Click, Sub()
                                          result = 2
                                          messageForm.Close()
                                      End Sub

            ' Bottom flow for buttons (+ countdown)
            Dim bottomFlow As New FlowLayoutPanel() With {
                        .FlowDirection = FlowDirection.LeftToRight,
                        .AutoSize = True,
                        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        .Margin = New Padding(0, 20, 0, 0)
                    }
            bottomFlow.Controls.Add(button1)
            bottomFlow.Controls.Add(button2)

            ' --- optional extra button, double distance from other buttons ---
            If (Not autoCloseSeconds.HasValue) AndAlso
       (Not String.IsNullOrEmpty(extraButtonText)) AndAlso
       (extraButtonAction IsNot Nothing) Then


                Dim extraButton As New System.Windows.Forms.Button() With {
                            .Text = extraButtonText,
                            .AutoSize = True,
                            .Font = standardFont,
                                         .Margin = New System.Windows.Forms.Padding(10, button1.Margin.Top, 0, button1.Margin.Bottom)
                        }

                AddHandler extraButton.Click,
            Sub()
                Try
                    extraButtonAction.Invoke()
                Catch ex As System.Exception
                    ' Optional: log or handle exception
                End Try
                If CloseAfterExtra Then messageForm.Close()
            End Sub

                bottomFlow.Controls.Add(extraButton)
            End If


            If autoCloseSeconds.HasValue Then
                bottomFlow.Controls.Add(countdownLabel)
            End If
            mainFlow.Controls.Add(bottomFlow)

            messageForm.Controls.Add(mainFlow)


            ' Auto-close timer
            If autoCloseSeconds.HasValue Then
                Dim remaining = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Dim t As New System.Windows.Forms.Timer() With {.Interval = 1000}
                AddHandler t.Tick, Sub()
                                       remaining -= 1
                                       If remaining > 0 Then
                                           countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                                       Else
                                           t.Stop()
                                           result = 3
                                           messageForm.Close()
                                       End If
                                   End Sub
                t.Start()
            End If

            ' Show and return
            messageForm.TopMost = True
            messageForm.Opacity = 1
            messageForm.ShowDialog()
            messageForm.Activate()
            Return result
        End Function


        Public Shared Sub ShowCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional autoCloseSeconds As System.Nullable(Of Integer) = Nothing,
    Optional Defaulttext As String = " - execution continues meanwhile",
    Optional SeparateThread As Boolean = False,
    Optional extraButtonText As String = Nothing,
    Optional extraButtonAction As System.Action = Nothing,
    Optional CloseAfterExtra As Boolean = False
)
            If System.String.IsNullOrWhiteSpace(header) Then header = AN
            Dim isTruncated As System.Boolean = False
            If bodyText IsNot Nothing AndAlso bodyText.Length > 10000 Then
                bodyText = bodyText.Substring(0, 10000) & "(...)"
                isTruncated = True
            End If

            Dim messageForm As New System.Windows.Forms.Form() With {
        .Opacity = 0,
        .Text = header,
        .FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
        .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
        .MaximizeBox = False,
        .MinimizeBox = False,
        .ShowInTaskbar = False,
        .TopMost = True,
        .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font,
        .AutoSize = False
    }

            Dim bmpIcon As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            messageForm.Icon = System.Drawing.Icon.FromHandle(bmpIcon.GetHicon())

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            messageForm.Font = standardFont

            Dim wa As System.Drawing.Rectangle = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
            Dim paddingAll As System.Int32 = 20
            Dim gapAboveButtons As System.Int32 = 10 ' keep existing gap logic
            Dim spacerExtra As System.Int32 = 20    ' NEW: extra space between text and buttons
            Dim minContentWidth As System.Int32 = 360
            Dim startContentWidth As System.Int32 = 500
            Dim maxWindowWidth As System.Int32 = CInt(System.Math.Floor(wa.Width * 0.5))
            Dim maxWindowHeight As System.Int32 = CInt(System.Math.Floor(wa.Height * 0.9))

            Dim okButton As New System.Windows.Forms.Button() With {.Text = "OK", .AutoSize = True, .Font = standardFont, .Margin = New System.Windows.Forms.Padding(0)}
            Dim countdownLabel As New System.Windows.Forms.Label() With {.Font = standardFont, .AutoSize = True, .Margin = New System.Windows.Forms.Padding(8, 0, 0, 0)}
            Dim userClicked As System.Boolean = False
            AddHandler okButton.Click, Sub()
                                           userClicked = True
                                           messageForm.Close()
                                       End Sub

            Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
        .Margin = New System.Windows.Forms.Padding(0)
    }
            bottomFlow.Controls.Add(okButton)

            ' optional extra button
            If (Not autoCloseSeconds.HasValue) AndAlso
       (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso
       (extraButtonAction IsNot Nothing) Then


                Dim extraButton As New System.Windows.Forms.Button() With {
                            .Text = extraButtonText,
                            .AutoSize = True,
                            .Font = standardFont,
                            .Margin = New System.Windows.Forms.Padding(8, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                        }


                AddHandler extraButton.Click,
            Sub()
                Try
                    extraButtonAction.Invoke()
                Catch ex As System.Exception
                End Try
                If CloseAfterExtra Then messageForm.Close()
            End Sub
                bottomFlow.Controls.Add(extraButton)
            End If
            If autoCloseSeconds.HasValue Then bottomFlow.Controls.Add(countdownLabel)

            bottomFlow.PerformLayout()
            Dim bottomSize As System.Drawing.Size = bottomFlow.PreferredSize
            Dim reservedBottomHeight As System.Int32 = bottomSize.Height + gapAboveButtons

            Dim bodyLabel As New System.Windows.Forms.Label() With {
        .Text = If(bodyText, System.String.Empty),
        .Font = standardFont,
        .AutoSize = True,
        .Margin = New System.Windows.Forms.Padding(0)
    }

            Dim GetLabelPreferred As System.Func(Of System.Int32, System.Drawing.Size) =
        Function(w As System.Int32) As System.Drawing.Size
            bodyLabel.MaximumSize = New System.Drawing.Size(System.Math.Max(1, w), 0)
            Return bodyLabel.GetPreferredSize(New System.Drawing.Size(System.Math.Max(1, w), 0))
        End Function

            Dim contentWidth As System.Int32 = System.Math.Max(minContentWidth, System.Math.Min(startContentWidth, maxWindowWidth - 2 * paddingAll))
            Dim pref As System.Drawing.Size = GetLabelPreferred(contentWidth)
            Dim maxBodyHeightNoScroll As System.Int32 = System.Math.Max(100, maxWindowHeight - reservedBottomHeight - spacerExtra - 2 * paddingAll) ' include spacer in budget

            While (pref.Height > maxBodyHeightNoScroll) AndAlso ((contentWidth + 2 * paddingAll) < maxWindowWidth)
                Dim stepW As System.Int32 = System.Math.Max(24, (maxWindowWidth - 2 * paddingAll - contentWidth) \ 3)
                contentWidth = System.Math.Min(maxWindowWidth - 2 * paddingAll, contentWidth + stepW)
                pref = GetLabelPreferred(contentWidth)
            End While

            Dim needScroll As System.Boolean = pref.Height > maxBodyHeightNoScroll
            Dim usableTextWidth As System.Int32 = contentWidth
            If needScroll Then
                usableTextWidth = System.Math.Max(100, contentWidth - System.Windows.Forms.SystemInformation.VerticalScrollBarWidth)
                pref = GetLabelPreferred(usableTextWidth)
            End If

            Dim bodyPanelHeight As System.Int32 = If(needScroll, maxBodyHeightNoScroll, pref.Height)

            Dim bodyScrollPanel As New System.Windows.Forms.Panel() With {
        .AutoScroll = False,
        .AutoSize = False,
        .Size = New System.Drawing.Size(contentWidth, bodyPanelHeight),
        .Margin = New System.Windows.Forms.Padding(0),
        .Padding = New System.Windows.Forms.Padding(0)
    }
            bodyScrollPanel.HorizontalScroll.Enabled = False
            bodyScrollPanel.HorizontalScroll.Visible = False

            bodyLabel.MaximumSize = New System.Drawing.Size(usableTextWidth, 0)
            bodyScrollPanel.Controls.Add(bodyLabel)
            bodyLabel.Location = New System.Drawing.Point(0, 0)

            If needScroll Then
                bodyScrollPanel.AutoScroll = True
                bodyScrollPanel.AutoScrollMinSize = New System.Drawing.Size(usableTextWidth, pref.Height)
            End If

            ' --- MAIN TABLE: now 3 rows: [text][SPACER][buttons] ---
            Dim table As New System.Windows.Forms.TableLayoutPanel() With {
        .Dock = System.Windows.Forms.DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3, ' NEW
        .Padding = New System.Windows.Forms.Padding(paddingAll),
        .AutoSize = False,
        .Margin = New System.Windows.Forms.Padding(0)
    }
            table.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0F))
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, bodyPanelHeight))  ' text
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, spacerExtra))       ' NEW spacer
            table.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))                    ' buttons

            table.Controls.Add(bodyScrollPanel, 0, 0)

            ' NEW spacer: exact +20 px above the buttons
            Dim spacer As New System.Windows.Forms.Panel() With {.Height = spacerExtra, .Width = 1, .Margin = New System.Windows.Forms.Padding(0)}
            table.Controls.Add(spacer, 0, 1)

            Dim bottomHost As New System.Windows.Forms.Panel() With {.AutoSize = True, .Margin = New System.Windows.Forms.Padding(0)}
            bottomHost.Padding = New System.Windows.Forms.Padding(0, gapAboveButtons, 0, 0)
            bottomHost.Controls.Add(bottomFlow)
            table.Controls.Add(bottomHost, 0, 2)

            messageForm.Controls.Clear()
            messageForm.Controls.Add(table)

            ' Final size: include spacerExtra
            Dim clientW As System.Int32 = contentWidth + 2 * paddingAll
            Dim clientH As System.Int32 = bodyPanelHeight + spacerExtra + reservedBottomHeight + 2 * paddingAll
            clientW = System.Math.Min(clientW, maxWindowWidth)
            clientH = System.Math.Min(clientH, maxWindowHeight)
            messageForm.ClientSize = New System.Drawing.Size(clientW, clientH)

            If autoCloseSeconds.HasValue Then
                Dim remaining As System.Int32 = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Dim t As New System.Windows.Forms.Timer() With {.Interval = 1000}
                AddHandler t.Tick,
            Sub()
                remaining -= 1
                If remaining > 0 Then
                    countdownLabel.Text = $"(closes in {remaining} seconds{Defaulttext})"
                Else
                    t.Stop()
                    If Not userClicked Then
                        messageForm.Close()
                    End If
                End If
            End Sub
                t.Start()

                messageForm.Opacity = 1
                If SeparateThread Then
                    messageForm.BringToFront()
                    messageForm.Focus()
                    messageForm.Activate()

                    AddHandler messageForm.Shown,
                            Sub(sender, e)
                                messageForm.TopMost = False  ' Reset first
                                messageForm.TopMost = True   ' Then set again
                                messageForm.Activate()
                                messageForm.BringToFront()
                            End Sub
                    messageForm.ShowDialog()
                Else
                    messageForm.Show()
                    System.Windows.Forms.Application.DoEvents()
                End If
            Else

                messageForm.BringToFront()
                messageForm.Focus()
                messageForm.Activate()

                AddHandler messageForm.Shown,
                        Sub(sender, e)
                            messageForm.TopMost = False  ' Reset first
                            messageForm.TopMost = True   ' Then set again
                            messageForm.Activate()
                            messageForm.BringToFront()
                        End Sub

                messageForm.Opacity = 1
                messageForm.ShowDialog()
            End If
        End Sub





        Public Shared Sub ShowRTFCustomMessageBox(ByVal bodyText As String, Optional header As String = AN, Optional autoCloseSeconds As Integer? = Nothing, Optional Defaulttext As String = " - execution continues meanwhile")

            Dim RTFMessageForm As New System.Windows.Forms.Form()
            Dim bodyLabel As New System.Windows.Forms.RichTextBox()
            Dim okButton As New System.Windows.Forms.Button()
            Dim countdownLabel As New System.Windows.Forms.Label()
            Dim Truncated As Boolean = False

            If String.IsNullOrWhiteSpace(header) Then header = AN

            ' Form attributes
            RTFMessageForm.Opacity = 0
            RTFMessageForm.Text = header
            RTFMessageForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            RTFMessageForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            RTFMessageForm.MaximizeBox = True
            RTFMessageForm.MinimizeBox = True
            RTFMessageForm.ShowInTaskbar = False
            RTFMessageForm.TopMost = True
            RTFMessageForm.KeyPreview = True

            ' Autoscale for fonts & DPI
            RTFMessageForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            RTFMessageForm.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)

            RTFMessageForm.MinimumSize = New System.Drawing.Size(650, 335)

            ' Icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            RTFMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Standard font
            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            ' Body RTF box
            ' Body RTF box
            bodyLabel.Font = standardFont
            bodyLabel.ReadOnly = True
            bodyLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
            bodyLabel.BackColor = RTFMessageForm.BackColor
            bodyLabel.TabStop = False
            bodyLabel.Rtf = bodyText
            bodyLabel.Location = New System.Drawing.Point(20, 20)
            bodyLabel.Width = 600
            bodyLabel.Height = 200
            ' Anchor to all sides so it resizes with the form
            bodyLabel.Anchor = System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right _
                     Or System.Windows.Forms.AnchorStyles.Bottom
            RTFMessageForm.Controls.Add(bodyLabel)


            ' OK button & countdown label setup
            okButton.Font = standardFont
            okButton.Text = "OK"
            okButton.AutoSize = True

            countdownLabel.Font = standardFont
            countdownLabel.AutoSize = True

            ' Bottom panel to hold button + countdown, docked so it moves with resizing
            Dim bottomPanel As New System.Windows.Forms.Panel()
            bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom
            bottomPanel.Padding = New System.Windows.Forms.Padding(20)  ' 20px padding on all sides
            bottomPanel.Height = okButton.PreferredSize.Height + bottomPanel.Padding.Top + bottomPanel.Padding.Bottom
            RTFMessageForm.Controls.Add(bottomPanel)

            ' Add controls into panel
            bottomPanel.Controls.Add(okButton)
            bottomPanel.Controls.Add(countdownLabel)
            okButton.Location = New System.Drawing.Point(bottomPanel.Padding.Left, bottomPanel.Padding.Top)
            countdownLabel.Location = New System.Drawing.Point(okButton.Right + 10, bottomPanel.Padding.Top)

            ' Ensure bodyLabel resizes when form is resized
            AddHandler RTFMessageForm.Resize, Sub(sender As Object, e As EventArgs)
                                                  Dim availableWidth As Integer = RTFMessageForm.ClientSize.Width - bodyLabel.Left - 20
                                                  Dim availableHeight As Integer = RTFMessageForm.ClientSize.Height - bottomPanel.Height - bodyLabel.Top - 20
                                                  bodyLabel.Size = New System.Drawing.Size(availableWidth, availableHeight)
                                              End Sub

            ' Handlers
            Dim userClicked As Boolean = False
            AddHandler okButton.Click, Sub(sender As Object, e As EventArgs)
                                           userClicked = True
                                           RTFMessageForm.Close()
                                           RTFMessageForm = Nothing
                                       End Sub
            AddHandler RTFMessageForm.KeyDown, Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
                                                   If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                       userClicked = True
                                                       RTFMessageForm.Close()
                                                       RTFMessageForm = Nothing
                                                       e.SuppressKeyPress = True
                                                   End If
                                               End Sub
            AddHandler RTFMessageForm.Shown, Sub(sender As Object, e As EventArgs)
                                                 ' Trigger initial resize layout
                                                 RTFMessageForm.PerformLayout()
                                                 RTFMessageForm.Activate()
                                             End Sub

            ' Initial form sizing: ensure 20px padding around button and RTF label sizing
            Dim formWidth As Integer = Math.Max(RTFMessageForm.MinimumSize.Width, bodyLabel.Width + 40)
            Dim formHeight As Integer = Math.Max(RTFMessageForm.MinimumSize.Height,
                                         bodyLabel.Bottom + 20 + bottomPanel.Height)
            RTFMessageForm.ClientSize = New System.Drawing.Size(formWidth, formHeight)

            ' Auto-close timer
            If autoCloseSeconds.HasValue AndAlso autoCloseSeconds > 0 Then
                Dim remainingTime As Integer = autoCloseSeconds.Value
                countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"

                Dim timer As New System.Windows.Forms.Timer()
                timer.Interval = 1000
                AddHandler timer.Tick, Sub(sender As Object, e As EventArgs)
                                           remainingTime -= 1
                                           If remainingTime > 0 Then
                                               countdownLabel.Text = $"(closes in {remainingTime} seconds{Defaulttext})"
                                           Else
                                               timer.Stop()
                                               If Not userClicked Then
                                                   RTFMessageForm.Close()
                                               End If
                                           End If
                                       End Sub
                timer.Start()


                RTFMessageForm.BringToFront()
                RTFMessageForm.Focus()
                RTFMessageForm.Activate()

                AddHandler RTFMessageForm.Shown,
                                        Sub(sender, e)
                                            RTFMessageForm.TopMost = False  ' Reset first
                                            RTFMessageForm.TopMost = True   ' Then set again
                                            RTFMessageForm.Activate()
                                            RTFMessageForm.BringToFront()
                                        End Sub

                RTFMessageForm.Opacity = 1
                RTFMessageForm.Show()
                RTFMessageForm.BringToFront()
                RTFMessageForm.Activate()
                System.Windows.Forms.Application.DoEvents()
            Else


                RTFMessageForm.BringToFront()
                RTFMessageForm.Focus()
                RTFMessageForm.Activate()

                AddHandler RTFMessageForm.Shown,
                                        Sub(sender, e)
                                            RTFMessageForm.TopMost = False  ' Reset first
                                            RTFMessageForm.TopMost = True   ' Then set again
                                            RTFMessageForm.Activate()
                                            RTFMessageForm.BringToFront()
                                        End Sub

                RTFMessageForm.Opacity = 1
                RTFMessageForm.ShowDialog()
            End If

        End Sub


        Public Shared Sub ShowHTMLCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional Defaulttext As String = " - execution continues meanwhile",
    Optional extraButtonText As String = Nothing,
    Optional extraButtonAction As System.Action = Nothing,
    Optional CloseAfterExtra As Boolean = False
)
            Dim t As New Thread(Sub()
                                    ' Create and configure form
                                    Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                                .Opacity = 0,
                                .Text = header,
                                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                                .MaximizeBox = True,
                                .MinimizeBox = True,
                                .ShowInTaskbar = True,
                                .TopMost = False,
                                .KeyPreview = True,
                                .MinimumSize = New System.Drawing.Size(800, 500),
                                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                            }

                                    ' Header fallback
                                    If String.IsNullOrWhiteSpace(header) Then
                                        HTMLMessageForm.Text = AN
                                    End If

                                    ' Set the icon
                                    Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                                    HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                    ' Standard font
                                    Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                                    HTMLMessageForm.Font = standardFont

                                    ' WebBrowser mit 10px Margin
                                    Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                                .AllowNavigation = False,
                                .WebBrowserShortcutsEnabled = False,
                                .ScrollBarsEnabled = True,
                                .ScriptErrorsSuppressed = True,
                                .DocumentText = bodyText,
                                .Dock = System.Windows.Forms.DockStyle.Fill,
                                .BackColor = HTMLMessageForm.BackColor,
                                .Margin = New System.Windows.Forms.Padding(20)
                            }
                                    AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                                                  If htmlBrowser.Document?.Body IsNot Nothing Then
                                                                                      ' Body-Style mit 10px Margin innen
                                                                                      htmlBrowser.Document.Body.Style =
                                                                                  $"background-color: rgb({HTMLMessageForm.BackColor.R}, {HTMLMessageForm.BackColor.G}, {HTMLMessageForm.BackColor.B}); " &
                                                                                  "font-family: 'Segoe UI'; font-size: 9pt; margin: 20px;"
                                                                                  End If
                                                                              End Sub

                                    ' OK button
                                    Dim okButton As New System.Windows.Forms.Button() With {
                                .Text = "OK",
                                .AutoSize = True,
                                .Font = standardFont,
                                .Margin = New System.Windows.Forms.Padding(0)
                            }
                                    AddHandler okButton.Click, Sub()
                                                                   HTMLMessageForm.Close()
                                                               End Sub

                                    ' Form‐level Escape
                                    AddHandler HTMLMessageForm.KeyDown, Sub(sender2, e2)
                                                                            If e2.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                                                HTMLMessageForm.Close()
                                                                                e2.SuppressKeyPress = True
                                                                            End If
                                                                        End Sub

                                    ' Activate on shown
                                    AddHandler HTMLMessageForm.Shown, Sub(sender2, e2)
                                                                          HTMLMessageForm.Activate()
                                                                      End Sub

                                    ' Bottom flow panel
                                    Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                .Dock = System.Windows.Forms.DockStyle.Bottom,
                                .AutoSize = True,
                                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                .Padding = New System.Windows.Forms.Padding(20)
                            }
                                    bottomFlow.Controls.Add(okButton)

                                    ' Optional extra button
                                    If (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                                        Dim extraButton As New System.Windows.Forms.Button() With {
                                            .Text = extraButtonText,
                                            .AutoSize = True,
                                            .Font = standardFont,
                                            .Margin = New System.Windows.Forms.Padding(10, okButton.Margin.Top, 0, okButton.Margin.Bottom)
                                        }

                                        AddHandler extraButton.Click,
                                            Sub()
                                                Try
                                                    ' Execute the action - recursive ShowHTMLCustomMessageBox calls
                                                    ' will spawn their own STA threads, so this is safe
                                                    extraButtonAction.Invoke()
                                                Catch ex As System.Exception
                                                    ' Swallow to keep dialog functional
                                                End Try
                                                If CloseAfterExtra Then HTMLMessageForm.Close()
                                            End Sub

                                        bottomFlow.Controls.Add(extraButton)
                                    End If

                                    ' Compose form
                                    HTMLMessageForm.Controls.Add(htmlBrowser)
                                    HTMLMessageForm.Controls.Add(bottomFlow)

                                    HTMLMessageForm.BringToFront()
                                    HTMLMessageForm.Focus()
                                    HTMLMessageForm.Activate()

                                    AddHandler HTMLMessageForm.Shown,
                                        Sub(sender, e)
                                            HTMLMessageForm.TopMost = False  ' Reset first
                                            HTMLMessageForm.TopMost = True   ' Then set again
                                            HTMLMessageForm.Activate()
                                            HTMLMessageForm.BringToFront()
                                        End Sub

                                    HTMLMessageForm.Opacity = 1

                                    ' Show dialog

                                    HTMLMessageForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
        End Sub


        Public Shared Sub oldShowHTMLCustomMessageBox(
    ByVal bodyText As String,
    Optional header As String = AN,
    Optional Defaulttext As String = " - execution continues meanwhile"
)
            Dim t As New Thread(Sub()
                                    ' Create and configure form
                                    Dim HTMLMessageForm As New System.Windows.Forms.Form() With {
                                .Opacity = 0,
                                .Text = header,
                                .FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable,
                                .StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                                .MaximizeBox = True,
                                .MinimizeBox = True,
                                .ShowInTaskbar = True,
                                .TopMost = False,
                                .KeyPreview = True,
                                .MinimumSize = New System.Drawing.Size(800, 500),
                                .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                            }

                                    ' Header fallback
                                    If String.IsNullOrWhiteSpace(header) Then
                                        HTMLMessageForm.Text = AN
                                    End If

                                    ' Set the icon
                                    Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                                    HTMLMessageForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                                    ' Standard font
                                    Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                                    HTMLMessageForm.Font = standardFont

                                    ' WebBrowser mit 10px Margin
                                    Dim htmlBrowser As New System.Windows.Forms.WebBrowser() With {
                                .AllowNavigation = False,
                                .WebBrowserShortcutsEnabled = False,
                                .ScrollBarsEnabled = True,
                                .ScriptErrorsSuppressed = True,
                                .DocumentText = bodyText,
                                .Dock = System.Windows.Forms.DockStyle.Fill,
                                .BackColor = HTMLMessageForm.BackColor,
                                .Margin = New System.Windows.Forms.Padding(20)
                            }
                                    AddHandler htmlBrowser.DocumentCompleted, Sub(sender2, e2)
                                                                                  If htmlBrowser.Document?.Body IsNot Nothing Then
                                                                                      ' Body-Style mit 10px Margin innen
                                                                                      htmlBrowser.Document.Body.Style =
                                                                                  $"background-color: rgb({HTMLMessageForm.BackColor.R}, {HTMLMessageForm.BackColor.G}, {HTMLMessageForm.BackColor.B}); " &
                                                                                  "font-family: 'Segoe UI'; font-size: 9pt; margin: 20px;"
                                                                                  End If
                                                                              End Sub

                                    ' OK button
                                    Dim okButton As New System.Windows.Forms.Button() With {
                                .Text = "OK",
                                .AutoSize = True,
                                .Font = standardFont,
                                .Margin = New System.Windows.Forms.Padding(0) ' kein zusätzlicher Abstand hier
                            }
                                    AddHandler okButton.Click, Sub()
                                                                   HTMLMessageForm.Close()
                                                               End Sub

                                    ' Form‐level Escape
                                    AddHandler HTMLMessageForm.KeyDown, Sub(sender2, e2)
                                                                            If e2.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                                                HTMLMessageForm.Close()
                                                                                e2.SuppressKeyPress = True
                                                                            End If
                                                                        End Sub

                                    ' Activate on shown
                                    AddHandler HTMLMessageForm.Shown, Sub(sender2, e2)
                                                                          HTMLMessageForm.Activate()
                                                                      End Sub

                                    ' Bottom flow panel
                                    Dim bottomFlow As New System.Windows.Forms.FlowLayoutPanel() With {
                                .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                .Dock = System.Windows.Forms.DockStyle.Bottom,
                                .AutoSize = True,
                                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                                .Padding = New System.Windows.Forms.Padding(20)
                            }
                                    bottomFlow.Controls.Add(okButton)

                                    ' Compose form
                                    HTMLMessageForm.Controls.Add(htmlBrowser)
                                    HTMLMessageForm.Controls.Add(bottomFlow)

                                    HTMLMessageForm.BringToFront()
                                    HTMLMessageForm.Focus()
                                    HTMLMessageForm.Activate()

                                    AddHandler HTMLMessageForm.Shown,
                                        Sub(sender, e)
                                            HTMLMessageForm.TopMost = False  ' Reset first
                                            HTMLMessageForm.TopMost = True   ' Then set again
                                            HTMLMessageForm.Activate()
                                            HTMLMessageForm.BringToFront()
                                        End Sub

                                    HTMLMessageForm.Opacity = 1

                                    ' Show dialog

                                    HTMLMessageForm.ShowDialog()
                                End Sub)
            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
        End Sub


        Public Shared Function ShowCustomVariableInputForm(
                                            ByVal prompt As String,
                                            ByVal header As String,
                                            ByRef params() As InputParameter,
                                            Optional extraButtonText As System.String = Nothing,
                                            Optional extraButtonAction As System.Action = Nothing,
                                            Optional CloseAfterExtra As System.Boolean = False
                                        ) As Boolean
            If String.IsNullOrWhiteSpace(header) Then header = String.Empty

            Dim inputForm As New Form() With {
                .Text = header,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .StartPosition = FormStartPosition.CenterScreen,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .AutoScaleMode = AutoScaleMode.Font,
                .AutoScaleDimensions = New SizeF(6.0F, 13.0F),
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .KeyPreview = True ' allow form to see Ctrl+Enter before controls
            }

            ' Set icon
            Dim bmpIcon As New Bitmap(My.Resources.Red_Ink_Logo)
            inputForm.Icon = Icon.FromHandle(bmpIcon.GetHicon())

            ' Layout
            Dim mainLayout As New TableLayoutPanel() With {
                .ColumnCount = 2,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12),
                .GrowStyle = TableLayoutPanelGrowStyle.AddRows
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            ' Prompt label
            Dim promptLabel As New System.Windows.Forms.Label() With {
                .Text = prompt,
                .AutoSize = True,
                .MaximumSize = New Size(600, 0),
                .Margin = New Padding(0, 0, 0, 12)
            }
            mainLayout.Controls.Add(promptLabel, 0, 0)
            mainLayout.SetColumnSpan(promptLabel, 2)

            ' Component container + tooltip
            Dim components As New System.ComponentModel.Container()
            Dim toolTip As New System.Windows.Forms.ToolTip(components) With {
                .ShowAlways = True
            }

            For i As Integer = 0 To params.Length - 1
                Dim param = params(i)
                Dim rawValue As Object = param.Value

                Dim lbl As New System.Windows.Forms.Label() With {
                    .Text = param.Name & ":",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 0, 8, 8)
                }
                mainLayout.Controls.Add(lbl, 0, i + 1)

                Dim ctrl As Control

                ' RULES:
                ' 1. If value Is Nothing -> show DISABLED CheckBox (unchecked).
                ' 2. If value Is Boolean -> show enabled CheckBox with that state.
                ' 3. Else if options exist -> ComboBox.
                ' 4. Else -> TextBox.
                Dim isNothing As Boolean = (rawValue Is Nothing)
                Dim isBool As Boolean = TypeOf rawValue Is Boolean

                Dim sentinelDisabled As String = "<<disabled>>"
                Dim disableForSentinel As Boolean =
                    (TypeOf rawValue Is String AndAlso
                     String.Equals(CStr(rawValue), sentinelDisabled, System.StringComparison.Ordinal))

                If disableForSentinel Then rawValue = ""

                If isNothing OrElse isBool Then
                    Dim initial As Boolean = If(isBool, CBool(rawValue), False)
                    Dim chk As New System.Windows.Forms.CheckBox() With {
                        .Checked = initial,
                        .AutoSize = True,
                        .Anchor = AnchorStyles.Left,
                        .Margin = New Padding(0, 0, 0, 8),
                        .Enabled = Not isNothing
                    }
                    If isNothing Then
                        chk.BackColor = SystemColors.Control
                        toolTip.SetToolTip(chk, "Not available")
                    End If
                    ctrl = chk

                ElseIf param.Options IsNot Nothing AndAlso param.Options.Count > 0 AndAlso TypeOf rawValue Is String Then
                    Dim cb As New System.Windows.Forms.ComboBox() With {
                        .DropDownStyle = ComboBoxStyle.DropDownList,
                        .MaxDropDownItems = 5,
                        .IntegralHeight = False,
                        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                        .Margin = New Padding(0, 0, 0, 12),
                        .MinimumSize = New Size(400, 0)
                    }
                    cb.Items.AddRange(param.Options.ToArray())
                    If param.Options.Contains(CStr(rawValue)) Then cb.SelectedItem = rawValue

                    ' Adjust dropdown width
                    Dim maxItemWidth As Integer = 0
                    For Each it In cb.Items
                        Dim w = TextRenderer.MeasureText(CStr(it), cb.Font).Width
                        If w > maxItemWidth Then maxItemWidth = w
                    Next
                    Dim needsScroll = cb.Items.Count > cb.MaxDropDownItems
                    Dim scrollW = If(needsScroll, SystemInformation.VerticalScrollBarWidth, 0)
                    cb.DropDownWidth = Math.Max(cb.DropDownWidth, maxItemWidth + scrollW + 16)

                    ' Tooltip if truncated
                    Dim updateToolTip As EventHandler =
                        Sub(sender As Object, eArgs As EventArgs)
                            Dim combo = DirectCast(sender, ComboBox)
                            Dim t = combo.Text
                            Dim tw = TextRenderer.MeasureText(t, combo.Font).Width
                            Dim usable = Math.Max(0, combo.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6)
                            If tw > usable Then
                                toolTip.SetToolTip(combo, t)
                            Else
                                toolTip.SetToolTip(combo, Nothing)
                            End If
                        End Sub
                    AddHandler cb.SelectedIndexChanged, updateToolTip
                    AddHandler cb.TextChanged, updateToolTip
                    AddHandler cb.Resize, updateToolTip
                    AddHandler cb.MouseEnter, updateToolTip
                    updateToolTip(cb, EventArgs.Empty)

                    ctrl = cb

                Else
                    Dim txt As New TextBox() With {
                        .Text = rawValue.ToString(),
                        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                        .Margin = New Padding(0, 0, 0, 8)
                    }
                    If TypeOf rawValue Is String Then
                        txt.MinimumSize = New Size(400, 0)
                    Else
                        txt.MinimumSize = New Size(50, 0)
                    End If
                    ctrl = txt
                End If

                If disableForSentinel Then
                    ctrl.Enabled = False
                    toolTip.SetToolTip(ctrl, "Not available")
                End If

                param.InputControl = ctrl
                mainLayout.Controls.Add(ctrl, 1, i + 1)
            Next

            ' Buttons
            Dim buttonFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 8, 12, 12)
            }
            Dim btnOK As New Button() With {.Text = "OK", .AutoSize = True, .DialogResult = DialogResult.OK}
            Dim btnCancel As New Button() With {.Text = "Cancel", .AutoSize = True, .DialogResult = DialogResult.Cancel}

            ' Add in this order so visual order is [OK][Cancel] with RightToLeft
            buttonFlow.Controls.Add(btnCancel)
            buttonFlow.Controls.Add(btnOK)

            ' Ensure Tab order prefers OK when tabbing out of the last field
            btnOK.TabIndex = 0
            btnCancel.TabIndex = 2 ' will move to 1 if no extra button is added

            ' Optional extra button: same behavior as ShowCustomMessageBox
            Dim extraButton As System.Windows.Forms.Button = Nothing
            If (Not System.String.IsNullOrEmpty(extraButtonText)) AndAlso (extraButtonAction IsNot Nothing) Then
                extraButton = New System.Windows.Forms.Button() With {
                    .Text = extraButtonText,
                    .AutoSize = True,
                    .Margin = New System.Windows.Forms.Padding(8, btnOK.Margin.Top, 0, btnOK.Margin.Bottom)
                }
                AddHandler extraButton.Click,
                    Sub()
                        Try
                            extraButtonAction.Invoke()
                        Catch ex As System.Exception
                            ' swallow to keep dialog functional; mirror ShowCustomMessageBox behavior
                        End Try
                        If CloseAfterExtra Then
                            inputForm.DialogResult = DialogResult.Cancel ' do not commit changes implicitly
                            inputForm.Close()
                        End If
                    End Sub

                ' Place the extra button to the left of OK (RightToLeft flow)
                buttonFlow.Controls.Add(extraButton)

                ' Tab order: OK first, then extra, then Cancel
                extraButton.TabIndex = 1
            Else
                ' No extra button: let Cancel be second
                btnCancel.TabIndex = 1
            End If

            inputForm.Controls.Add(mainLayout)
            inputForm.Controls.Add(buttonFlow)

            ' Ctrl+Enter should trigger OK anywhere on the form
            AddHandler inputForm.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.KeyCode = Keys.Enter AndAlso e.Control Then
                        btnOK.PerformClick()
                        e.SuppressKeyPress = True
                        e.Handled = True
                    End If
                End Sub

            Dim result = inputForm.ShowDialog()

            If result = DialogResult.OK Then
                For Each param In params
                    ' Skip disabled controls: keep existing Value
                    If param.InputControl IsNot Nothing AndAlso Not param.InputControl.Enabled Then
                        Continue For
                    End If
                    Try
                        If TypeOf param.InputControl Is System.Windows.Forms.ComboBox Then
                            Dim cb = DirectCast(param.InputControl, System.Windows.Forms.ComboBox)
                            param.Value = If(cb.SelectedItem IsNot Nothing, cb.SelectedItem.ToString(), cb.Text)
                        ElseIf TypeOf param.Value Is Boolean Then
                            param.Value = CType(param.InputControl, System.Windows.Forms.CheckBox).Checked
                        ElseIf TypeOf param.Value Is Integer Then
                            Dim valI As Integer
                            If Integer.TryParse(CType(param.InputControl, TextBox).Text, valI) Then
                                param.Value = valI
                            Else
                                Throw New Exception($"Invalid value for {param.Name}.")
                            End If
                        ElseIf TypeOf param.Value Is Double Then
                            Dim valD As Double
                            Dim inputText As String = CType(param.InputControl, TextBox).Text.Trim()

                            ' Normalize: replace comma with dot, then parse with invariant culture
                            Dim normalizedInput As String = inputText.Replace(","c, "."c)

                            If Double.TryParse(normalizedInput, NumberStyles.Float, CultureInfo.InvariantCulture, valD) Then
                                param.Value = valD
                            Else
                                Throw New Exception($"Invalid value for {param.Name}.")
                            End If
                        Else
                            ' Generic / string
                            If TypeOf param.InputControl Is TextBox Then
                                param.Value = CType(param.InputControl, TextBox).Text
                            End If
                        End If
                    Catch ex As Exception
                        ShowCustomMessageBox($"{ex.Message} Using original ('{If(param.Value Is Nothing, "Nothing", param.Value)}').")
                    End Try
                Next
            End If

            inputForm.Dispose()
            Return (result = DialogResult.OK)
        End Function

        Public Shared Function ShowCustomWindow(
            introLine As String,
            ByVal bodyText As String,
            finalRemark As String,
            header As String,
            Optional NoRTF As Boolean = False,
            Optional Getfocus As Boolean = False,
            Optional InsertMarkdown As Boolean = False,
            Optional TransferToPane As Boolean = False,
            Optional parentWindowHwnd As IntPtr = Nothing,
            Optional PreserveLiterals As Boolean = False
        ) As String

            ' Ursprünglichen Text merken
            Dim OriginalText As String = bodyText

            ' --- Abstände & Konstanten ---
            Const leftMargin As Integer = 10
            Const rightPadding As Integer = 10
            Const spacing As Integer = 10
            Const gapButtons As Integer = 10
            Const remarkToButtonSpacing As Integer = 20
            Const bottomPadding As Integer = 20

            ' --- Controls anlegen ---
            Dim styledForm As New System.Windows.Forms.Form()
            Dim introLabel As New System.Windows.Forms.Label()
            Dim bodyTextBox As New RichTextBox()
            Dim finalRemarkLabel As New System.Windows.Forms.Label()
            Dim btnEdited As New System.Windows.Forms.Button()
            Dim btnOriginal As New System.Windows.Forms.Button()
            Dim btnMark As New System.Windows.Forms.Button()
            Dim btnPane As New System.Windows.Forms.Button()
            Dim btnCancel As New System.Windows.Forms.Button()
            Dim toolStrip As New System.Windows.Forms.ToolStrip()
            Dim lblHint As New System.Windows.Forms.Label() With {
        .AutoSize = False,
        .TextAlign = ContentAlignment.MiddleRight
    }

            ' --- Screen / Max-Größe berechnen ---
            Dim scrW = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width
            Dim scrH = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height
            Dim maxW = scrW \ 2
            Dim maxH = Math.Min(scrH \ 2, (maxW * 9) \ 16)
            maxW = Math.Min(maxW, (maxH * 16) \ 9)

            ' --- Fallback–Minima ---
            Const minFormWStatic As Integer = 400
            Const minFormHStatic As Integer = 300

            ' --- Formular-Eigenschaften ---
            styledForm.Text = header
            styledForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            styledForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            styledForm.MaximizeBox = True
            styledForm.MinimizeBox = False
            styledForm.ShowInTaskbar = False
            styledForm.TopMost = True
            styledForm.CancelButton = btnCancel
            styledForm.MinimumSize = New System.Drawing.Size(minFormWStatic, minFormHStatic)

            ' Icon
            Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
            styledForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

            ' Einheitliche Schrift
            Dim stdFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            styledForm.Font = stdFont

            ' --- Intro-Label ---
            introLabel.Text = introLine
            introLabel.Font = stdFont
            introLabel.AutoSize = False
            introLabel.Location = New System.Drawing.Point(leftMargin, spacing)
            introLabel.Width = maxW - leftMargin - rightPadding
            introLabel.Height = introLabel.PreferredHeight
            introLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            styledForm.Controls.Add(introLabel)

            ' --- Buttons ---
            btnEdited.Text = "OK, use edited text"
            Dim szE = TextRenderer.MeasureText(btnEdited.Text, stdFont)
            btnEdited.Size = New Size(szE.Width + 20, szE.Height + 10)

            btnOriginal.Text = "OK, use original text"
            Dim szO = TextRenderer.MeasureText(btnOriginal.Text, stdFont)
            btnOriginal.Size = New Size(szO.Width + 20, szE.Height + 10)

            If TransferToPane Then
                btnPane.Text = "Transfer to pane"
                Dim szP = TextRenderer.MeasureText(btnPane.Text, stdFont)
                btnPane.Size = New Size(szP.Width + 20, szE.Height + 10)
                styledForm.Controls.Add(btnPane)
            End If

            If InsertMarkdown Then
                btnMark.Text = "Insert original text with formatting"
                Dim szM = TextRenderer.MeasureText(btnMark.Text, stdFont)
                btnMark.Size = New Size(szM.Width + 20, szE.Height + 10)
                styledForm.Controls.Add(btnMark)
            End If

            btnCancel.Text = "Cancel"
            Dim szC = TextRenderer.MeasureText(btnCancel.Text, stdFont)
            btnCancel.Size = New Size(szC.Width + 20, szE.Height + 10)

            styledForm.Controls.Add(btnEdited)
            styledForm.Controls.Add(btnOriginal)
            styledForm.Controls.Add(btnCancel)

            ' --- BodyTextBox (align with CustomPaneControl) ---
            bodyTextBox.Font = New System.Drawing.Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
            bodyTextBox.Multiline = True
            bodyTextBox.ScrollBars = RichTextBoxScrollBars.Vertical
            bodyTextBox.WordWrap = True
            bodyTextBox.HideSelection = False
            bodyTextBox.DetectUrls = True
            bodyTextBox.Location = New System.Drawing.Point(leftMargin, introLabel.Bottom + spacing)
            bodyTextBox.Width = maxW - leftMargin - rightPadding
            bodyTextBox.Height = maxH - introLabel.Bottom - spacing
            bodyTextBox.MinimumSize = New Size(bodyTextBox.Width, bodyTextBox.Height)
            bodyTextBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            styledForm.Controls.Add(bodyTextBox)

            ' LinkClicked: open directly (no Ctrl modifier), like CustomPaneControl
            AddHandler bodyTextBox.LinkClicked,
        Sub(senderObj As Object, e As LinkClickedEventArgs)
            Try
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(e.LinkText) With {.UseShellExecute = True})
            Catch
                ' ignore
            End Try
        End Sub

            ' Copy handler: match CustomPaneControl behavior
            AddHandler bodyTextBox.KeyDown,
        Sub(sender As Object, e As System.Windows.Forms.KeyEventArgs)
            If (e.Control AndAlso (e.KeyCode = Keys.C OrElse e.KeyCode = Keys.Insert)) Then
                Try
                    If Not NoRTF Then
                        SharedMethods.CopySelectionExcludingTrailingNbsp(bodyTextBox)
                    Else
                        If bodyTextBox.SelectionLength > 0 Then
                            SharedMethods.PutInClipboard(bodyTextBox.SelectedText)
                        Else
                            SharedMethods.PutInClipboard(bodyTextBox.Text)
                        End If
                    End If
                    e.Handled = True
                Catch
                    ' fallback to default if anything goes wrong
                End Try
            End If
            ' Do not intercept Ctrl+A (same as CustomPaneControl)
        End Sub

            ' --- Optionales End-Label ---
            Dim hasRemark = Not String.IsNullOrEmpty(finalRemark)
            If hasRemark Then
                finalRemarkLabel.Text = finalRemark
                finalRemarkLabel.Font = stdFont
                finalRemarkLabel.AutoSize = False
                finalRemarkLabel.Width = bodyTextBox.MinimumSize.Width
                finalRemarkLabel.Height = finalRemarkLabel.GetPreferredSize(New Size(finalRemarkLabel.Width, 0)).Height
                finalRemarkLabel.Anchor = AnchorStyles.Left Or AnchorStyles.Right
                styledForm.Controls.Add(finalRemarkLabel)
            End If

            ' --- ToolStrip ---
            toolStrip.Dock = DockStyle.None
            For Each sym In New String() {"B", "I", "U", "•"}
                Dim tsb As New ToolStripButton(sym) With {
            .Font = New System.Drawing.Font(stdFont, If(sym = "B",
                                                FontStyle.Bold,
                                                If(sym = "I",
                                                   FontStyle.Italic,
                                                   If(sym = "U",
                                                      FontStyle.Underline,
                                                      FontStyle.Regular)))),
            .Name = "tsb" & sym
        }
                AddHandler tsb.Click,
            Sub(s, e)
                If bodyTextBox.SelectionLength > 0 Then
                    Select Case DirectCast(s, ToolStripButton).Name
                        Case "tsbB"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Bold)
                        Case "tsbI"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Italic)
                        Case "tsbU"
                            bodyTextBox.SelectionFont = New System.Drawing.Font(bodyTextBox.SelectionFont, bodyTextBox.SelectionFont.Style Xor FontStyle.Underline)
                        Case "tsb•"
                            bodyTextBox.SelectionIndent = If(bodyTextBox.SelectionIndent = 20, 0, 20)
                            bodyTextBox.SelectionBullet = Not bodyTextBox.SelectionBullet
                            bodyTextBox.BulletIndent = If(bodyTextBox.BulletIndent = 15, 0, 15)
                    End Select
                End If
            End Sub
                toolStrip.Items.Add(tsb)
            Next
            styledForm.Controls.Add(toolStrip)

            ' Hint label (update text; no Ctrl+Click requirement)
            lblHint.Text = "Click a link to open"
            lblHint.Font = New System.Drawing.Font(stdFont, FontStyle.Italic)
            lblHint.ForeColor = Color.DimGray
            lblHint.Height = szE.Height + 6
            styledForm.Controls.Add(lblHint)

            ' --- Dynamische Mindestgröße ---
            Dim bodyTop = bodyTextBox.Top
            Dim bodyMinH = bodyTextBox.MinimumSize.Height
            Dim remHeight = If(hasRemark,
               finalRemarkLabel.GetPreferredSize(New Size(bodyTextBox.MinimumSize.Width, 0)).Height,
               0)
            Dim btnH = btnEdited.Height

            Dim dynamicMinH = bodyTop +
              bodyMinH +
              If(hasRemark,
                 spacing + remHeight + remarkToButtonSpacing,
                 remarkToButtonSpacing) +
              btnH +
              bottomPadding

            Dim w1 = leftMargin + bodyTextBox.MinimumSize.Width + rightPadding
            Dim introMinW = leftMargin + introLabel.PreferredWidth + rightPadding
            Dim totalBtnW = btnEdited.Width + gapButtons + btnOriginal.Width +
            If(InsertMarkdown, gapButtons + btnMark.Width, 0) +
            If(TransferToPane, gapButtons + btnPane.Width, 0) +
            gapButtons + btnCancel.Width
            Dim w3 = leftMargin + totalBtnW + rightPadding
            Dim dynamicMinW = Math.Max(Math.Max(w1, introMinW), w3)

            styledForm.MinimumSize = New Size(
        Math.Max(minFormWStatic, dynamicMinW),
        Math.Max(minFormHStatic, dynamicMinH)
    )

            ' --- Resize Handler ---
            AddHandler styledForm.Resize,
        Sub(s, e)
            Dim fW = styledForm.ClientSize.Width
            Dim fH = styledForm.ClientSize.Height

            introLabel.Width = fW - leftMargin - rightPadding

            Dim newW = fW - leftMargin - rightPadding
            bodyTextBox.Width = Math.Max(bodyTextBox.MinimumSize.Width, newW)

            Dim usedBelow = If(hasRemark,
                               spacing + finalRemarkLabel.Height + remarkToButtonSpacing,
                               remarkToButtonSpacing) +
                            btnH + bottomPadding
            Dim availH = fH - bodyTop - usedBelow
            bodyTextBox.Height = Math.Max(bodyTextBox.MinimumSize.Height, availH)

            If hasRemark Then
                finalRemarkLabel.Width = bodyTextBox.Width
                finalRemarkLabel.Height = finalRemarkLabel.GetPreferredSize(New Size(finalRemarkLabel.Width, 0)).Height
                finalRemarkLabel.Location = New System.Drawing.Point(leftMargin, bodyTextBox.Bottom + spacing)
            End If

            Dim btnY = fH - btnH - bottomPadding
            btnEdited.Location = New System.Drawing.Point(leftMargin, btnY)
            btnOriginal.Location = New System.Drawing.Point(btnEdited.Right + gapButtons, btnY)

            Dim nextX = btnOriginal.Right
            If InsertMarkdown Then
                btnMark.Location = New System.Drawing.Point(btnOriginal.Right + gapButtons, btnY)
                nextX = btnMark.Right
            End If
            If TransferToPane Then
                btnPane.Location = New System.Drawing.Point(nextX + gapButtons, btnY)
                nextX = btnPane.Right
            End If
            btnCancel.Location = New System.Drawing.Point(nextX + gapButtons, btnY)

            ' Toolstrip above textbox right aligned
            toolStrip.Location = New System.Drawing.Point(
                leftMargin + bodyTextBox.Width - toolStrip.Width,
                bodyTextBox.Top - toolStrip.Height - spacing
            )
            toolStrip.BringToFront()

            ' Hint label aligns with right edge above buttons
            lblHint.Width = 180
            lblHint.Location = New System.Drawing.Point(fW - lblHint.Width - rightPadding, introLabel.Top)
        End Sub

            ' --- Initialgröße ---
            Dim initW = Math.Max(maxW, styledForm.MinimumSize.Width)
            Dim initH = Math.Max(maxH, styledForm.MinimumSize.Height)
            styledForm.ClientSize = New Size(initW, initH)
            styledForm.PerformLayout()
            styledForm.MinimumSize = styledForm.Size

            ' --- Content assignment (match CustomPaneControl) ---
            Dim rtf As String = Nothing
            If Not NoRTF Then
                rtf = MarkdownToRtfConverter.Convert(bodyText, PreserveLiterals)
                Debug.WriteLine("Converted RTF: " & rtf)
            End If

            Try
                If NoRTF Then
                    bodyTextBox.Text = bodyText
                Else
                    bodyTextBox.Rtf = rtf
                    ' Append NBSPs for hyperlinks (same as CustomPaneControl)
                    SharedMethods.AppendNbspForHyperlinks(bodyTextBox, rtf)
                End If
            Catch ex As System.ComponentModel.Win32Exception
                bodyTextBox.Text = bodyText
            Catch
                bodyTextBox.Text = bodyText
            End Try

            ' Ensure URL detection is enabled (same as CustomPaneControl)
            bodyTextBox.DetectUrls = True
            bodyTextBox.Select(0, 0)

            Dim OriginalTextBox As String = bodyTextBox.Text

            ' --- Button-Handler ---
            Dim returnValue As String = String.Empty

            AddHandler btnEdited.Click,
        Sub()
            returnValue = If(NoRTF, bodyTextBox.Text, bodyTextBox.Rtf)
            styledForm.DialogResult = DialogResult.OK
            styledForm.Close()
        End Sub

            AddHandler btnOriginal.Click,
        Sub()
            returnValue = If(NoRTF, OriginalText, If(rtf, bodyText))
            styledForm.DialogResult = DialogResult.OK
            styledForm.Close()
        End Sub

            If InsertMarkdown Then
                AddHandler btnMark.Click,
            Sub()
                returnValue = "Markdown"
                styledForm.DialogResult = DialogResult.OK
                styledForm.Close()
            End Sub
            End If

            If TransferToPane Then
                AddHandler btnPane.Click,
            Sub()
                If bodyTextBox.Text.Trim() = OriginalTextBox.Trim() OrElse
                   ShowCustomYesNoBox($"Your changes will be lost and the pane will again show the original text (unless you put it in the clipboard manually). Continue?", "Yes", "No") = 1 Then
                    returnValue = "Pane"
                    styledForm.DialogResult = DialogResult.OK
                    styledForm.Close()
                End If
            End Sub
            End If

            AddHandler btnCancel.Click,
        Sub()
            returnValue = String.Empty
            styledForm.DialogResult = DialogResult.Cancel
            styledForm.Close()
        End Sub

            ' --- Dialog anzeigen ---
            styledForm.BringToFront()
            styledForm.Focus()
            styledForm.Activate()

            AddHandler styledForm.Shown,
                    Sub(sender, e)
                        styledForm.TopMost = False  ' Reset first
                        styledForm.TopMost = True   ' Then set again
                        styledForm.Activate()
                        styledForm.BringToFront()
                    End Sub

            If parentWindowHwnd <> IntPtr.Zero Then
                styledForm.ShowDialog(New WindowWrapper(parentWindowHwnd))
            ElseIf Getfocus Then
                Dim officeHwnd As IntPtr = GetOfficeApplicationHwnd()
                If officeHwnd <> IntPtr.Zero Then
                    styledForm.ShowDialog(New WindowWrapper(officeHwnd))
                Else
                    styledForm.ShowDialog()
                End If
            Else
                styledForm.ShowDialog()
            End If

            Return returnValue
        End Function


        Public Class InputParameter
            Public Property Name As System.String
            Public Property Value As System.Object
            Public Property Options As System.Collections.Generic.List(Of System.String)
            Public Property InputControl As System.Windows.Forms.Control

            ' Wichtig: parameterloser Ctor (wird bei "New InputParameter() With {...}" benötigt)
            Public Sub New()
                Me.Options = New System.Collections.Generic.List(Of System.String)()
            End Sub

            ' Constructor für einfache Parameter
            Public Sub New(ByVal name As System.String, ByVal value As System.Object)
                Me.New()
                Me.Name = name
                Me.Value = value
            End Sub

            ' Overload für Parameter mit Options
            Public Sub New(ByVal name As System.String,
                   ByVal value As System.Object,
                   ByVal options As System.Collections.Generic.IEnumerable(Of System.String))
                Me.New()
                Me.Name = name
                Me.Value = value
                If options IsNot Nothing Then
                    Me.Options = New System.Collections.Generic.List(Of System.String)(options)
                End If
            End Sub
        End Class



    End Class
End Namespace
