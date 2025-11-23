' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Namespace SharedLibrary

    Partial Class MultiModelSelectorForm
        Inherits System.Windows.Forms.Form

        Private lblTitle As System.Windows.Forms.Label
        Private txtFilter As System.Windows.Forms.TextBox
        Private chkList As System.Windows.Forms.CheckedListBox
        Private chkReset As System.Windows.Forms.CheckBox
        Private btnOK As Button
        Private btnCancel As Button
        Private pnlButtons As System.Windows.Forms.FlowLayoutPanel
        Private outer As System.Windows.Forms.TableLayoutPanel

        Private displayToModel As New System.Collections.Generic.Dictionary(Of String, ModelConfig)(System.StringComparer.OrdinalIgnoreCase)
        Private seenDisplays As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
        Private allDisplayItems As New System.Collections.Generic.List(Of String)
        Private preselectKey As String = Nothing
        Private ReadOnly altModels As System.Collections.Generic.List(Of ModelConfig)

        ' Persist checked selections across filtering
        Private ReadOnly selectedLabels As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
        Private isUpdating As Boolean = False

        ' EM_SETCUEBANNER to show a cue banner ("placeholder") on Win32 edit controls
        Private Const EM_SETCUEBANNER As Integer = &H1501
        <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
        End Function

        Public ReadOnly Property SelectedModels As System.Collections.Generic.List(Of ModelConfig)
            Get
                ' Return models based on the persisted selectedLabels (not only the currently visible items)
                Dim result As New System.Collections.Generic.List(Of ModelConfig)
                For Each key In selectedLabels
                    If displayToModel.ContainsKey(key) Then
                        result.Add(displayToModel(key))
                    End If
                Next
                Return result
            End Get
        End Property

        Public ReadOnly Property UseDefault As Boolean
            Get
                Return chkReset.Checked
            End Get
        End Property

        Public Sub New(models As System.Collections.Generic.List(Of ModelConfig),
                   preselect As System.String,
                   Optional title As System.String = Nothing,
                   Optional resetChecked As System.Boolean = True)
            Me.altModels = If(models, New System.Collections.Generic.List(Of ModelConfig))
            Me.preselectKey = preselect
            InitializeComponent(title, resetChecked)
            PopulateList()
            ApplyPreselection()
        End Sub

        Private Sub InitializeComponent(Optional title As System.String = Nothing, Optional resetChecked As System.Boolean = True)
            Me.Text = If(String.IsNullOrWhiteSpace(title), SharedMethods.AN & " - Select Alternate Models", title)
            Me.Icon = Icon.FromHandle((New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)).GetHicon())
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.MinimizeBox = True
            Me.MaximizeBox = True
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            Me.Width = 520
            Me.Height = 460
            Me.MinimumSize = New System.Drawing.Size(520, 460)

            Me.outer = New System.Windows.Forms.TableLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New System.Windows.Forms.Padding(16, 12, 16, 12)
            }
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))
            Me.outer.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

            Me.lblTitle = New System.Windows.Forms.Label() With {
                .Text = "Select one or more alternate models:",
                .Dock = System.Windows.Forms.DockStyle.Top,
                .Height = 28,
                .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }

            Me.txtFilter = New System.Windows.Forms.TextBox() With {
                .Dock = System.Windows.Forms.DockStyle.Top
            }
            AddHandler Me.txtFilter.HandleCreated,
                Sub()
                    Try
                        Dim showEvenIfFocused As IntPtr = CType(1, IntPtr)
                        SendMessage(Me.txtFilter.Handle, EM_SETCUEBANNER, showEvenIfFocused, "Filter models…")
                    Catch
                    End Try
                End Sub
            AddHandler Me.txtFilter.TextChanged, AddressOf OnFilterChanged

            Me.chkList = New System.Windows.Forms.CheckedListBox() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .CheckOnClick = True
            }
            AddHandler Me.chkList.DoubleClick, AddressOf OnListDoubleClick
            AddHandler Me.chkList.ItemCheck, AddressOf OnItemCheck

            Me.chkReset = New System.Windows.Forms.CheckBox() With {
                .Text = "Reset to default model after use",
                .Dock = System.Windows.Forms.DockStyle.Top,
                .Checked = resetChecked,
                .Visible = False            ' Hide it for the time being
            }

            Me.pnlButtons = New System.Windows.Forms.FlowLayoutPanel() With {
                .Dock = System.Windows.Forms.DockStyle.Fill,
                .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                .Padding = New System.Windows.Forms.Padding(0, 8, 0, 0),
                .AutoSize = True
            }

            ' Match ModelSelectorForm: autosize, GrowAndShrink, and padding (10,5,10,5)
            Me.btnOK = New System.Windows.Forms.Button() With {
                .Text = "OK",
                .DialogResult = System.Windows.Forms.DialogResult.OK,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5)
            }
            Me.btnCancel = New System.Windows.Forms.Button() With {
                .Text = "Cancel",
                .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                .AutoSize = True,
                .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5)
            }
            Me.pnlButtons.Controls.Add(Me.btnOK)
            Me.pnlButtons.Controls.Add(Me.btnCancel)

            Me.outer.Controls.Add(Me.lblTitle, 0, 0)
            Me.outer.Controls.Add(Me.txtFilter, 0, 1)
            Me.outer.Controls.Add(Me.chkList, 0, 2)
            Me.outer.Controls.Add(Me.chkReset, 0, 3)
            Me.outer.Controls.Add(Me.pnlButtons, 0, 4)
            Me.Controls.Add(Me.outer)

            Me.AcceptButton = Me.btnOK
            Me.CancelButton = Me.btnCancel
        End Sub

        Private Function MakeUniqueDisplay(baseText As System.String) As System.String
            Dim s As System.String = If(String.IsNullOrWhiteSpace(baseText), "(Unnamed model)", baseText.Trim())
            Dim unique As System.String = s
            Dim suffix As Integer = 2
            While seenDisplays.Contains(unique)
                unique = s & " (" & suffix.ToString() & ")"
                suffix += 1
            End While
            seenDisplays.Add(unique)
            Return unique
        End Function

        Private Sub PopulateList()
            displayToModel.Clear()
            seenDisplays.Clear()
            allDisplayItems.Clear()

            For Each m In altModels
                Dim display As System.String = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                Dim unique As System.String = MakeUniqueDisplay(display)
                displayToModel(unique) = m
                allDisplayItems.Add(unique)
            Next

            ' Build initial visible list (no filter), preserving any pre-existing selections
            isUpdating = True
            Try
                Me.chkList.Items.Clear()
                For Each label In allDisplayItems
                    Me.chkList.Items.Add(label, selectedLabels.Contains(label))
                Next
            Finally
                isUpdating = False
            End Try
        End Sub

        Private Sub OnFilterChanged(sender As System.Object, e As System.EventArgs)
            Dim filter As System.String = If(Me.txtFilter.Text, String.Empty).Trim().ToLowerInvariant()

            isUpdating = True
            Me.chkList.BeginUpdate()
            Try
                Me.chkList.Items.Clear()
                For Each itemText In allDisplayItems
                    If filter.Length = 0 OrElse itemText.ToLowerInvariant().Contains(filter) Then
                        ' Restore check state from the persisted selection set
                        Dim isChecked = selectedLabels.Contains(itemText)
                        Me.chkList.Items.Add(itemText, isChecked)
                    End If
                Next
            Finally
                Me.chkList.EndUpdate()
                isUpdating = False
            End Try
        End Sub

        Private Sub OnItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs)
            If isUpdating Then Return
            Dim label As String = Me.chkList.Items(e.Index).ToString()
            If e.NewValue = CheckState.Checked Then
                selectedLabels.Add(label)
            Else
                selectedLabels.Remove(label)
            End If
        End Sub

        Private Sub OnListDoubleClick(sender As System.Object, e As System.EventArgs)
            Dim idx As Integer = Me.chkList.SelectedIndex
            If idx >= 0 Then
                Dim state As Boolean = Not Me.chkList.GetItemChecked(idx)
                Me.chkList.SetItemChecked(idx, state)
            End If
        End Sub

        Private Sub ApplyPreselection()
            If String.IsNullOrWhiteSpace(preselectKey) Then Return

            ' Try by label first
            For i = 0 To Me.chkList.Items.Count - 1
                Dim label As System.String = Me.chkList.Items(i).ToString()
                If String.Equals(label, preselectKey, System.StringComparison.OrdinalIgnoreCase) Then
                    Me.chkList.SetItemChecked(i, True)
                    ' selectedLabels will be updated by ItemCheck handler
                    Return
                End If
            Next

            ' Fallback: try to match underlying ModelDescription/Model
            Dim idxToCheck As Integer = -1
            For j = 0 To Me.chkList.Items.Count - 1
                Dim label As System.String = Me.chkList.Items(j).ToString()
                If displayToModel.ContainsKey(label) Then
                    Dim mc As ModelConfig = displayToModel(label)
                    If String.Equals(mc.ModelDescription, preselectKey, System.StringComparison.OrdinalIgnoreCase) _
                   OrElse String.Equals(mc.Model, preselectKey, System.StringComparison.OrdinalIgnoreCase) Then
                        idxToCheck = j
                        Exit For
                    End If
                End If
            Next
            If idxToCheck >= 0 Then
                Me.chkList.SetItemChecked(idxToCheck, True)
                ' selectedLabels will be updated by ItemCheck handler
            End If
        End Sub
    End Class


End Namespace
