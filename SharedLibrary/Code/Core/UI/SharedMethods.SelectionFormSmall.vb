' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Structure SelectionItem
            Public ReadOnly DisplayText As String
            Public ReadOnly Value As Integer

            Public Sub New(text As String, value As Integer)
                Me.DisplayText = text
                Me.Value = value
            End Sub

            Public Overrides Function ToString() As String
                Return DisplayText
            End Function
        End Structure


        Friend NotInheritable Class SelectionFormSmall
            Inherits System.Windows.Forms.Form

            Private ReadOnly _lst As System.Windows.Forms.ListBox
            Private ReadOnly _lbl As System.Windows.Forms.Label
            Private _result As Integer = 0

            Friend Sub New(items As IReadOnlyList(Of SelectionItem),
                   defaultValue As Integer,
                   promptText As String,
                   Optional headerText As String = Nothing)

                Const baseWidth As Integer = 400
                Const sidePadding As Integer = 10
                Const bottomPadding As Integer = 24

                Me.SuspendLayout()

                Dim stdFont As New System.Drawing.Font("Segoe UI", 9.0F,
                                               System.Drawing.FontStyle.Regular,
                                               System.Drawing.GraphicsUnit.Point)
                Me.Font = stdFont
                Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

                If String.IsNullOrWhiteSpace(headerText) Then headerText = AN
                Me.Text = headerText

                Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                Me.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())

                Me.ClientSize = New System.Drawing.Size(baseWidth, 100)
                Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
                Me.KeyPreview = True

                _lbl = New System.Windows.Forms.Label With {
            .AutoSize = True,
            .Text = promptText,
            .Location = New System.Drawing.Point(sidePadding, sidePadding),
            .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                      System.Windows.Forms.AnchorStyles.Left Or
                      System.Windows.Forms.AnchorStyles.Right
        }
                Controls.Add(_lbl)

                ' Important: do NOT anchor Bottom yet (avoid early stretch)
                _lst = New System.Windows.Forms.ListBox With {
            .IntegralHeight = False,
            .SelectionMode = System.Windows.Forms.SelectionMode.One,
            .Anchor = System.Windows.Forms.AnchorStyles.Top Or
                      System.Windows.Forms.AnchorStyles.Left Or
                      System.Windows.Forms.AnchorStyles.Right
        }
                Dim visibleRows As Integer = Math.Min(5, items.Count)
                _lst.ItemHeight = CInt(stdFont.GetHeight())
                Dim desiredListHeight As Integer = _lst.ItemHeight * visibleRows + 9

                _lst.Location = New System.Drawing.Point(sidePadding, _lbl.Bottom + 10)
                _lst.Width = baseWidth - (2 * sidePadding)
                _lst.Height = desiredListHeight
                Controls.Add(_lst)

                For Each it In items : _lst.Items.Add(it) : Next

                Dim defIdx As Integer = items.ToList().FindIndex(Function(it) it.Value = defaultValue)
                If defIdx >= 0 Then
                    _lst.SelectedIndex = defIdx
                    _result = items(defIdx).Value
                End If

                ' Now compute final ClientSize based on desired list height + padding
                Dim requiredHeight As Integer = _lst.Top + desiredListHeight + bottomPadding
                Me.ClientSize = New System.Drawing.Size(baseWidth, requiredHeight)

                ' After ClientSize is finalized, enable Bottom anchoring
                _lst.Anchor = System.Windows.Forms.AnchorStyles.Top Or
                      System.Windows.Forms.AnchorStyles.Left Or
                      System.Windows.Forms.AnchorStyles.Right Or
                      System.Windows.Forms.AnchorStyles.Bottom

                ' Ensure width matches final client width
                _lst.Width = Me.ClientSize.Width - (2 * sidePadding)

                ' Optional: keep a reasonable minimum, using current size
                Me.MinimumSize = Me.Size

                AddHandler _lst.KeyDown,
            Sub(s, e)
                If e.KeyCode = System.Windows.Forms.Keys.Enter Then AcceptCurrentSelection()
            End Sub
                AddHandler _lst.DoubleClick, Sub() AcceptCurrentSelection()

                AddHandler Me.KeyDown,
            Sub(sender, e)
                If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                    _result = 0
                    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Close()
                End If
            End Sub

                AddHandler Me.FormClosing,
            Sub(s, e)
                If Me.DialogResult <> System.Windows.Forms.DialogResult.OK Then _result = 0
            End Sub

                ' Keep padding and width on resize
                AddHandler Me.Resize,
            Sub()
                _lbl.Width = Me.ClientSize.Width - (2 * sidePadding)
                _lst.Width = Me.ClientSize.Width - (2 * sidePadding)
                Dim newHeight = Me.ClientSize.Height - _lst.Top - bottomPadding
                If newHeight > 40 Then _lst.Height = newHeight
            End Sub

                Me.ResumeLayout(False)
                Me.PerformLayout()
                _lst.Focus()
            End Sub

            Private Sub AcceptCurrentSelection()
                If _lst.SelectedIndex >= 0 Then
                    Dim item As SelectionItem = DirectCast(_lst.SelectedItem, SelectionItem)
                    _result = item.Value
                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Close()
                End If
            End Sub

            Friend ReadOnly Property Result As Integer
                Get
                    Return _result
                End Get
            End Property
        End Class


        Public Shared Function SelectValue(items As IEnumerable(Of SelectionItem),
                                   defaultValue As Integer,
                                   Optional prompt As String = "Please choose …",
                                   Optional header As String = Nothing) As Integer

            If items Is Nothing Then
                System.Windows.Forms.MessageBox.Show("SelectValue Error: Items collection must not be null.")
                Return 0
            End If

            Using frm As New SelectionFormSmall(items.ToList(), defaultValue, prompt, header)
                frm.ShowDialog()
                Return frm.Result            ' now returns the correct integer
            End Using
        End Function



    End Class

End Namespace