' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary
    Public Class ModelSelectorForm
        Inherits Form

        Private lblTitle As System.Windows.Forms.Label
        Private lstModels As ListBox
        Private chkReset As System.Windows.Forms.CheckBox
        Private btnOK As Button
        Private btnCancel As Button

        Private alternativeModels As List(Of ModelConfig)
        Private hasDefaultEntry As Boolean

        Public Shared ReadOnly ButtonTextPadding As System.Windows.Forms.Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)

        ' The selected alternative model (if any).
        Public Property SelectedModel As ModelConfig = Nothing
        ' True if the default configuration is to be used.
        Public Property UseDefault As Boolean = True

        Public Sub New(ByVal iniFilePath As String, ByVal context As ISharedContext, ByVal Title As String, ByVal ListType As String, ByVal OptionText As String, Optional UseCase As Integer = 1)

            ' UseCase 1 = Model Selection (with Default) UseCase 2 = Model Selection (without Default)

            OptionChecked = True

            ' --- DPI- und Font-Skalierung aktivieren ---
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Icon = Icon.FromHandle((New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)).GetHicon())
            Me.Text = Title

            ' Haupt-TableLayoutPanel mit 4 Zeilen
            Dim tlpMain As New System.Windows.Forms.TableLayoutPanel() With {
                                            .Dock = System.Windows.Forms.DockStyle.Fill,
                                            .ColumnCount = 1,
                                            .RowCount = 4
                                        }
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Zeile 1: Label
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F)) ' Zeile 2: ListBox
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Zeile 3: Checkbox
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Zeile 4: Buttons

            ' Zeile 1: Label (shrinks & grows, 20px Padding)
            lblTitle = New System.Windows.Forms.Label() With {
                                            .Text = ListType,
                                            .AutoSize = True,
                                            .Dock = System.Windows.Forms.DockStyle.Fill,
                                            .Margin = New System.Windows.Forms.Padding(20, 20, 20, 0)
                                        }
            tlpMain.Controls.Add(lblTitle, 0, 0)

            ' Zeile 2: ListBox (shrinks & grows, 20px Padding)
            lstModels = New System.Windows.Forms.ListBox() With {
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .Margin = New System.Windows.Forms.Padding(20)
                                    }
            tlpMain.Controls.Add(lstModels, 0, 1)


            ' Zeile 3: Checkbox (grows but not shrink, 20px Padding)
            chkReset = New System.Windows.Forms.CheckBox() With {
                                        .Text = OptionText,
                                        .Checked = OptionChecked,
                                        .AutoSize = True,
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .Margin = New System.Windows.Forms.Padding(20, 0, 20, 0)
                                    }

            If OptionText <> "" Then
                tlpMain.Controls.Add(chkReset, 0, 2)
            End If

            ' Zeile 4: Buttons (links-nach-rechts, grows but not shrink, 20px Padding)
            Dim flpButtons As New System.Windows.Forms.FlowLayoutPanel() With {
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                        .AutoSize = True,
                                        .Margin = New System.Windows.Forms.Padding(20)
                                    }
            btnOK = New System.Windows.Forms.Button() With {
                                        .Text = "OK",
                                        .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                                    }
            AddHandler btnOK.Click, AddressOf btnOK_Click
            flpButtons.Controls.Add(btnOK)

            btnCancel = New System.Windows.Forms.Button() With {
                                        .Text = "Cancel",
                                        .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                                    }
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            flpButtons.Controls.Add(btnCancel)

            tlpMain.Controls.Add(flpButtons, 0, 3)

            Me.Controls.Add(tlpMain)
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            ' Modelle laden
            alternativeModels = LoadAlternativeModels(iniFilePath, context)
            If UseCase = 1 Then
                lstModels.Items.Add("Default = " & context.INI_Model_2)
                hasDefaultEntry = True
            Else
                hasDefaultEntry = False
            End If
            For Each model In alternativeModels
                Dim displayText As String = If(String.IsNullOrEmpty(model.ModelDescription), model.Model, model.ModelDescription)
                lstModels.Items.Add(displayText)
            Next
            lstModels.SelectedIndex = 0
            AddHandler lstModels.DoubleClick, AddressOf lstModels_DoubleClick

            Me.ClientSize = New System.Drawing.Size(580, 450)
            Me.MinimumSize = Me.Size
        End Sub

        Private Sub lstModels_DoubleClick(sender As Object, e As System.EventArgs)
            If lstModels.SelectedIndex >= 0 Then
                btnOK.PerformClick()
            End If
        End Sub


        Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
            MyBase.OnHandleCreated(e)
            Dim dpiScale As Single = Me.DeviceDpi / 96.0F
            If dpiScale <> 1.0F Then
                Me.Scale(New System.Drawing.SizeF(dpiScale, dpiScale))
            End If
        End Sub



        Private Sub btnOK_Click(sender As Object, e As EventArgs)
            Try
                If hasDefaultEntry AndAlso lstModels.SelectedIndex = 0 Then
                    UseDefault = True
                Else
                    UseDefault = False
                    ' adjust the index offset by 1 if there was a default entry
                    Dim offset As Integer = If(hasDefaultEntry, 1, 0)
                    Dim idx As Integer = lstModels.SelectedIndex - offset
                    If idx >= 0 AndAlso idx < alternativeModels.Count Then
                        SelectedModel = alternativeModels(idx)
                    End If
                End If

                ' If the checkbox is unchecked and a non-default model is selected, set OriginalConfigurationLoaded to False.
                If chkReset IsNot Nothing Then
                    If Not chkReset.Checked AndAlso Not UseDefault Then
                        originalConfigLoaded = False
                    End If
                    OptionChecked = chkReset.Checked
                Else
                    OptionChecked = True
                    If Not UseDefault Then
                        originalConfigLoaded = False
                    End If
                End If

                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As System.Exception
                MessageBox.Show("Error processing selection: " & ex.Message)
            End Try
        End Sub

        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

    End Class

End Namespace
