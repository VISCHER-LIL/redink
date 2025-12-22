' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DragDropForm.vb
' Purpose: Provides a drag-and-drop interface for file selection with browse fallback.
'          Stores the selected file path and returns DialogResult.OK upon selection.
'
' Architecture:
'  - Drag-and-Drop Support: Enables AllowDrop and handles DragEnter/DragDrop events
'    to accept file drops (takes first file from drop operation).
'  - Browse Button: Opens OpenFileDialog with configurable filter (uses global
'    settings from Globals.ThisAddIn.DragDropFormFilter or default supported extensions).
'  - Customization: Form title and label text can be configured via Globals.ThisAddIn
'    properties (DragDropFormLabel).
'  - Result: Exposes SelectedFilePath property containing the chosen file path;
'    sets DialogResult.OK and closes form upon successful selection.
' =============================================================================

Imports System.Windows.Forms
Imports System.Drawing

Public Class DragDropForm

    Private _selectedFilePath As String = String.Empty

    ''' <summary>
    ''' Gets the file path selected by the user via drag-and-drop or browse dialog.
    ''' </summary>
    Public ReadOnly Property SelectedFilePath As String
        Get
            Return _selectedFilePath
        End Get
    End Property

    ''' <summary>
    ''' Initializes the form with drag-and-drop enabled and optional custom label text.
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ' Ensure drag and drop is enabled
        Me.AllowDrop = True
        ' Adjust form properties as needed
        Me.Text = "Drag & Drop Your File or Click Browse"
        If Globals.ThisAddIn.DragDropFormLabel <> "" Then
            Me.Label2.Text = Globals.ThisAddIn.DragDropFormLabel
        End If
    End Sub

    ''' <summary>
    ''' Sets the form icon from application resources on load.
    ''' </summary>
    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
        Dim icon As Icon = Icon.FromHandle(bmp.GetHicon())
        Me.Icon = icon
        ' Dispose bitmap to release GDI resources
        bmp.Dispose()
    End Sub

    ''' <summary>
    ''' Handles drag-enter event to accept file drops with copy effect.
    ''' </summary>
    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        ' Check if the data being dragged is a file
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ''' <summary>
    ''' Handles drag-drop event to capture the first dropped file and close the form with DialogResult.OK.
    ''' </summary>
    Private Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Try
            ' Retrieve the file list
            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing AndAlso files.Length > 0 Then
                _selectedFilePath = files(0) ' Take first file
                ' Close form automatically once a file is dropped
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        Catch ex As System.Exception
            MessageBox.Show($"Error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Opens file browse dialog with configurable filter to select a file and close the form with DialogResult.OK.
    ''' Uses Globals.ThisAddIn.DragDropFormFilter if set, otherwise applies default supported file extensions.
    ''' </summary>
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Using ofd As New OpenFileDialog()

            If Globals.ThisAddIn.DragDropFormFilter = "" Then

                ' Default filter covering text, document, and presentation formats
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.pptx;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Rich Text Files (*.rtf)|*.rtf|" &
                             "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                             "PDF Files (*.pdf)|*.pdf|" &
                             "Powerpoint Files (*.pptx)|*.pptx|" &
                             "All Files (*.*)|*.*"

            Else

                ofd.Filter = Globals.ThisAddIn.DragDropFormFilter

            End If

            ofd.Title = "Select a File"
            ofd.Multiselect = False

            If ofd.ShowDialog() = DialogResult.OK Then
                _selectedFilePath = ofd.FileName
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

End Class