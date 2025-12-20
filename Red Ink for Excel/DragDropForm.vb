Imports System.Windows.Forms
Imports System.Drawing

Public Class DragDropForm

    Private _selectedFilePath As String = String.Empty

    Public ReadOnly Property SelectedFilePath As String
            Get
                Return _selectedFilePath
            End Get
        End Property

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

    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
        Me.Icon = Icon.FromHandle(bmp.GetHicon())
    End Sub

    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
            ' Check if the data being dragged is a file
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.None
            End If
        End Sub

        Private Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
            Try
                ' Retrieve the file list
                Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
                If files IsNot Nothing AndAlso files.Length > 0 Then
                    _selectedFilePath = files(0) ' Take first file
                    ' Optionally close form automatically once a file is dropped
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                End If
            Catch ex As System.Exception
                MessageBox.Show($"Error: {ex.Message}")
            End Try
        End Sub

        Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
            Using ofd As New OpenFileDialog()

            If Globals.ThisAddIn.DragDropFormFilter = "" Then

                ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm||" &
                             "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Rich Text Files (*.rtf)|*.rtf|" &
                             "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                             "PDF Files (*.pdf)|*.pdf"

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