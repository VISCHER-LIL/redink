' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Namespace SharedLibrary

    Friend Module ClipboardHelper

        Private Sub SafeReleaseCom(obj As Object)
            Try
                If obj IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(obj) Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
                End If
            Catch
                ' ignore
            End Try
        End Sub

        Friend Function TryGetClipboardObject(ByRef mimeType As String, ByRef base64 As String) As Boolean
            Dim succeeded As Boolean = False
            Dim localMimeType As String = Nothing
            Dim localBase64 As String = Nothing

            Dim t As New System.Threading.Thread(
Sub()
    Try
        ' 1) Outlook attachment (FileGroupDescriptorW / FileGroupDescriptor + FileContents)
        Dim hasW = System.Windows.Forms.Clipboard.ContainsData("FileGroupDescriptorW")
        Dim hasA = System.Windows.Forms.Clipboard.ContainsData("FileGroupDescriptor")
        If hasW OrElse hasA Then
            Dim fmt = If(hasW, "FileGroupDescriptorW", "FileGroupDescriptor")
            Dim fgObj = System.Windows.Forms.Clipboard.GetData(fmt)
            Dim fgStream = TryCast(fgObj, System.IO.MemoryStream)
            Try
                If fgStream IsNot Nothing Then
                    Using reader As New System.IO.BinaryReader(fgStream, System.Text.Encoding.Unicode, leaveOpen:=False)
                        ' skip itemCount + fixed fields
                        reader.ReadInt32() ' itemCount
                        reader.BaseStream.Seek(4 + 16 + 8 + 8 + 8 + 4 + 4, System.IO.SeekOrigin.Current)
                        ' read filename (up to 260 WCHARs)
                        Dim nameChars As New System.Collections.Generic.List(Of Char)
                        For i = 0 To 259
                            Dim ch As Char = reader.ReadChar()
                            If ch = ChrW(0) Then Exit For
                            nameChars.Add(ch)
                        Next
                        Dim fileName As String = New String(nameChars.ToArray())

                        ' pull the raw attachment bytes
                        Dim contentObj = System.Windows.Forms.Clipboard.GetData("FileContents")
                        Dim contentStream = TryCast(contentObj, System.IO.Stream)
                        Try
                            If contentStream IsNot Nothing Then
                                Using ms As New System.IO.MemoryStream()
                                    contentStream.CopyTo(ms)
                                    Dim bytes() As Byte = ms.ToArray()

                                    ' 2) WAV-header sniff
                                    If bytes.Length >= 12 AndAlso
                                       System.Text.Encoding.ASCII.GetString(bytes, 0, 4) = "RIFF" AndAlso
                                       System.Text.Encoding.ASCII.GetString(bytes, 8, 4) = "WAVE" Then

                                        localMimeType = "audio/wav"
                                    Else
                                        ' 3) fallback to extension-based mapping
                                        Dim ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant()
                                        Select Case ext
                                            Case ".wav" : localMimeType = "audio/wav"
                                            Case ".mp3" : localMimeType = "audio/mpeg"
                                            Case ".txt" : localMimeType = "text/plain"
                                            Case ".png" : localMimeType = "image/png"
                                            Case ".jpg", ".jpeg" : localMimeType = "image/jpeg"
                                            Case Else : localMimeType = "application/octet-stream"
                                        End Select
                                    End If

                                    localBase64 = System.Convert.ToBase64String(bytes)
                                    succeeded = True
                                    Exit Sub
                                End Using
                            End If
                        Finally
                            ' Ensure we drop COM references that can hold the clipboard data object
                            If contentStream IsNot Nothing Then contentStream.Dispose()
                            SafeReleaseCom(contentObj)
                        End Try
                    End Using
                End If
            Finally
                ' BinaryReader.Dispose closes fgStream; also release COM wrapper if any
                SafeReleaseCom(fgObj)
            End Try
        End If

        ' 2) File-drop (Explorer copy)
        If System.Windows.Forms.Clipboard.ContainsFileDropList() Then
            Dim files = System.Windows.Forms.Clipboard.GetFileDropList()
            If files.Count > 0 Then
                Dim path = files(0)
                Dim mresult = MimeHelper.GetFileMimeTypeAndBase64(path)
                localMimeType = mresult.MimeType.Trim()
                localBase64 = mresult.EncodedData.Trim()
                succeeded = True
                Exit Sub
            End If
        End If

        ' 3) Raw WAV stream
        If System.Windows.Forms.Clipboard.ContainsAudio() Then
            Using audioStream As System.IO.Stream = System.Windows.Forms.Clipboard.GetAudioStream()
                Using ms As New System.IO.MemoryStream()
                    audioStream.CopyTo(ms)
                    localBase64 = System.Convert.ToBase64String(ms.ToArray())
                    localMimeType = "audio/wav"
                    succeeded = True
                    Exit Sub
                End Using
            End Using
        End If

        ' 4) RTF
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Rtf) Then
            localMimeType = "application/rtf"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Rtf)))
            succeeded = True : Exit Sub
        End If

        ' 5) HTML
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Html) Then
            localMimeType = "text/html"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Html)))
            succeeded = True : Exit Sub
        End If

        ' 6) CSV
        If System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.CommaSeparatedValue) Then
            localMimeType = "text/csv"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.CommaSeparatedValue)))
            succeeded = True : Exit Sub
        End If

        ' 7) Plain text
        If System.Windows.Forms.Clipboard.ContainsText() Then
            localMimeType = "text/plain"
            localBase64 = System.Convert.ToBase64String(
                                System.Text.Encoding.UTF8.GetBytes(
                                    System.Windows.Forms.Clipboard.GetText()))
            succeeded = True : Exit Sub
        End If

        ' 8) Image (Bitmap → PNG)
        If System.Windows.Forms.Clipboard.ContainsImage() Then
            Using img As System.Drawing.Image = System.Windows.Forms.Clipboard.GetImage()
                Using ms As New System.IO.MemoryStream()
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    localMimeType = "image/png"
                    localBase64 = System.Convert.ToBase64String(ms.ToArray())
                    succeeded = True : Exit Sub
                End Using
            End Using
        End If

        ' 9) EMF → Bitmap → PNG
        If NativeClipboardX.OpenClipboard(IntPtr.Zero) Then
            Try
                If NativeClipboardX.IsClipboardFormatAvailable(NativeClipboardX.CF_ENHMETAFILE) Then
                    Dim src As IntPtr = NativeClipboardX.GetClipboardData(NativeClipboardX.CF_ENHMETAFILE)
                    If src <> IntPtr.Zero Then
                        Dim clone As IntPtr = NativeClipboardX.CopyEnhMetaFile(src, Nothing)
                        Try
                            Using emf As New System.Drawing.Imaging.Metafile(clone, False)
                                Using bmp As New System.Drawing.Bitmap(emf.Width, emf.Height)
                                    Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
                                        g.DrawImage(emf, 0, 0)
                                        Using out As New System.IO.MemoryStream()
                                            bmp.Save(out, System.Drawing.Imaging.ImageFormat.Png)
                                            localMimeType = "image/png"
                                            localBase64 = System.Convert.ToBase64String(out.ToArray())
                                            succeeded = True
                                        End Using
                                    End Using
                                End Using
                            End Using
                        Finally
                            ' Always free the duplicated handle
                            NativeClipboardX.DeleteEnhMetaFile(clone)
                        End Try
                        If succeeded Then Exit Sub
                    End If
                End If
            Finally
                NativeClipboardX.CloseClipboard()
            End Try
        End If

    Catch
        ' suppress all exceptions
    End Try
End Sub)

            t.SetApartmentState(System.Threading.ApartmentState.STA)
            t.Start()
            t.Join()

            If succeeded Then
                mimeType = localMimeType
                base64 = localBase64
            End If

            Return succeeded
        End Function

    End Module
End Namespace