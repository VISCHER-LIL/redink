' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices

Namespace SharedLibrary

    Public Module MimeHelper

        ' P/Invoke to urlmon.dll for MIME sniffing
        <DllImport("urlmon.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Function FindMimeFromData(
    ByVal pBC As IntPtr,
    <MarshalAs(UnmanagedType.LPWStr)> ByVal pwzUrl As String,
    <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.I1, SizeParamIndex:=3)> ByVal pBuffer As Byte(),
    ByVal cbSize As UInteger,
    <MarshalAs(UnmanagedType.LPWStr)> ByVal pwzMimeProposed As String,
    ByVal dwMimeFlags As UInteger,
    ByRef ppwzMimeOut As IntPtr,
    ByVal dwReserved As UInteger
) As Integer
        End Function

        Public Function GetFileMimeTypeAndBase64(
    ByVal filePath As String
) As (MimeType As String, EncodedData As String)
            Try
                ' 1) sniff the MIME type
                Dim mime As String = GetMimeType(filePath)

                ' 2) read and Base64-encode
                Dim bytes As Byte() = File.ReadAllBytes(filePath)
                Dim b64 As String = System.Convert.ToBase64String(bytes)

                Return (mime, b64)
            Catch ex As System.Exception
                Throw New System.Exception("Error determining MIME type or encoding data: " & ex.Message, ex)
            End Try
        End Function

        ' Uses FindMimeFromData to inspect the first 256 bytes of the file and return a MIME type.
        Private Function GetMimeType(ByVal filePath As String) As String
            Dim buffer(255) As Byte
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                Dim read As Integer = fs.Read(buffer, 0, buffer.Length)
                If read = 0 Then
                    Throw New System.Exception("Unable to read from file: " & filePath)
                End If
            End Using

            Dim mimePtr As IntPtr = IntPtr.Zero
            Dim hr As Integer = FindMimeFromData(
        IntPtr.Zero,
        filePath,
        buffer,
        CUInt(buffer.Length),
        Nothing,
        0,
        mimePtr,
        0
    )

            If hr <> 0 Then
                Throw New System.Exception($"FindMimeFromData failed with HRESULT 0x{hr:X8}")
            End If

            Dim mime As String = Marshal.PtrToStringUni(mimePtr)
            Marshal.FreeCoTaskMem(mimePtr)

            Return mime
        End Function

    End Module

End Namespace