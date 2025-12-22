' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Namespace SharedLibrary

    Friend Module NativeClipboardX
        Friend Const CF_ENHMETAFILE As UInteger = 14

        <DllImport("user32.dll", SetLastError:=True)>
        Friend Function OpenClipboard(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Friend Function CloseClipboard() As Boolean
        End Function

        <DllImport("user32.dll")>
        Friend Function IsClipboardFormatAvailable(fmt As UInteger) As Boolean
        End Function

        <DllImport("user32.dll")>
        Friend Function GetClipboardData(fmt As UInteger) As IntPtr
        End Function

        <DllImport("gdi32.dll")>
        Friend Function CopyEnhMetaFile(hEmfSrc As IntPtr,
                                    lpszFile As String) As IntPtr
        End Function

        <DllImport("gdi32.dll")>
        Friend Function DeleteEnhMetaFile(hemf As IntPtr) As Boolean
        End Function
    End Module

End Namespace