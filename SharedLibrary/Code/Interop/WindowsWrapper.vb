' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WindowsWrapper.vb
' Purpose: Provides a lightweight wrapper to convert native window handles (HWND)
'          into managed IWin32Window instances for use with Windows Forms dialogs
'          and other APIs that require a parent window reference.
'
' Public Types:
'   WindowWrapper - Implements IWin32Window to bridge native/managed window ownership.
'
' Key Dependencies:
'   System.Windows.Forms.IWin32Window - .NET Framework interface for window handles.
'
' Thread Safety:
'   Immutable after construction; safe for concurrent read access.
'   The wrapped handle itself must remain valid for the lifetime of this instance.
'
' Performance Notes:
'   Zero-overhead wrapper; single IntPtr field with no allocations beyond the object itself.
'
' Usage Pattern:
'   Typical use case is wrapping Office application window handles (Word/Excel/Outlook)
'   to parent modal dialogs correctly:
'     Dim owner As New WindowWrapper(New IntPtr(Application.Hwnd))
'     MessageBox.Show(owner, "Message", "Title")
'
' External Libraries:
'   None (Framework-only).
'
' Change Log:
'   (Add entries as modifications are made)
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Windows.Forms

Namespace SharedLibrary

    ''' <summary>
    ''' Wraps a native window handle (HWND) to provide IWin32Window implementation
    ''' for use with managed Windows Forms APIs that require parent window references.
    ''' </summary>
    ''' <remarks>
    ''' This class is particularly useful when hosting dialogs from Office add-ins,
    ''' where the Office application window handle must be wrapped to ensure proper
    ''' modal dialog parenting and Z-order behavior.
    ''' </remarks>
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        ''' <summary>
        ''' Initializes a new instance of the WindowWrapper class with the specified window handle.
        ''' </summary>
        ''' <param name="handle">The native window handle (HWND) to wrap.</param>
        ''' <remarks>
        ''' The caller is responsible for ensuring the handle remains valid for the lifetime
        ''' of this wrapper instance. Invalid or destroyed window handles will cause undefined
        ''' behavior when used with Windows Forms APIs.
        ''' </remarks>
        Public Sub New(handle As IntPtr)
            _hwnd = handle
        End Sub

        ''' <summary>
        ''' Gets the window handle that this wrapper represents.
        ''' </summary>
        ''' <value>The native window handle (HWND).</value>
        Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

End Namespace