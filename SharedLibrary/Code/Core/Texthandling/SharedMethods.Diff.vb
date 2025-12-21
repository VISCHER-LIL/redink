' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Partial Public Class SharedMethods

        Public Class Diff
            Public Enum Operation
                Equal
                Insert
                Delete
            End Enum

            Public Property Op As Operation
            Public Property Text As String

            Public Sub New(op As Operation, text As String)
                Me.Op = op
                Me.Text = text
            End Sub
        End Class

    End Class

End Namespace