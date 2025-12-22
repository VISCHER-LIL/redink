' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class ModelConfig
        Public Property APIKey As String
        Public Property APIKeyBack As String
        Public Property Temperature As String
        Public Property Timeout As Long
        Public Property MaxOutputToken As Integer
        Public Property Model As String
        Public Property Endpoint As String
        Public Property HeaderA As String
        Public Property HeaderB As String
        Public Property APICall As String
        Public Property APICall_Object As String
        Public Property Response As String
        Public Property Anon As String
        Public Property TokenCount As String
        Public Property APIEncrypted As Boolean
        Public Property APIKeyPrefix As String
        Public Property OAuth2 As Boolean
        Public Property OAuth2ClientMail As String
        Public Property OAuth2Scopes As String
        Public Property OAuth2Endpoint As String
        Public Property OAuth2ATExpiry As Long
        Public Property ModelDescription As String
        Public Property DecodedAPI As String
        Public Property TokenExpiry As DateTime
        Public Property Parameter1 As String
        Public Property Parameter2 As String
        Public Property Parameter3 As String
        Public Property Parameter4 As String
        Public Property MergePrompt As String
        Public Property QueryPrompt As String

        Public Function Clone() As ModelConfig
            Return DirectCast(Me.MemberwiseClone(), ModelConfig)
        End Function
    End Class


End Namespace