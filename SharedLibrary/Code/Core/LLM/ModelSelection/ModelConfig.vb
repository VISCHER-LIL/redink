' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ModelConfig.vb
' Purpose: Defines a mutable data container that holds the currently selected LLM/model configuration
'          and related runtime state used by the model selection and invocation infrastructure.
'
' Architecture:
'  - Configuration Carrier: Instances of `ModelConfig` are populated from INI/config dictionaries (e.g. via
'    `SharedMethods.CreateModelConfigFromDict`) and can be applied back to an `ISharedContext` (e.g. via
'    `SharedMethods.ApplyModelConfig`).
'  - Credentials: Stores API key material and related flags (e.g., encryption/prefix handling) and limited
'    OAuth2 configuration values required by callers.
'  - Runtime/Diagnostics State: Stores request/response related strings and token expiry information that
'    callers may fill during runtime.
'  - Cloning: `Clone()` returns a shallow copy via `MemberwiseClone` for snapshotting the current config.
' =============================================================================

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