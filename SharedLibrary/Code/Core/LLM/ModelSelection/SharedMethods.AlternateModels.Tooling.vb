' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.AlternateModels.Tooling.vb
' Purpose: Provides tooling-related helpers for alternate model configurations.
'          This file is a `Partial` extension of `SharedMethods` and groups helper
'          logic that determines whether a given `ModelConfig` is tool-capable and
'          how such models should be represented in UI strings.
'
' Architecture:
'  - Tool Capability Detection: `ModelSupportsTooling` returns True when the model configuration
'    indicates tool support through `ModelConfig.Tool` or a non-empty `ModelConfig.APICall_ToolInstructions`.
'  - UI Display Normalization: `GetModelDisplayWithToolingSuffix` returns the display label used in UI,
'    based on `ModelConfig.ModelDescription` (fallback: `ModelConfig.Model`) and appends `ToolingSuffix`
'    when tool support is detected.
'
' Notes:
'  - `ToolingSuffix` is expected to be defined elsewhere in the `SharedMethods` partial type
'    (typically in a constants/definitions file). This file only applies the suffix.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.IO
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Partial Public Class SharedMethods


        ''' <summary>
        ''' Determines if a model supports tooling based on its configuration.
        ''' </summary>
        ''' <param name="config">ModelConfig to check.</param>
        ''' <returns>True if the model supports tooling.</returns>
        Public Shared Function ModelSupportsTooling(config As ModelConfig) As Boolean
            If config Is Nothing Then Return False
            Return config.Tool OrElse
                   Not String.IsNullOrWhiteSpace(config.APICall_ToolInstructions)
        End Function

        ''' <summary>
        ''' Gets the display description for a model, including tooling suffix if applicable.
        ''' </summary>
        ''' <param name="config">ModelConfig to get description for.</param>
        ''' <returns>Display description with appropriate suffix.</returns>
        Public Shared Function GetModelDisplayWithToolingSuffix(config As ModelConfig) As String
            Dim baseDesc = If(Not String.IsNullOrWhiteSpace(config.ModelDescription),
                              config.ModelDescription, config.Model)

            If ModelSupportsTooling(config) Then
                If Not baseDesc.EndsWith(ToolingSuffix) Then
                    baseDesc &= ToolingSuffix
                End If
            End If

            Return baseDesc
        End Function

    End Class

End Namespace