' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.IO

Namespace SharedLibrary

    Partial Public Class SharedMethods


        ' Creates a ModelConfig object from a dictionary of key/value pairs.
        Public Shared Function CreateModelConfigFromDict(ByVal configDict As Dictionary(Of String, String), context As ISharedContext, Description As String) As ModelConfig
            Dim mc As New ModelConfig()
            Try
                mc.APIKey = If(configDict.ContainsKey("APIKey"), configDict("APIKey"), "")
                mc.Endpoint = If(configDict.ContainsKey("Endpoint"), configDict("Endpoint"), "")
                mc.HeaderA = If(configDict.ContainsKey("HeaderA"), configDict("HeaderA"), "")
                mc.HeaderB = If(configDict.ContainsKey("HeaderB"), configDict("HeaderB"), "")
                mc.Response = If(configDict.ContainsKey("Response"), configDict("Response"), "")
                mc.APICall = If(configDict.ContainsKey("APICall"), configDict("APICall"), "")
                mc.APICall_Object = If(configDict.ContainsKey("APICall_Object"), configDict("APICall_Object"), "")
                mc.Timeout = If(configDict.ContainsKey("Timeout"), CLng(configDict("Timeout")), 0)
                mc.MaxOutputToken = If(configDict.ContainsKey("MaxOutputToken_2"), CInt(configDict("MaxOutputToken")), 0)
                mc.Temperature = If(configDict.ContainsKey("Temperature"), configDict("Temperature"), "")
                mc.Model = If(configDict.ContainsKey("Model"), configDict("Model"), "")
                mc.APIEncrypted = ParseBoolean(configDict, "APIKeyEncrypted")
                mc.APIKeyPrefix = If(configDict.ContainsKey("APIKeyPrefix"), configDict("APIKeyPrefix"), "")
                mc.OAuth2 = ParseBoolean(configDict, "OAuth2")
                mc.OAuth2ClientMail = If(configDict.ContainsKey("OAuth2ClientMail"), configDict("OAuth2ClientMail"), "")
                mc.OAuth2Scopes = If(configDict.ContainsKey("OAuth2Scopes"), configDict("OAuth2Scopes"), "")
                mc.OAuth2Endpoint = If(configDict.ContainsKey("OAuth2Endpoint"), configDict("OAuth2Endpoint"), "")
                mc.OAuth2ATExpiry = If(configDict.ContainsKey("OAuth2ATExpiry"), CLng(configDict("OAuth2ATExpiry")), 3600)
                mc.Parameter1 = If(configDict.ContainsKey("Parameter1"), configDict("Parameter1"), "")
                mc.Parameter2 = If(configDict.ContainsKey("Parameter2"), configDict("Parameter2"), "")
                mc.Parameter3 = If(configDict.ContainsKey("Parameter3"), configDict("Parameter3"), "")
                mc.Parameter4 = If(configDict.ContainsKey("Parameter4"), configDict("Parameter4"), "")
                mc.MergePrompt = If(configDict.ContainsKey("MergePrompt"), configDict("MergePrompt"), context.SP_MergePrompt)
                mc.QueryPrompt = If(configDict.ContainsKey("QueryPrompt"), configDict("QueryPrompt"), "")
                mc.ModelDescription = Description

                mc.APIKeyBack = mc.APIKey

                ' Additional configurations for OAuth2
                mc.TokenExpiry = Microsoft.VisualBasic.DateAndTime.DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateTime.Now)
                mc.DecodedAPI = ""

                ' Check and decrypt API keys
                If mc.OAuth2 Then
                    mc.APIKey = Trim(Replace(RealAPIKeyMC(mc.APIKey, True, mc, context), "\n", ""))
                Else
                    mc.DecodedAPI = RealAPIKeyMC(mc.APIKey, False, mc, context)
                End If

            Catch ex As System.Exception
                MessageBox.Show("Error in CreateModelConfigFromDict: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return mc
        End Function



        ' Extracts the current configuration from the shared context using the same style.
        Public Shared Function GetCurrentConfig(ByVal context As ISharedContext) As ModelConfig
            Dim mc As New ModelConfig()
            Try
                ' Here we simulate reading from a config dictionary by using the context values.
                mc.APIKey = If(String.IsNullOrEmpty(context.INI_APIKey_2), "", context.INI_APIKey_2)
                mc.APIKeyBack = If(String.IsNullOrEmpty(context.INI_APIKeyBack_2), "", context.INI_APIKeyBack_2)
                mc.Endpoint = If(String.IsNullOrEmpty(context.INI_Endpoint_2), "", context.INI_Endpoint_2)
                mc.HeaderA = If(String.IsNullOrEmpty(context.INI_HeaderA_2), "", context.INI_HeaderA_2)
                mc.HeaderB = If(String.IsNullOrEmpty(context.INI_HeaderB_2), "", context.INI_HeaderB_2)
                mc.Response = If(String.IsNullOrEmpty(context.INI_Response_2), "", context.INI_Response_2)
                mc.Anon = If(String.IsNullOrEmpty(context.INI_Anon_2), "", context.INI_Anon_2)
                mc.TokenCount = If(String.IsNullOrEmpty(context.INI_TokenCount_2), "", context.INI_TokenCount_2)
                mc.APICall = If(String.IsNullOrEmpty(context.INI_APICall_2), "", context.INI_APICall_2)
                mc.APICall_Object = If(String.IsNullOrEmpty(context.INI_APICall_Object_2), "", context.INI_APICall_Object_2)
                mc.Timeout = context.INI_Timeout_2
                mc.MaxOutputToken = context.INI_MaxOutputToken_2
                mc.Temperature = If(String.IsNullOrEmpty(context.INI_Temperature_2), "", context.INI_Temperature_2)
                mc.Model = If(String.IsNullOrEmpty(context.INI_Model_2), "", context.INI_Model_2)
                mc.APIEncrypted = context.INI_APIEncrypted_2
                mc.APIKeyPrefix = If(String.IsNullOrEmpty(context.INI_APIKeyPrefix_2), "", context.INI_APIKeyPrefix_2)
                mc.OAuth2 = context.INI_OAuth2_2
                mc.OAuth2ClientMail = If(String.IsNullOrEmpty(context.INI_OAuth2ClientMail_2), "", context.INI_OAuth2ClientMail_2)
                mc.OAuth2Scopes = If(String.IsNullOrEmpty(context.INI_OAuth2Scopes_2), "", context.INI_OAuth2Scopes_2)
                mc.OAuth2Endpoint = If(String.IsNullOrEmpty(context.INI_OAuth2Endpoint_2), "", context.INI_OAuth2Endpoint_2)
                mc.OAuth2ATExpiry = context.INI_OAuth2ATExpiry_2
                mc.MergePrompt = If(String.IsNullOrEmpty(context.SP_MergePrompt), "", context.SP_MergePrompt)
                mc.DecodedAPI = context.DecodedAPI_2
                mc.TokenExpiry = context.TokenExpiry_2

            Catch ex As System.Exception
                MessageBox.Show("Error in GetCurrentConfig: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return mc
        End Function

        ' Applies the given ModelConfig to the shared context using the assignment style.
        Public Shared Sub ApplyModelConfig(ByVal context As ISharedContext, ByVal config As ModelConfig, Optional ByRef ErrorFlag As Boolean = False)
            Try
                context.INI_APIKey_2 = If(Not String.IsNullOrEmpty(config.APIKey), config.APIKey, "")
                context.INI_APIKeyBack_2 = If(Not String.IsNullOrEmpty(config.APIKeyBack), config.APIKeyBack, "")
                context.INI_Endpoint_2 = If(Not String.IsNullOrEmpty(config.Endpoint), config.Endpoint, "")
                context.INI_HeaderA_2 = If(Not String.IsNullOrEmpty(config.HeaderA), config.HeaderA, "")
                context.INI_HeaderB_2 = If(Not String.IsNullOrEmpty(config.HeaderB), config.HeaderB, "")
                context.INI_Response_2 = If(Not String.IsNullOrEmpty(config.Response), config.Response, "")
                context.INI_Anon_2 = If(Not String.IsNullOrEmpty(config.Anon), config.Anon, "")
                context.INI_TokenCount_2 = If(Not String.IsNullOrEmpty(config.TokenCount), config.TokenCount, "")
                context.INI_APICall_2 = If(Not String.IsNullOrEmpty(config.APICall), config.APICall, "")
                context.INI_APICall_Object_2 = If(Not String.IsNullOrEmpty(config.APICall_Object), config.APICall_Object, "")
                context.INI_Timeout_2 = If(config.Timeout <> 0, config.Timeout, 0)
                context.INI_MaxOutputToken_2 = If(config.MaxOutputToken <> 0, config.MaxOutputToken, 0)
                context.INI_Temperature_2 = If(Not String.IsNullOrEmpty(config.Temperature), config.Temperature, "")
                context.INI_Model_2 = If(Not String.IsNullOrEmpty(config.Model), config.Model, "")
                context.INI_APIEncrypted_2 = config.APIEncrypted
                context.INI_APIKeyPrefix_2 = If(Not String.IsNullOrEmpty(config.APIKeyPrefix), config.APIKeyPrefix, "")
                context.INI_OAuth2_2 = config.OAuth2
                context.INI_OAuth2ClientMail_2 = If(Not String.IsNullOrEmpty(config.OAuth2ClientMail), config.OAuth2ClientMail, "")
                context.INI_OAuth2Scopes_2 = If(Not String.IsNullOrEmpty(config.OAuth2Scopes), config.OAuth2Scopes, "")
                context.INI_OAuth2Endpoint_2 = If(Not String.IsNullOrEmpty(config.OAuth2Endpoint), config.OAuth2Endpoint, "")
                context.INI_OAuth2ATExpiry_2 = If(config.OAuth2ATExpiry <> 0, config.OAuth2ATExpiry, 3600)
                context.DecodedAPI_2 = config.DecodedAPI
                context.TokenExpiry_2 = config.TokenExpiry
                context.INI_Model_Parameter1 = If(Not String.IsNullOrEmpty(config.Parameter1), config.Parameter1, "")
                context.INI_Model_Parameter2 = If(Not String.IsNullOrEmpty(config.Parameter2), config.Parameter2, "")
                context.INI_Model_Parameter3 = If(Not String.IsNullOrEmpty(config.Parameter3), config.Parameter3, "")
                context.INI_Model_Parameter4 = If(Not String.IsNullOrEmpty(config.Parameter4), config.Parameter4, "")
                context.SP_MergePrompt = If(Not String.IsNullOrEmpty(config.MergePrompt), config.MergePrompt, "")
                SP_QueryPrompt = If(Not String.IsNullOrEmpty(config.QueryPrompt), config.QueryPrompt, "")

                ErrorFlag = False

            Catch ex As System.Exception
                If Not ErrorFlag Then
                    MessageBox.Show("Error in ApplyModelConfig: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ErrorFlag = True
            End Try
        End Sub

        ' Restores the default configuration (passed in as originalConfig).
        Public Shared Sub RestoreDefaults(ByVal context As ISharedContext, ByVal originalConfig As ModelConfig)
            ApplyModelConfig(context, originalConfig)
        End Sub


        ' Loads alternative model configurations from an INI file.
        Public Shared Function LoadAlternativeModels(ByVal iniFilePath As String, context As ISharedContext) As List(Of ModelConfig)
            Dim models As New List(Of ModelConfig)()
            Try
                If Not File.Exists(iniFilePath) Then
                    ShowCustomMessageBox($"INI file for alternative models not found (update {AN2}.ini): " & iniFilePath)
                    Return models
                End If

                Dim currentDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim Description As String = ""
                For Each XLine In File.ReadAllLines(iniFilePath)
                    Dim trimmedLine As String = XLine.Trim()
                    ' Skip empty lines and comments.
                    If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                        Continue For
                    End If

                    ' Section header (e.g., [Model1]) indicates a new model.
                    If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                        If currentDict.Count > 0 Then
                            models.Add(CreateModelConfigFromDict(currentDict, context, Description))
                            currentDict.Clear()
                        End If
                        Description = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim()
                        Continue For
                    End If

                    ' Parse key=value lines.
                    Dim tokens() As String = trimmedLine.Split(New Char() {"="c}, 2)
                    If tokens.Length = 2 Then
                        Dim key As String = tokens(0).Trim()
                        Dim value As String = tokens(1).Trim()
                        ' Store the key/value pair.
                        If Not currentDict.ContainsKey(key) Then
                            currentDict.Add(key, value)
                        Else
                            currentDict(key) = value
                        End If
                    End If
                Next
                ' Add the last model if any.
                If currentDict.Count > 0 Then
                    models.Add(CreateModelConfigFromDict(currentDict, context, Description))
                End If
            Catch ex As System.Exception
                ShowCustomMessageBox($"Error reading INI file for alternative models ({iniFilePath}): " & ex.Message)
            End Try
            Return models
        End Function


        Public Shared originalConfig As ModelConfig
        Public Shared OptionChecked As Boolean = False
        Public Shared originalConfigLoaded As Boolean = False
        Public Shared SelectedAlternateModels As List(Of ModelConfig)
        Public Shared LastAlternateModel As String = ""

        ' Displays the model selection form and applies the chosen configuration.
        Public Shared Function ShowModelSelection(ByVal context As ISharedContext, iniFilePath As String, Optional Title As String = "Freestyle", Optional Listtype As String = "Select the model you want to use:", Optional OptionText As String = "Reset to default model after use", Optional UseCase As Integer = 1) As Boolean
            Try
                ' Back up the current (default) configuration.

                originalConfig = GetCurrentConfig(context)
                originalConfigLoaded = True

                Dim selector As New ModelSelectorForm(iniFilePath, context, Title, Listtype, OptionText, UseCase)
                If selector.ShowDialog() = DialogResult.OK Then
                    If selector.UseDefault AndAlso UseCase = 1 Then
                        RestoreDefaults(context, originalConfig)
                    ElseIf selector.SelectedModel IsNot Nothing Then
                        ApplyModelConfig(context, selector.SelectedModel)
                    End If

                    If selector.SelectedModel IsNot Nothing Then
                        Dim m As ModelConfig = selector.SelectedModel
                        LastAlternateModel = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                    End If

                    Return True
                Else
                    Return False
                End If
            Catch ex As System.Exception
                MessageBox.Show("Error in ShowModelSelection: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function


        Public Shared Function ShowMultipleModelSelection(context As ISharedContext,
                                                  modelPath As String) As Boolean
            Try
                Dim iniPath As String = ExpandEnvironmentVariables(modelPath)
                If String.IsNullOrWhiteSpace(iniPath) OrElse Not System.IO.File.Exists(iniPath) Then
                    System.Windows.Forms.MessageBox.Show("The configured alternate model path does not exist.", AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                    Return False
                End If

                Dim alternativeModels As System.Collections.Generic.List(Of ModelConfig) = LoadAlternativeModels(iniPath, context)
                If alternativeModels Is Nothing OrElse alternativeModels.Count = 0 Then
                    System.Windows.Forms.MessageBox.Show("No alternate model configurations found in the specified file.", AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                    Return False
                End If

                Using form As New MultiModelSelectorForm(alternativeModels, LastAlternateModel, AN & " - Select Alternate Models", True)
                    If form.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                        Return False
                    End If

                    SelectedAlternateModels = form.SelectedModels
                    If SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 Then
                        Return False
                    End If

                    Return True
                End Using
            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show("Error during multi-model selection: " & ex.Message, AN, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Return False
            End Try
        End Function



        ' Retrieves and applies the first model whose section contains a flag parameter named Task
        ' with a truthy value (True/Yes/Wahr/Ja/1). Returns True if applied; False if none found.
        Public Shared Function GetSpecialTaskModel(ByVal context As ISharedContext,
                                               ByVal iniFilePath As String,
                                               ByVal Task As String,
                                               Optional ByVal UseCase As Integer = 1) As Boolean
            If String.IsNullOrWhiteSpace(Task) Then Return False
            Try
                If Not File.Exists(iniFilePath) Then
                    ShowCustomMessageBox($"INI file for alternative models not found (update {AN2}.ini): " & iniFilePath)
                    Return False
                End If

                ' Backup current (default) config like ShowModelSelection
                originalConfigLoaded = False
                originalConfig = GetCurrentConfig(context)
                originalConfigLoaded = True

                Dim normalizedTask As String = Task.Trim()
                Dim truthy = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "true", "yes", "wahr", "ja", "on"
                }

                Dim currentDict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim description As String = ""

                Dim applyIfMatch As Func(Of Boolean) =
                    Function()
                        If currentDict.Count = 0 Then Return False
                        If currentDict.ContainsKey(normalizedTask) Then
                            Dim raw As String = currentDict(normalizedTask)
                            If raw Is Nothing Then raw = ""
                            ' Strip inline comments ; or # (common INI patterns)
                            Dim scIdx = raw.IndexOf(";"c)
                            If scIdx >= 0 Then raw = raw.Substring(0, scIdx)
                            Dim hashIdx = raw.IndexOf("#"c)
                            If hashIdx >= 0 Then raw = raw.Substring(0, hashIdx)
                            raw = raw.Trim()
                            ' Remove surrounding quotes if present
                            If raw.Length >= 2 AndAlso ((raw.StartsWith("""") AndAlso raw.EndsWith("""")) OrElse (raw.StartsWith("'") AndAlso raw.EndsWith("'"))) Then
                                raw = raw.Substring(1, raw.Length - 2).Trim()
                            End If
                            Dim lowered = raw.ToLowerInvariant()
                            If truthy.Contains(lowered) OrElse lowered = "1" Then
                                Dim mc = CreateModelConfigFromDict(currentDict, context, description)
                                ApplyModelConfig(context, mc)
                                Return True
                            End If
                        End If
                        Return False
                    End Function

                For Each rawLine In File.ReadAllLines(iniFilePath)
                    Dim line = rawLine.Trim()
                    If line.Length = 0 OrElse line.StartsWith(";") OrElse line.StartsWith("#") Then
                        Continue For
                    End If

                    ' Section header
                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        If applyIfMatch() Then
                            Return True
                        End If
                        currentDict.Clear()
                        description = line.Substring(1, line.Length - 2).Trim()
                        Continue For
                    End If

                    ' key=value
                    Dim tokens = line.Split(New Char() {"="c}, 2)
                    If tokens.Length = 2 Then
                        Dim key = tokens(0).Trim()
                        Dim value = tokens(1).Trim()
                        currentDict(key) = value
                    End If
                Next

                ' Final section
                If applyIfMatch() Then
                    Return True
                End If

                Return False

            Catch ex As Exception
                MessageBox.Show("Error in GetSpecialModel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

    End Class
End Namespace
