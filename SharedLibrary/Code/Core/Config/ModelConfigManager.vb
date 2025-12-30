' Part of: Red Ink Shared Library by David Rosenthal
' 2025-12-06 Modelmanager implemented by Roman Kost

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Manages multiple LLM model configurations loaded from redink.ini.
    ''' - Detects numbered model variants (APIKey_1, Endpoint_1, Model_1, ...).
    ''' - Stores them in memory.
    ''' - Applies the selected model to the primary INI_* fields in the shared context.
    ''' - Tracks the currently selected model number.
    ''' 
    ''' This class only manipulates PRIMARY fields (INI_APIKey, INI_Model, ...).
    ''' It does NOT touch the secondary API slot (INI_*_2) used by SecondAPI / alternate models.
    ''' </summary>
    Public Class ModelConfigManager

        Private Const MultiModelPrefix As String = "MultiModel_"
        Private Const MaxModelNumber As Integer = 100
        ' Keep this aligned with SharedMethods.LoadConfig.INIValuesMissing to ensure both checks require the same core fields.
        Private Shared ReadOnly RequiredModelKeys As String() = {"Endpoint", "Model", "APICall", "Response"}

        ' Storage for all detected model configurations (per model number)
        Private Shared ReadOnly _availableModels As New Dictionary(Of Integer, Dictionary(Of String, String))()

        ' Currently selected model number (1-based), default 1
        Private Shared _currentModelNumber As Integer = 1

        ' Display names for each model (e.g. the Model_N value)
        Private Shared ReadOnly _modelDisplayNames As New Dictionary(Of Integer, String)()

        ''' <summary>
        ''' Detects and stores all model configurations from the parsed INI dictionary.
        ''' </summary>
        Public Shared Sub DetectAndStoreModels(configDict As Dictionary(Of String, String))
            _availableModels.Clear()
            _modelDisplayNames.Clear()

            ' Preferred format: MultiModel_<Key>_<Number>
            Dim perModel = ExtractPrefixedModels(configDict)

            ' Validate each detected model
            For Each n In perModel.Keys.OrderBy(Function(x) x)
                Dim mc = perModel(n)
                If ValidateModel(mc, $"#{n}") Then
                    _availableModels(n) = mc
                    _modelDisplayNames(n) = GetModelDisplayNameForConfig(mc, n)
                End If
            Next

            ' Backward compatibility: if no numbered models, extract a "primary" model as Model 1
            If _availableModels.Count = 0 Then
                Dim primary = ExtractPrimaryModelConfig(configDict)
                If ValidateModel(primary, "primary (no suffix)") Then
                    _availableModels(1) = primary
                    _modelDisplayNames(1) = GetModelDisplayNameForConfig(primary, 1)
                End If
            End If

            ' If previous current model is no longer available, fall back to 1
            If _availableModels.Count > 0 Then
                If Not _availableModels.ContainsKey(_currentModelNumber) Then
                    _currentModelNumber = 1
                End If
            Else
                _currentModelNumber = 1
            End If
        End Sub

        ''' <summary>
        ''' Returns the list of available model numbers.
        ''' </summary>
        Public Shared Function GetAvailableModels() As List(Of Integer)
            Return _availableModels.Keys.OrderBy(Function(n) n).ToList()
        End Function

        ''' <summary>
        ''' Returns a user-friendly display name for a model.
        ''' </summary>
        Public Shared Function GetModelDisplayName(modelNumber As Integer) As String
            If _modelDisplayNames.ContainsKey(modelNumber) Then
                Return _modelDisplayNames(modelNumber)
            End If
            Return $"Model {modelNumber}"
        End Function

        ''' <summary>
        ''' Returns the currently selected model number (1-based).
        ''' </summary>
        Public Shared Function GetCurrentModelNumber() As Integer
            Return _currentModelNumber
        End Function

        ''' <summary>
        ''' Selects a model by copying its configuration into the primary INI_* fields.
        ''' Returns False if the model number is not available or if an error occurs.
        ''' </summary>
        Public Shared Function SelectModel(context As ISharedContext, modelNumber As Integer) As Boolean
            If context Is Nothing Then Return False
            If Not _availableModels.ContainsKey(modelNumber) Then Return False

            Dim mc = _availableModels(modelNumber)

            Try
                CopyPropertiesToContext(context, mc)
                UpdateDecodedAPI(context)

                _currentModelNumber = modelNumber

                ' Persist selection to settings; ignore failures
                Try
                    My.Settings.SelectedModelNumber = modelNumber
                    My.Settings.Save()
                Catch
                End Try

                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Loads the saved model selection from settings.
        ''' Returns 1 if not found or invalid.
        ''' </summary>
        Public Shared Function LoadSavedModelNumber() As Integer
            Try
                Dim n = My.Settings.SelectedModelNumber
                If n > 0 Then Return n
            Catch
            End Try
            Return 1
        End Function

        ' -----------------------
        '  Internal helpers
        ' -----------------------

        Private Shared Function ExtractPrefixedModels(configDict As Dictionary(Of String, String)) As Dictionary(Of Integer, Dictionary(Of String, String))
            Dim perModel As New Dictionary(Of Integer, Dictionary(Of String, String))()

            For Each kvp In configDict
                Dim key = kvp.Key
                If Not key.StartsWith(MultiModelPrefix, StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim remainder = key.Substring(MultiModelPrefix.Length)
                Dim idx = remainder.LastIndexOf("_"c)
                If idx <= 0 OrElse idx = remainder.Length - 1 Then Continue For

                Dim baseKey = remainder.Substring(0, idx)
                Dim suffix = remainder.Substring(idx + 1)

                Dim n As Integer
                If Integer.TryParse(suffix, n) AndAlso n > 0 AndAlso n <= MaxModelNumber Then
                    If Not perModel.ContainsKey(n) Then
                        perModel(n) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    End If

                    perModel(n)(baseKey) = kvp.Value
                End If
            Next

            Return perModel
        End Function
        Private Shared Function ExtractPrimaryModelConfig(configDict As Dictionary(Of String, String)) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each kvp In configDict
                Dim key = kvp.Key
                ' skip keys that clearly belong to numbered or secondary models
                Dim idx = key.LastIndexOf("_"c)
                Dim numericSuffix As Integer
                If idx > 0 AndAlso Integer.TryParse(key.Substring(idx + 1), numericSuffix) AndAlso numericSuffix > 0 AndAlso numericSuffix <= MaxModelNumber Then
                    ' belongs to model N, handled elsewhere
                    Continue For
                End If
                result(key) = kvp.Value
            Next

            Return result
        End Function

        Private Shared Function ValidateModel(config As Dictionary(Of String, String), modelLabel As String) As Boolean
            If config Is Nothing OrElse config.Count = 0 Then
                Dim emptyMessage = BuildMissingKeyMessage(modelLabel, Nothing, True, config)
                Debug.WriteLine($"[ModelConfigManager] {emptyMessage}")
                MessageBox.Show(emptyMessage,
                                "Model configuration invalid",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return False
            End If

            Dim missing As New List(Of String)()

            For Each key In RequiredModelKeys
                If Not config.ContainsKey(key) OrElse String.IsNullOrWhiteSpace(config(key)) Then
                    missing.Add(key)
                End If
            Next

            If missing.Count > 0 Then
                Dim detailedMsg = BuildMissingKeyMessage(modelLabel, missing, False, config)
                Debug.WriteLine($"[ModelConfigManager] {detailedMsg}")
                MessageBox.Show(detailedMsg,
                                "Model configuration invalid",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return False
            End If

            Return True
        End Function

        Private Shared Function BuildMissingKeyMessage(modelLabel As String,
                                                       missingKeys As List(Of String),
                                                       noEntries As Boolean,
                                                       modelConfig As Dictionary(Of String, String)) As String
            Dim prefix As String
            Dim fixHints As New List(Of String)()

            Dim labelNumber As Integer
            Dim hasNumericLabel As Boolean = False

            If Not String.IsNullOrWhiteSpace(modelLabel) AndAlso modelLabel.StartsWith("#", StringComparison.Ordinal) Then
                Dim suffix = modelLabel.TrimStart("#"c)
                If Integer.TryParse(suffix, labelNumber) Then
                    hasNumericLabel = True
                End If
            End If

            If hasNumericLabel Then
                Dim identifierSuffix As String = ""
                prefix = $"Model #{labelNumber}"
                If missingKeys IsNot Nothing Then
                    For Each key In missingKeys
                        fixHints.Add($"MultiModel_{key}_{labelNumber}")
                    Next
                Else
                    fixHints.Add($"MultiModel_<Key>_{labelNumber}")
                End If

                Dim itemNameExists As Boolean = modelConfig IsNot Nothing AndAlso modelConfig.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("ItemName"))
                Dim modelNameExists As Boolean = modelConfig IsNot Nothing AndAlso modelConfig.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("Model"))

                If itemNameExists Then
                    identifierSuffix = $"MultiModel_ItemName_{labelNumber} = '{modelConfig("ItemName")}'"
                ElseIf modelNameExists Then
                    identifierSuffix = $"MultiModel_Model_{labelNumber} = '{modelConfig("Model")}'"
                Else
                    identifierSuffix = $"MultiModel suffix #{labelNumber}"
                End If

                prefix = $"{prefix} ({identifierSuffix})"
            Else
                prefix = $"Model {modelLabel}"
                If missingKeys IsNot Nothing Then
                    fixHints.AddRange(missingKeys)
                Else
                    fixHints.Add("base INI keys (e.g., Endpoint, Model, APICall, Response)")
                End If

                If modelConfig IsNot Nothing Then
                    Dim primaryName As String = ""
                    If modelConfig.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("ItemName")) Then
                        primaryName = $"ItemName = '{modelConfig("ItemName")}'"
                    ElseIf modelConfig.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(modelConfig("Model")) Then
                        primaryName = $"Model = '{modelConfig("Model")}'"
                    End If

                    If Not String.IsNullOrWhiteSpace(primaryName) Then
                        prefix = $"{prefix} ({primaryName})"
                    End If
                End If
            End If

            Dim baseMessage As String
            If noEntries Then
                baseMessage = $"{prefix} did not contain any configuration entries."
            Else
                baseMessage = $"{prefix} is missing required fields: {String.Join(", ", missingKeys)}."
            End If

            Dim hintText = $"Please set the following INI keys: {String.Join(", ", fixHints.Distinct())}"
            Return $"{baseMessage}{Environment.NewLine}{hintText}"
        End Function

        Private Shared Function GetModelDisplayNameForConfig(config As Dictionary(Of String, String), n As Integer) As String
            If config.ContainsKey("ItemName") AndAlso Not String.IsNullOrWhiteSpace(config("ItemName")) Then
                Return config("ItemName")
            End If
            If config.ContainsKey("Model") AndAlso Not String.IsNullOrWhiteSpace(config("Model")) Then
                Return config("Model")
            End If
            Return $"Model {n}"
        End Function

        Private Shared Sub CopyPropertiesToContext(context As ISharedContext,
                                                   model As Dictionary(Of String, String))
            ' Strings
            context.INI_APIKey = GetValue(model, "APIKey", "")
            context.INI_APIKeyBack = context.INI_APIKey
            context.INI_APIKeyPrefix = GetValue(model, "APIKeyPrefix", "")
            context.INI_Endpoint = GetValue(model, "Endpoint", "")
            context.INI_Model = GetValue(model, "Model", "")
            context.INI_Temperature = GetValue(model, "Temperature", "")
            context.INI_HeaderA = GetValue(model, "HeaderA", "")
            context.INI_HeaderB = GetValue(model, "HeaderB", "")
            context.INI_Response = GetValue(model, "Response", "")
            context.INI_APICall = GetValue(model, "APICall", "")
            context.INI_APICall_Object = GetValue(model, "APICall_Object", "")
            context.INI_Anon = GetValue(model, "Anon", "")
            context.INI_TokenCount = GetValue(model, "TokenCount", "")
            context.INI_PreCorrection = GetValue(model, "PreCorrection", "")
            context.INI_PostCorrection = GetValue(model, "PostCorrection", "")
            context.INI_OAuth2ClientMail = GetValue(model, "OAuth2ClientMail", "")
            context.INI_OAuth2Scopes = GetValue(model, "OAuth2Scopes", "")
            context.INI_OAuth2Endpoint = GetValue(model, "OAuth2Endpoint", "")

            ' Numerics
            context.INI_Timeout = GetLongValue(model, "Timeout", 0)
            context.INI_MaxOutputToken = GetIntValue(model, "MaxOutputToken", 0)
            context.INI_OAuth2ATExpiry = GetLongValue(model, "OAuth2ATExpiry", 3600)

            ' Booleans
            context.INI_APIEncrypted = GetBoolValue(model, "APIKeyEncrypted", False)
            context.INI_DoubleS = GetBoolValue(model, "DoubleS", False)
            context.INI_Clean = GetBoolValue(model, "Clean", False)
            context.INI_NoDash = GetBoolValue(model, "NoEmDash", False)
            context.INI_OAuth2 = GetBoolValue(model, "OAuth2", False)
        End Sub

        Private Shared Sub UpdateDecodedAPI(context As ISharedContext)
            ' Mirrors logic from InitializeConfig for primary API
            If context.INI_OAuth2 Then
                context.INI_APIKey = Trim(SharedMethods.RealAPIKey(context.INI_APIKey, False, True, context)).Replace(vbLf, "").Replace(vbCr, "")
                ' If something goes wrong, RealAPIKey will already have shown an error via SharedMethods
            Else
                context.DecodedAPI = SharedMethods.RealAPIKey(context.INI_APIKey, False, False, context)
            End If
        End Sub

        Private Shared Function GetValue(config As Dictionary(Of String, String),
                                         key As String,
                                         defaultValue As String) As String
            If config.ContainsKey(key) Then
                Return config(key)
            End If
            Return defaultValue
        End Function

        Private Shared Function GetLongValue(config As Dictionary(Of String, String),
                                             key As String,
                                             defaultValue As Long) As Long
            If config.ContainsKey(key) Then
                Dim v As Long
                If Long.TryParse(config(key), v) Then
                    Return v
                End If
            End If
            Return defaultValue
        End Function

        Private Shared Function GetIntValue(config As Dictionary(Of String, String),
                                            key As String,
                                            defaultValue As Integer) As Integer
            If config.ContainsKey(key) Then
                Dim v As Integer
                If Integer.TryParse(config(key), v) Then
                    Return v
                End If
            End If
            Return defaultValue
        End Function

        Private Shared Function GetBoolValue(config As Dictionary(Of String, String),
                                             key As String,
                                             defaultValue As Boolean) As Boolean
            If config.ContainsKey(key) Then
                Dim raw = config(key).Trim().ToLowerInvariant()
                Return raw = "yes" OrElse raw = "true" OrElse raw = "ja" OrElse raw = "wahr" OrElse raw = "1"
            End If
            Return defaultValue
        End Function

    End Class

End Namespace


