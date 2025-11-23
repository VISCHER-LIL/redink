' =============================================================================
' File: InitialConfig.vb
' Part of: Red Ink Shared Library
' Purpose: First-run installation wizard form for configuring Red Ink LLM API access.
'          Provides provider-specific templates (OpenAI, Azure, Google Vertex, etc.),
'          validates user input, supports remote configuration updates, and generates
'          INI files for Word/Excel/Outlook add-ins.
'
' Copyright: David Rosenthal, david.rosenthal@vischer.com
' License: May only be used with an appropriate license (see redink.ai)
'
' Architecture Overview:
' ----------------------
' InitialConfig is a modal wizard dialog that appears when no redink.ini file exists.
' It uses a provider-centric design where each LLM provider (OpenAI, Azure, Google, etc.)
' has a template of 9-13 configuration fields stored as List(Of AppConfigurationVariable).
'
' Data Flow:
' 1. PrepareConfigData() builds provider dictionary from hardcoded defaults
' 2. TryOverrideDefaultsFromRemote() optionally downloads updated templates from server
' 3. User selects provider from ComboBox -> LoadConfigForSelectedProvider() generates UI
' 4. User edits TextBoxes -> SaveCurrentInputToConfig() syncs to AppConfigurationVariable.CurrentValue
' 5. User clicks OK -> ValidateAllConfigs() enforces rules -> btnOK_Click() maps to ISharedContext
' 6. CreateAppConfig() writes INI files to %APPDATA%\redink.ini (or app-specific paths)
'
' Provider Templates (6 built-in):
' ---------------------------------
' Each provider defines 9-13 fields via AppConfigurationVariable:
'   Common fields (all providers):
'   - INI_APIKey: API key or private key
'   - INI_Temperature: LLM temperature (0.0-2.0)
'   - INI_Timeout: HTTP timeout in milliseconds
'   - INI_Model: Model identifier (gpt-4.1, gemini-2.5-pro, etc.)
'   - INI_Endpoint: API endpoint URL with {model} and {apikey} placeholders
'   - INI_HeaderA/B: HTTP header name and value template
'   - INI_APICall: JSON request body with {promptsystem}, {promptuser}, {temperature} placeholders
'   - INI_Response: JSON response field to extract (e.g., "content" or "text")
'
'   OAuth2-specific fields (Google Vertex only):
'   - INI_OAuth2ClientMail: Service account email
'   - INI_OAuth2Scopes: OAuth2 scopes (space-separated)
'   - INI_OAuth2Endpoint: Token endpoint URL
'   - INI_OAuth2ATExpiry: Access token lifetime in seconds
'
' Remote Configuration Updates:
' ------------------------------
' TryOverrideDefaultsFromRemote() downloads provider templates from RemoteDefaultsUrl.
' INI format specification:
'   [ProviderName]
'   Field1 = DisplayName|VarName|VarType|ValidationRule|DefaultValue
'   Field2 = DisplayName|VarName|VarType|ValidationRule|DefaultValue
'   ...
'   Note = Optional user-facing hint text
'
' Example remote INI:
'   [OpenAI]
'   Field1 = API Key:|INI_APIKey|String|NotEmpty|
'   Field2 = Temperature:|INI_Temperature|String|2.0|0.2
'   Field3 = Model:|INI_Model|String|NotEmpty|gpt-4.1
'   Note = Payment method required even for free accounts.
'
' AreDifferent() performs deep comparison (provider-by-provider, field-by-field).
' User is prompted only if differences detected. Network errors are silent (never block wizard).
'
' Validation Rules (ValidateAllConfigs):
' ---------------------------------------
' Enforced via AppConfigurationVariable.ValidationRule:
'   - "NotEmpty": String.IsNullOrWhiteSpace check
'   - "E-Mail": Contains "@" check (simple heuristic)
'   - "Hyperlink": StartsWith("http://") Or StartsWith("https://")
'   - ">0": Integer.TryParse && value > 0
'   - "0.0-2.0": Double.TryParse && value in range [0.0, 2.0]
'   - "\d+\.\d+" (e.g., "2.0"): Max value validation, accepts "." or "," as decimal separator
'
' Decimal separator normalization: Replaces "," with ".", rejects multiple decimal points.
'
' State Management:
' -----------------
' CurrentValue preservation during provider switching:
' 1. User selects "OpenAI", enters API key "sk-abc123...", temperature "0.5"
' 2. User switches to "Google Gemini" via cmbProvider dropdown
' 3. cmbProvider_SelectedIndexChanged fires:
'    a. SaveCurrentInputToSpecificConfig(OpenAI vars) copies TextBox.Text -> CurrentValue
'    b. _activeProvider = "Google Gemini"
'    c. LoadConfigForSelectedProvider() clears panel, creates new TextBoxes from Gemini defaults
' 4. User switches back to "OpenAI"
'    a. LoadConfigForSelectedProvider() reads OpenAI vars' preserved CurrentValue ("sk-abc123...", "0.5")
'    b. TextBoxes repopulated with saved values
'
' UI Layout (responsive):
' -----------------------
' Target width = Min(1050px, 80% screen width) for wrapping/sizing.
' Control hierarchy (top to bottom):
'   1. PictureBox (logo 50x50) + "Welcome to Red Ink" label
'   2. LinkLabel with instructions (wraps to target width)
'   3. "Select API provider:" label + ComboBox (6 providers)
'   4. LinkLabel with additional info
'   5. "Configuration For {Provider}:" label (updated on provider change)
'   6. panelConfig (scrollable, dynamic Label+TextBox pairs, auto-height)
'   7. "Use this config for Red Ink:" label + chkWord/chkOutlook/chkExcel
'   8. btnOK + btnCancel
'   9. invisibleLabel (forces form height adjustment)
'
' Panel height adjusts via PanelConfig_SizeChanged event -> repositions all controls below.
'
' Output Files (CreateAppConfig):
' --------------------------------
' Generates INI files for checked applications:
'   - Word: %APPDATA%\redink.ini (SharedMethods.GetDefaultINIPath("Word"))
'   - Excel: App-specific path if chkExcel checked
'   - Outlook: App-specific path if chkOutlook checked
'
' INI format:
'   ; Red Ink configuration file (automatically generated)
'   ; Minimum configuration for OpenAI
'   APIKey = sk-abc123...
'   Endpoint = https://api.openai.com/v1/chat/completions
'   HeaderA = Authorization
'   HeaderB = Bearer {apikey}
'   Temperature = 0.2
'   Model = gpt-4.1
'   OAuth2 = False
'   ...
'
' VarName "INI_APIKey" is stripped to "APIKey" for INI output (remove "INI_" prefix).
' Temperature is normalized to dot decimal separator (0.2 not 0,2) via CultureInfo.InvariantCulture.
'
' Error Handling:
' ---------------
'   - Remote download failures: Debug.WriteLine, silent (never block wizard)
'   - Validation failures: ShowCustomMessageBox, return False, keep wizard open
'   - File I/O errors: ShowCustomMessageBox in CreateAppConfig()
'   - Provider switching errors: Try/Catch in cmbProvider_SelectedIndexChanged (ignore timing issues)
'   - Parsing errors: TryParseRemoteDefaults returns False, wizard continues with built-in defaults
'
' Thread Safety:
' --------------
' NOT thread-safe. Must run on UI thread (standard WinForms requirement).
' Remote download uses synchronous Wait() on HttpClient.GetStringAsync task.
'
' Performance:
' ------------
'   - Form initialization: <100ms for built-in providers
'   - Remote defaults check: 10s timeout, user-prompted (opt-in)
'   - UI rendering: Dynamic control creation for 9-13 fields per provider (~50ms)
'   - Validation: Inline per-field checks, <1ms total
'   - INI file write: <10ms per file
'
' Extension Points:
' -----------------
'   - Add provider: Extend PrepareConfigData() with new SubAdd() call
'   - Custom validation: Add case to ValidateAllConfigs()
'   - UI controls: Modify LoadConfigForSelectedProvider() to support ComboBox, NumericUpDown, etc.
'   - Remote source: Change RemoteDefaultsUrl to organization-specific config server
'   - Localization: Replace hardcoded English strings with resource file lookups
'
' Maintenance Notes:
' ------------------
'   - When adding provider: Update PrepareConfigData, optionally add to defaultOrder list
'   - When adding VarName: Update btnOK_Click Select Case for ISharedContext mapping
'   - When adding validation rule: Update ValidateAllConfigs() and AppConfigurationVariable.vb docs
'   - Test with various screen sizes (min 900px, max 80% screen width)
'   - Test decimal separator handling with European regional settings (comma separator)
'
' Dependencies:
' -------------
'   - AppConfigurationVariable.vb: Data model for configuration fields
'   - ISharedContext (SharedContext.vb): Configuration storage interface
'   - SharedMethods: UI helpers (ShowCustomMessageBox, ShowCustomYesNoBox, GetDefaultINIPath, etc.)
'   - System.Windows.Forms: Form, Controls, DialogResult
'   - System.Net.Http: HttpClient for remote defaults download
'   - System.Drawing: Font, Color, Size, Point, Bitmap
'
' Known Limitations:
' ------------------
'   - VarType="String" for all fields; UI always renders TextBox (no ComboBox, NumericUpDown, etc.)
'   - Email validation is simplistic (contains "@" check)
'   - Remote defaults use custom INI format (not standard INI parser)
'   - No undo/redo for configuration edits
'   - No field-level tooltips (only provider-level notes at bottom)
'
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' First-run installation wizard for Red Ink LLM configuration.
    ''' Guides users through provider selection, credential entry, and INI file generation.
    ''' </summary>
    ''' <remarks>
    ''' Supports 6 providers (OpenAI, Azure, Google Gemini/Vertex, MTF, SafeSwissCloud).
    ''' Each provider has 9-13 configuration fields with validation.
    ''' State preserved when switching providers. Optional remote config updates.
    ''' See file header (lines 1-200) for architecture details.
    ''' </remarks>
    Public Class InitialConfig
        Inherits Form

        ''' <summary>Shared configuration context passed ByRef from host add-in (Word/Excel/Outlook).</summary>
        Private _context As ISharedContext

        ''' <summary>Provider selection dropdown. Items populated from providerConfigs keys.</summary>
        Private WithEvents cmbProvider As ComboBox

        ''' <summary>Checkbox: Apply configuration to Word add-in (creates %APPDATA%\redink.ini).</summary>
        Private chkWord As System.Windows.Forms.CheckBox

        ''' <summary>Checkbox: Apply configuration to Outlook add-in (creates separate config file).</summary>
        Private chkOutlook As System.Windows.Forms.CheckBox

        ''' <summary>Checkbox: Apply configuration to Excel add-in (creates separate config file).</summary>
        Private chkExcel As System.Windows.Forms.CheckBox

        ''' <summary>Scrollable panel containing dynamically generated Label+TextBox pairs for selected provider.</summary>
        Private panelConfig As Panel

        ''' <summary>Label displaying "Configuration For {ProviderName}:" above configuration fields.</summary>
        Private lblCurrentProvider As System.Windows.Forms.Label

        ''' <summary>Provider-to-configuration mapping. Key: Provider name, Value: List of 9-13 AppConfigurationVariable instances.</summary>
        Private providerConfigs As New Dictionary(Of String, List(Of AppConfigurationVariable))(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Optional per-provider notes displayed below configuration fields as gray Label.</summary>
        Private providerNotes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>List of controls (Label, TextBox, note Label) currently displayed in panelConfig.</summary>
        Private currentConfigControls As New List(Of Control)

        ''' <summary>OK button: Validates input, maps to ISharedContext, writes INI files, closes wizard.</summary>
        Private btnOK As Button

        ''' <summary>Cancel button: Sets DialogResult.Cancel, InitialConfigFailed = True, closes wizard.</summary>
        Private btnCancel As Button

        ''' <summary>Target width for form content (900px base + 150px, capped at 80% of screen width).</summary>
        Private ReadOnly _targetWidth As Integer

        ''' <summary>Flag to prevent event handlers from firing during InitializeComponent().</summary>
        Private isInitializing As Boolean = False

        ''' <summary>Invisible 1x10 label positioned at bottom of form to force height recalculation.</summary>
        Private invisibleLabel As New System.Windows.Forms.Label() With {
            .Size = New System.Drawing.Size(1, 10),
            .Visible = True
        }

        ''' <summary>Base width constant for form layout (900px). Actual target width = OverallWidth + 150, capped at 80% screen width.</summary>
        Private Const OverallWidth As Integer = 900

        ''' <summary>Label for "Use this config for Red Ink:" checkbox row. Positioned dynamically below panelConfig.</summary>
        Private lblUseThisConfig As System.Windows.Forms.Label

        ''' <summary>Tracks the provider whose fields are currently displayed in panelConfig. Used for state management during provider switching.</summary>
        Private _activeProvider As String = "OpenAI"

        ''' <summary>
        ''' Constructor: Initializes wizard with shared configuration context and responsive layout.
        ''' </summary>
        ''' <param name="context">Shared configuration interface passed ByRef from host add-in.</param>
        ''' <remarks>
        ''' Computes _targetWidth = Min(1050px, 80% screen width) for responsive layout.
        ''' Sets form size, disables AutoScroll, calls InitializeComponent(), sets FormBorderStyle = Fixed3D.
        ''' </remarks>
        Public Sub New(ByRef context As ISharedContext)
            _context = context
            _targetWidth = Math.Min(OverallWidth + 150, CInt(Screen.PrimaryScreen.WorkingArea.Width * 0.8))
            Me.Size = New System.Drawing.Size(_targetWidth + 20, 800)
            Me.AutoScroll = False
            Me.AutoSize = True
            Me.InitializeComponent()
            Me.FormBorderStyle = FormBorderStyle.Fixed3D
        End Sub


        ''' <summary>
        ''' Builds the wizard UI by creating and positioning all form controls.
        ''' </summary>
        ''' <remarks>
        ''' Creates responsive layout (900-1050px width) with header, provider ComboBox, 
        ''' dynamic configuration panel, application checkboxes, and OK/Cancel buttons.
        ''' ComboBox width: Min(670, Max(300, available space)). Panel width: _targetWidth.
        ''' See file header for detailed layout structure.
        ''' </remarks>
        Private Sub InitializeComponent()
            isInitializing = True

            ' Configure form properties
            Me.Text = $"{SharedMethods.AN} Initial Configuration Wizard"
            Me.FormBorderStyle = FormBorderStyle.None
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.BackColor = ColorTranslator.FromWin32(&H8000000F)
            Me.ControlBox = False  ' No min/max/close buttons
            Me.AutoScroll = True

            Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = standardFont

            ' PictureBox (Logo)
            Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo_Large)
            Dim pictureBox As New PictureBox() With {
            .Image = bmp,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
            pictureBox.SetBounds(10, 10, 50, 50)
            Me.Controls.Add(pictureBox)

            ' Label "Welcome to {AN}" next to the logo
            Dim lblWelcome As New System.Windows.Forms.Label() With {
            .Text = $"Welcome to {SharedMethods.AN}",
            .AutoSize = True,
            .Font = New System.Drawing.Font("Segoe UI", 12.0F, FontStyle.Bold, GraphicsUnit.Point)
        }
            lblWelcome.Location = New System.Drawing.Point(pictureBox.Right + 10, pictureBox.Top + (pictureBox.Height \ 2) - (lblWelcome.Height \ 2))
            Me.Controls.Add(lblWelcome)

            ' Resolve DefaultINIPath for Word (expanded for display in instructions)
            Dim defaultWordPath As String = ""
            Try
                If SharedMethods.DefaultINIPaths IsNot Nothing AndAlso SharedMethods.DefaultINIPaths.ContainsKey("Word") Then
                    defaultWordPath = SharedMethods.DefaultINIPaths("Word")
                    defaultWordPath = SharedMethods.ExpandEnvironmentVariables(defaultWordPath)
                End If
            Catch
                ' Ignore errors resolving Word path; use empty string
            End Try

            ' LinkLabel with instructions (wraps to target width for responsive layout)
            Dim lblInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(_targetWidth, 0),
            .Text =
                $"No configuration file '{SharedMethods.AN2}.ini' was found, in which all settings " &
                "can be made locally or centrally. Therefore, you can make the basic settings here, " &
                "which will then be saved to such a file. You can then expand it manually (e.g., to add more models); go to 'Settings', then 'Expert Config'. " &
                $"How all this works is explained in the manual, which you can find at {SharedMethods.AN4}." &
                If(String.IsNullOrWhiteSpace(defaultWordPath),
                   "",
                   $" {AN2} will be stored at {defaultWordPath} for Word, which will also be used by Excel and Outlook unless they have their own {SharedMethods.AN2}.ini.")
        }
            lblInfo.Location = New System.Drawing.Point(10, pictureBox.Bottom + 15)
            AddHandler lblInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            lblInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblInfo)

            ' Label + ComboBox "Select API provider:"
            Dim lblWhichAI As New System.Windows.Forms.Label() With {
            .Text = "Select API provider:",
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold)
        }
            lblWhichAI.Location = New System.Drawing.Point(10, lblInfo.Bottom + 20)
            Me.Controls.Add(lblWhichAI)

            ' ComboBox with responsive width (expanded from original 520px, max _targetWidth - label - margins)
            cmbProvider = New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = Math.Min(520 + 150, Math.Max(300, _targetWidth - lblWhichAI.Right - 30))
        }
            cmbProvider.Location = New System.Drawing.Point(lblWhichAI.Right + 10, lblWhichAI.Top - 2)
            Me.Controls.Add(cmbProvider)

            ' Second LinkLabel row (below combo, indented)
            Dim lblMoreInfo As New LinkLabel() With {
            .AutoSize = True,
            .MaximumSize = New Size(_targetWidth - 20, 0),
            .Text = $"Note: More on how to obtain access to one of these providers is described on {SharedMethods.AN4}. Getting an API access is not expensive. You can use the below form also for other providers. If this does not work or you need to configure more, abort and do it manually before restarting your application."
        }
            lblMoreInfo.Location = New System.Drawing.Point(30, cmbProvider.Bottom + 5)
            AddHandler lblMoreInfo.LinkClicked, AddressOf LinkLabel_LinkClicked
            lblMoreInfo.Links.Add(New LinkLabel.Link() With {
            .LinkData = $"{SharedMethods.AN4}",
            .Start = lblMoreInfo.Text.IndexOf($"{SharedMethods.AN4}", StringComparison.Ordinal),
            .Length = $"{SharedMethods.AN4}".Length
        })
            Me.Controls.Add(lblMoreInfo)

            ' Label for "Configuration for <AI Provider>:"
            lblCurrentProvider = New System.Windows.Forms.Label() With {
            .AutoSize = True,
            .Font = New System.Drawing.Font(standardFont, FontStyle.Bold),
            .Location = New System.Drawing.Point(10, lblMoreInfo.Bottom + 20)
        }
            Me.Controls.Add(lblCurrentProvider)

            ' Panel for dynamic input fields (scrollable, responsive width)
            panelConfig = New Panel() With {
            .AutoScroll = True,
            .Location = New System.Drawing.Point(10, lblCurrentProvider.Bottom + 5),
            .Width = _targetWidth
        }
            AddHandler panelConfig.SizeChanged, AddressOf PanelConfig_SizeChanged
            Me.Controls.Add(panelConfig)

            ' Build provider configuration data before populating combo
            PrepareConfigData()

            ' Populate combo in preferred order (OpenAI, Azure, Google Gemini, Google Vertex), then alphabetical
            Dim defaultOrder As New List(Of String) From {
            "OpenAI",
            "Microsoft Azure OpenAI Services",
            "Google Gemini",
            "Google Vertex"
        }
            For Each providerName As String In defaultOrder
                If providerConfigs.ContainsKey(providerName) Then cmbProvider.Items.Add(providerName)
            Next
            For Each providerName As String In providerConfigs.Keys
                If cmbProvider.Items.IndexOf(providerName) = -1 Then cmbProvider.Items.Add(providerName)
            Next
            If cmbProvider.Items.Count > 0 Then
                cmbProvider.SelectedIndex = 0
                _activeProvider = cmbProvider.SelectedItem.ToString()
            End If

            lblUseThisConfig = New System.Windows.Forms.Label() With {
            .Text = $"Use this config for {SharedMethods.AN}:",
            .Font = New System.Drawing.Font(Me.Font, FontStyle.Bold),
            .AutoSize = True
        }
            lblUseThisConfig.Location = New System.Drawing.Point(10, panelConfig.Bottom + 10)
            Me.Controls.Add(lblUseThisConfig)

            chkWord = New System.Windows.Forms.CheckBox() With {
            .Text = "for Word",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Word")
        }
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            Me.Controls.Add(chkWord)

            chkOutlook = New System.Windows.Forms.CheckBox() With {
            .Text = "for Outlook (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Outlook")
        }
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkOutlook)

            chkExcel = New System.Windows.Forms.CheckBox() With {
            .Text = "for Excel (as separate config)",
            .AutoSize = True,
            .Checked = _context.RDV.StartsWith("Excel")
        }
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 17, lblUseThisConfig.Top)
            Me.Controls.Add(chkExcel)

            btnOK = New Button() With {
            .Text = "OK, save this configuration and continue",
            .AutoSize = True
        }
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            AddHandler btnOK.Click, AddressOf btnOK_Click
            Me.Controls.Add(btnOK)

            btnCancel = New Button() With {
            .Text = "Cancel",
            .AutoSize = True
        }
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Controls.Add(invisibleLabel)

            LoadConfigForSelectedProvider()
            isInitializing = False
        End Sub


        ''' <summary>
        ''' Event handler for panel size changes. Repositions controls below panel and adjusts form height.
        ''' </summary>
        Private Sub PanelConfig_SizeChanged(sender As Object, e As EventArgs)
            If isInitializing OrElse lblUseThisConfig Is Nothing Then Exit Sub
            Dim panel = DirectCast(sender, Panel)
            lblUseThisConfig.Location = New System.Drawing.Point(10, panel.Bottom + 20)
            chkWord.Location = New System.Drawing.Point(lblUseThisConfig.Right + 10, lblUseThisConfig.Top)
            chkOutlook.Location = New System.Drawing.Point(chkWord.Right + 20, lblUseThisConfig.Top)
            chkExcel.Location = New System.Drawing.Point(chkOutlook.Right + 20, lblUseThisConfig.Top)
            btnOK.Location = New System.Drawing.Point(10, lblUseThisConfig.Bottom + 20)
            btnCancel.Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
            invisibleLabel.Location = New System.Drawing.Point(10, btnCancel.Bottom + 10)
            Me.Height = invisibleLabel.Bottom + 20
        End Sub


        ''' <summary>
        ''' Builds provider configuration templates with default values and optional notes.
        ''' </summary>
        ''' <remarks>
        ''' Creates templates for 6 providers (OpenAI, Azure, Google Gemini, Google Vertex, MTF, SafeSwissCloud).
        ''' Each has 9-13 fields. Google Vertex includes 4 OAuth2 fields. Validation rules: NotEmpty, Hyperlink, E-Mail, >0, 0.0-2.0.
        ''' Calls TryOverrideDefaultsFromRemote() to check for updated templates. Performance: ~20ms local, 10s max remote (user opt-in).
        ''' See file header "Provider Templates" section for field-by-field documentation.
        ''' </remarks>
        Private Sub PrepareConfigData()
            providerConfigs.Clear()
            providerNotes.Clear()

            ' Helper lambda to add provider with cloned variable list (prevents reference sharing)
            Dim SubAdd As Action(Of String, List(Of AppConfigurationVariable)) =
                Sub(name As String, vars As List(Of AppConfigurationVariable))
                    Dim clone As New List(Of AppConfigurationVariable)
                    For Each v In vars
                        clone.Add(New AppConfigurationVariable With {
                            .DisplayName = v.DisplayName,
                            .VarName = v.VarName,
                            .VarType = v.VarType,
                            .ValidationRule = v.ValidationRule,
                            .DefaultValue = v.DefaultValue,
                            .CurrentValue = v.DefaultValue
                        })
                    Next
                    providerConfigs(name) = clone
                End Sub

            ' OPENAI provider (9 fields + payment note)
            SubAdd("OpenAI",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-4.1"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://api.openai.com/v1/chat/completions"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("OpenAI") = "Note: When generating the API key with OpenAI, make sure you have added a valid payment method (e.g., credit card), even if you use ChatGPT for free or with an already paid subscription. You still need the payment method and a budget to pay for the actual consumption (costs are in our experience low)."

            ' MICROSOFT AZURE OPENAI SERVICES provider (9 fields, no note)
            SubAdd("Microsoft Azure OpenAI Services",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "0.0-2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-4.1"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://[your endpoint]/openai/deployments/[your deployment-id]/chat/completions?api-version=2024-06-01"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "api-key"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "{apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"", ""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("Microsoft Azure OpenAI Services") = ""

            ' GOOGLE GEMINI provider (9 fields, no note)
            SubAdd("Google Gemini",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gemini-2.5-pro"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "X-Goog-Api-Key"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "{apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""contents"": [{""role"": ""user"",""parts"": [{ ""text"": ""{promptsystem} {promptuser}"" }]}], ""generationConfig"": {""temperature"": {temperature}}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "text"}
                })
            providerNotes("Google Gemini") = ""

            ' GOOGLE VERTEX provider (13 fields: 9 common + 4 OAuth2, includes service account note)
            SubAdd("Google Vertex",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "Private Key (barebones, not PEM):", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gemini-2.5-pro"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://europe-west1-aiplatform.googleapis.com/v1/projects/[your project ID]/locations/europe-west1/publishers/google/models/{model}:generateContent"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""contents"": [{""role"": ""user"", ""parts"":[{""text"": ""{promptsystem} {promptuser}""}]}], ""generationConfig"": {""temperature"": {temperature}}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "text"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 'client_mail':", .VarName = "INI_OAuth2ClientMail", .VarType = "String", .ValidationRule = "E-Mail", .DefaultValue = "[service account mail]]@[your project ID].iam.gserviceaccount.com"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 'scopes':", .VarName = "INI_OAuth2Scopes", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "https://www.googleapis.com/auth/cloud-platform"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 Endpoint:", .VarName = "INI_OAuth2Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://oauth2.googleapis.com/token"},
                    New AppConfigurationVariable With {.DisplayName = "OAuth2 Access Token Expiry (ms):", .VarName = "INI_OAuth2ATExpiry", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "3600"}
                })
            providerNotes("Google Vertex") = "Note: Requires OAuth2 service account to be configured via the GCP console. Private Key must be the raw key (not PEM)."

            ' MTF provider (9 fields, no note)
            SubAdd("MTF",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "meta-llama-ai"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://api.ai.mtf.cloud/chatbot/ask"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("MTF") = ""

            ' SAFESWISSCLOUD provider (9 fields, no note)
            SubAdd("SafeSwissCloud",
                New List(Of AppConfigurationVariable) From {
                    New AppConfigurationVariable With {.DisplayName = "API Key:", .VarName = "INI_APIKey", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = ""},
                    New AppConfigurationVariable With {.DisplayName = "Temperature:", .VarName = "INI_Temperature", .VarType = "String", .ValidationRule = "0.0-2.0", .DefaultValue = "0.2"},
                    New AppConfigurationVariable With {.DisplayName = "Timeout (ms):", .VarName = "INI_Timeout", .VarType = "Integer", .ValidationRule = ">0", .DefaultValue = "200000"},
                    New AppConfigurationVariable With {.DisplayName = "Model:", .VarName = "INI_Model", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "gpt-oss-120b"},
                    New AppConfigurationVariable With {.DisplayName = "Endpoint:", .VarName = "INI_Endpoint", .VarType = "String", .ValidationRule = "Hyperlink", .DefaultValue = "https://llm01.safeswisscloud.ch/engines/{model}/chat/completions"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderA:", .VarName = "INI_HeaderA", .VarType = "String", .ValidationRule = "", .DefaultValue = "Authorization"},
                    New AppConfigurationVariable With {.DisplayName = "HeaderB:", .VarName = "INI_HeaderB", .VarType = "String", .ValidationRule = "", .DefaultValue = "Bearer {apikey}"},
                    New AppConfigurationVariable With {.DisplayName = "APICall:", .VarName = "INI_APICall", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "{""model"":   ""{model}"",  ""messages"": [{""role"": ""system"",""content"": ""{promptsystem}""},{""role"": ""user"",""content"": ""{promptuser}""}],""temperature"": {temperature}}"},
                    New AppConfigurationVariable With {.DisplayName = "Response tag:", .VarName = "INI_Response", .VarType = "String", .ValidationRule = "NotEmpty", .DefaultValue = "content"}
                })
            providerNotes("SafeSwissCloud") = ""

            ' Attempt to override with remote configuration if available and different
            TryOverrideDefaultsFromRemote()

        End Sub

        ''' <summary>Downloads remote configuration text from URL with timeout. Returns True on success.</summary>
        Private Function TryDownloadString(url As String, timeoutMs As Integer, ByRef content As String) As Boolean
            content = Nothing
            Try
                ' Ensure TLS 1.2 for HTTPS endpoints (many servers reject TLS 1.0/1.1)
                ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12

                Dim handler As New HttpClientHandler() With {
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
            }

                Using client As New System.Net.Http.HttpClient(handler)
                    client.Timeout = TimeSpan.FromMilliseconds(Math.Max(10000, timeoutMs)) ' 10s minimum

                    ' Use GetStringAsync for simplicity (auto-detects encoding from Content-Type header)
                    Dim readTask = client.GetStringAsync(url)
                    readTask.Wait()

                    If readTask.Status = TaskStatus.RanToCompletion Then
                        Dim s = readTask.Result
                        If Not String.IsNullOrWhiteSpace(s) Then
                            content = s
                            Return True
                        End If
                    End If
                End Using
            Catch ex As Exception
                ' Log error for diagnostics (visible in Output window during debugging)
                System.Diagnostics.Debug.WriteLine($"TryDownloadString error for {url}: {ex}")
            End Try
            Return False
        End Function

        ''' <summary>Parses remote INI-format configuration text into provider dictionaries. Custom INI parser for pipe-delimited field definitions.</summary>
        Private Function TryParseRemoteDefaults(ini As String,
                                            ByRef outConfigs As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                            ByRef outNotes As Dictionary(Of String, String)) As Boolean
            outConfigs = Nothing
            outNotes = Nothing
            If String.IsNullOrWhiteSpace(ini) Then Return False

            Try
                Dim cfg As New Dictionary(Of String, List(Of AppConfigurationVariable))(StringComparer.OrdinalIgnoreCase)
                Dim notes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                Dim lines = ini.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.None)
                Dim section As String = Nothing
                Dim sectionFields As New List(Of KeyValuePair(Of String, String))()

                Dim flushSection As Action =
                Sub()
                    If String.IsNullOrWhiteSpace(section) Then Return

                    Dim vars As New List(Of AppConfigurationVariable)()

                    ' Gather FieldN entries in numeric order (not alphabetical)
                    Dim ordered = sectionFields.
                        Where(Function(kv) kv.Key.StartsWith("Field", StringComparison.OrdinalIgnoreCase)).
                        Select(Function(kv)
                                   Dim numStr = New String(kv.Key.SkipWhile(Function(c) Not Char.IsDigit(c)).ToArray())
                                   Dim n As Integer = 0
                                   Integer.TryParse(numStr, n)
                                   Return New With {.Num = n, .Val = kv.Value}
                               End Function).
                        OrderBy(Function(x) x.Num).
                        ToList()

                    For Each f In ordered
                        Dim parts = (If(f.Val, "")).Split("|"c)
                        ' Expected format: DisplayName|VarName|VarType|ValidationRule|DefaultValue
                        If parts.Length >= 4 Then
                            Dim v As New AppConfigurationVariable With {
                                .DisplayName = parts(0).Trim(),
                                .VarName = parts(1).Trim(),
                                .VarType = parts(2).Trim(),
                                .ValidationRule = parts(3).Trim(),
                                .DefaultValue = If(parts.Length >= 5, parts(4), "")
                            }
                            v.CurrentValue = v.DefaultValue
                            vars.Add(v)
                        End If
                    Next

                    ' Extract optional Note entry
                    Dim noteValue = sectionFields.
                        FirstOrDefault(Function(kv) kv.Key.Equals("Note", StringComparison.OrdinalIgnoreCase)).Value
                    If Not String.IsNullOrWhiteSpace(noteValue) Then
                        notes(section) = noteValue.Trim()
                    End If

                    If vars.Count > 0 Then
                        cfg(section) = vars
                    End If

                    sectionFields.Clear()
                End Sub

                For Each raw In lines
                    Dim line = raw.Trim()
                    If line.Length = 0 Then Continue For
                    If line.StartsWith(";", StringComparison.Ordinal) OrElse line.StartsWith("#", StringComparison.Ordinal) Then Continue For

                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        flushSection()
                        section = line.Substring(1, line.Length - 2).Trim()
                        Continue For
                    End If

                    Dim eqIdx = line.IndexOf("="c)
                    If eqIdx > 0 AndAlso section IsNot Nothing Then
                        Dim key = line.Substring(0, eqIdx).Trim()
                        Dim val = line.Substring(eqIdx + 1).Trim()
                        sectionFields.Add(New KeyValuePair(Of String, String)(key, val))
                    End If
                Next

                flushSection()

                If cfg.Count > 0 Then
                    outConfigs = cfg
                    outNotes = notes
                    Return True
                End If
            Catch
                ' Parsing failure -> silently ignore, return False
            End Try

            Return False
        End Function

        ''' <summary>Compares two strings for equality with JSON-aware whitespace normalization.</summary>
        Private Function StringsEqual(a As String, b As String) As Boolean
            If Object.ReferenceEquals(a, b) Then Return True
            If a Is Nothing OrElse b Is Nothing Then
                Return String.IsNullOrEmpty(a) AndAlso String.IsNullOrEmpty(b)
            End If

            Dim sa = a.Replace(vbCrLf, vbLf).Trim()
            Dim sb = b.Replace(vbCrLf, vbLf).Trim()

            If String.Equals(sa, sb, StringComparison.Ordinal) Then Return True

            ' Heuristic: if both strings look JSON-ish, compare ignoring whitespace outside quotes
            Dim looksJsonA = (sa.IndexOf(":"c) >= 0) AndAlso (sa.Contains("{") OrElse sa.Contains("["))
            Dim looksJsonB = (sb.IndexOf(":"c) >= 0) AndAlso (sb.Contains("{") OrElse sb.Contains("["))

            If looksJsonA AndAlso looksJsonB Then
                Return String.Equals(StripWsOutsideQuotes(sa), StripWsOutsideQuotes(sb), StringComparison.Ordinal)
            End If

            Return False
        End Function


        ''' <summary>Removes all whitespace characters outside of quoted strings. Used by StringsEqual() for JSON-aware comparison.</summary>
        Private Function StripWsOutsideQuotes(s As String) As String
            Dim sb As New StringBuilder(s.Length)
            Dim inStr As Boolean = False
            Dim esc As Boolean = False

            For i As Integer = 0 To s.Length - 1
                Dim ch = s(i)
                If inStr Then
                    sb.Append(ch)
                    If esc Then
                        esc = False
                    ElseIf ch = "\"c Then
                        esc = True
                    ElseIf ch = """"c Then
                        inStr = False
                    End If
                Else
                    If ch = """"c Then
                        inStr = True
                        sb.Append(ch)
                    ElseIf Not Char.IsWhiteSpace(ch) Then
                        sb.Append(ch)
                    End If
                End If
            Next

            Return sb.ToString()
        End Function

        ''' <summary>Performs deep comparison of local vs. remote provider configurations. Returns True if differences detected.</summary>
        Private Function AreDifferent(localCfg As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                  localNotes As Dictionary(Of String, String),
                                  remoteCfg As Dictionary(Of String, List(Of AppConfigurationVariable)),
                                  remoteNotes As Dictionary(Of String, String)) As Boolean
            If remoteCfg Is Nothing OrElse remoteNotes Is Nothing Then
                System.Diagnostics.Debug.WriteLine("No remote config/notes available; treating as 'no differences'.")
                Return False
            End If

            Dim foundDiff As Boolean = False
            Dim S As Func(Of String, String) = Function(x) If(x, "<null>")

            ' Check for providers added/removed
            For Each p In remoteCfg.Keys
                If Not localCfg.ContainsKey(p) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider present in remote but missing locally: '{p}'")
                    foundDiff = True
                End If
            Next
            For Each p In localCfg.Keys
                If Not remoteCfg.ContainsKey(p) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider present locally but missing in remote: '{p}'")
                    foundDiff = True
                End If
            Next

            ' Per-provider comparison: check variables (added/removed/changed)
            For Each p In remoteCfg.Keys
                Dim rList = remoteCfg(p)
                Dim lList As List(Of AppConfigurationVariable) = Nothing
                If Not localCfg.TryGetValue(p, lList) Then
                    System.Diagnostics.Debug.WriteLine($"Difference: provider '{p}' exists in remote but not found in local config map.")
                    foundDiff = True
                    Continue For
                End If

                Dim rByName = rList.ToDictionary(Function(v) v.VarName, StringComparer.OrdinalIgnoreCase)
                Dim lByName = lList.ToDictionary(Function(v) v.VarName, StringComparer.OrdinalIgnoreCase)

                ' Check variables added/removed
                For Each k In rByName.Keys
                    If Not lByName.ContainsKey(k) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}]: variable added in remote: '{k}'")
                        foundDiff = True
                    End If
                Next
                For Each k In lByName.Keys
                    If Not rByName.ContainsKey(k) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}]: variable removed in remote: '{k}'")
                        foundDiff = True
                    End If
                Next

                ' Field-by-field comparison (including CurrentValue)
                For Each k In rByName.Keys
                    If Not lByName.ContainsKey(k) Then
                        Continue For
                    End If

                    Dim r = rByName(k)
                    Dim l = lByName(k)

                    If Not StringsEqual(l.DisplayName, r.DisplayName) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: DisplayName local='{S(l.DisplayName)}' remote='{S(r.DisplayName)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.VarType, r.VarType) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: VarType local='{S(l.VarType)}' remote='{S(r.VarType)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.ValidationRule, r.ValidationRule) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: ValidationRule local='{S(l.ValidationRule)}' remote='{S(r.ValidationRule)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.DefaultValue, r.DefaultValue) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: DefaultValue local='{S(l.DefaultValue)}' remote='{S(r.DefaultValue)}'")
                        foundDiff = True
                    End If
                    If Not StringsEqual(l.CurrentValue, r.CurrentValue) Then
                        System.Diagnostics.Debug.WriteLine($"Difference[{p}.{k}]: CurrentValue local='{S(l.CurrentValue)}' remote='{S(r.CurrentValue)}'")
                        foundDiff = True
                    End If
                Next
            Next

            ' Compare notes: check union of providers to catch added/removed notes
            Dim allProviders = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each k In localCfg.Keys : allProviders.Add(k) : Next
            For Each k In remoteCfg.Keys : allProviders.Add(k) : Next

            For Each p In allProviders
                Dim ln As String = Nothing
                Dim rn As String = Nothing
                localNotes.TryGetValue(p, ln)
                remoteNotes.TryGetValue(p, rn)
                If Not StringsEqual(ln, rn) Then
                    System.Diagnostics.Debug.WriteLine($"Difference[{p}]: Provider note changed. localLen={If(ln, "").Length} remoteLen={If(rn, "").Length}")
                    foundDiff = True
                End If
            Next

            If Not foundDiff Then
                System.Diagnostics.Debug.WriteLine("No differences detected between local and remote config/notes.")
            End If

            Return foundDiff
        End Function

        ''' <summary>
        ''' Attempts to download and apply remote provider configuration defaults.
        ''' Prompts user for download permission, compares with local defaults, optionally overrides.
        ''' </summary>
        ''' <remarks>
        ''' Prompts user to check RemoteDefaultsUrl for updated defaults. Downloads via TryDownloadString (10s timeout),
        ''' parses via TryParseRemoteDefaults, compares via AreDifferent. If differences found, prompts to override.
        ''' Never blocks wizard startup on network errors (all failures silent). Performance: ~200ms on fast network, 10s max timeout.
        ''' </remarks>
        Private Sub TryOverrideDefaultsFromRemote()

            Dim answer = ShowCustomYesNoBox($"You are about to run the {AN} Installation Wizard. Do you want to check on {RemoteDefaultsUrl} for updated default configuration information?", "Yes", "No, keep built-in")

            If answer <> 1 Then Return

            Try
                Dim remoteText As String = Nothing
                If Not TryDownloadString(RemoteDefaultsUrl, 10000, remoteText) Then Exit Sub

                Dim rCfg As Dictionary(Of String, List(Of AppConfigurationVariable)) = Nothing
                Dim rNotes As Dictionary(Of String, String) = Nothing
                If Not TryParseRemoteDefaults(remoteText, rCfg, rNotes) Then Exit Sub

                If rCfg Is Nothing OrElse rCfg.Count = 0 Then Exit Sub

                If AreDifferent(providerConfigs, providerNotes, rCfg, rNotes) Then
                    Dim choice = SharedMethods.ShowCustomYesNoBox(
                    "Updated Default provider configurations are available online. Do you want To load And use those instead Of the built-In defaults now?",
                    "Use online defaults",
                    "Keep built-In")

                    ' Convention: return 1 for first button ("Use online defaults")
                    If choice = 1 Then
                        providerConfigs = New Dictionary(Of String, List(Of AppConfigurationVariable))(rCfg, StringComparer.OrdinalIgnoreCase)
                        providerNotes = New Dictionary(Of String, String)(rNotes, StringComparer.OrdinalIgnoreCase)
                    End If
                Else
                    ShowCustomMessageBox("No updates found. Keeping built-in defaults.")

                End If
            Catch
                ' Never fail the wizard on remote errors (network, parsing, comparison exceptions all ignored)
            End Try
        End Sub



        ''' <summary>Event handler for provider selection change. Saves current provider's input, switches active provider, regenerates UI.</summary>
        Private Sub cmbProvider_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProvider.SelectedIndexChanged
            Try
                ' Step 1: Save values into the previously displayed provider
                Dim prevList As List(Of AppConfigurationVariable) = GetConfigListByName(_activeProvider)
                SaveCurrentInputToSpecificConfig(prevList)

                ' Step 2: Switch active provider to the newly selected one
                If cmbProvider.SelectedItem IsNot Nothing Then
                    _activeProvider = cmbProvider.SelectedItem.ToString()
                End If

                ' Step 3: Load UI for new provider
                LoadConfigForSelectedProvider()
            Catch
                ' Ignore minor UI timing issues (silent failure acceptable)
            End Try
        End Sub

        ''' <summary>Copies current UI input (TextBox values) to AppConfigurationVariable.CurrentValue for active provider.</summary>
        Private Sub SaveCurrentInputToConfig()
            Dim selectedList = GetSelectedConfigList()
            If selectedList Is Nothing OrElse currentConfigControls.Count = 0 Then Return

            For i As Integer = 0 To currentConfigControls.Count - 1
                Dim ctrl = currentConfigControls(i)
                If TypeOf ctrl Is System.Windows.Forms.Label Then
                    Dim labelText = CType(ctrl, System.Windows.Forms.Label).Text
                    Dim configVar = selectedList.FirstOrDefault(Function(x) x.DisplayName = labelText)
                    If configVar IsNot Nothing Then
                        If i + 1 < currentConfigControls.Count Then
                            Dim inputControl = currentConfigControls(i + 1)
                            If TypeOf inputControl Is TextBox Then
                                configVar.CurrentValue = CType(inputControl, TextBox).Text
                            End If
                        End If
                    End If
                End If
            Next
        End Sub


        ''' <summary>Saves current UI inputs into specified provider config list during provider switching.</summary>
        ''' <param name="targetConfig">Provider's variable list to update (from GetConfigListByName).</param>
        Private Sub SaveCurrentInputToSpecificConfig(targetConfig As List(Of AppConfigurationVariable))
            If targetConfig Is Nothing OrElse currentConfigControls.Count = 0 Then Return

            For i As Integer = 0 To currentConfigControls.Count - 1
                Dim ctrl As System.Windows.Forms.Control = currentConfigControls(i)
                If TypeOf ctrl Is System.Windows.Forms.Label Then
                    Dim labelText As String = CType(ctrl, System.Windows.Forms.Label).Text
                    Dim configVar As AppConfigurationVariable = targetConfig.FirstOrDefault(Function(x) x.DisplayName = labelText)
                    If configVar IsNot Nothing AndAlso i + 1 < currentConfigControls.Count Then
                        Dim inputControl As System.Windows.Forms.Control = currentConfigControls(i + 1)
                        If TypeOf inputControl Is System.Windows.Forms.TextBox Then
                            configVar.CurrentValue = CType(inputControl, System.Windows.Forms.TextBox).Text
                        End If
                    End If
                End If
            Next
        End Sub

        ''' <summary>Dynamically generates Label+TextBox pairs for selected provider's configuration fields.</summary>
        ''' <remarks>
        ''' Two-pass layout: Pass 1 calculates max label width, Pass 2 creates controls.
        ''' TextBox width = panelConfig.Width - maxLabelWidth - 30px. Sets panel height, triggers PanelConfig_SizeChanged.
        ''' Performance: ~30ms per provider switch (9-13 fields). See file header for layout details.
        ''' </remarks>
        Private Sub LoadConfigForSelectedProvider()
            Dim selectedList As List(Of AppConfigurationVariable) = GetConfigListByName(_activeProvider)
            If selectedList Is Nothing Then Return

            ' Clear panel contents
            panelConfig.Controls.Clear()
            currentConfigControls.Clear()

            ' Update header label
            lblCurrentProvider.Text = "Configuration For " & _activeProvider & ":"

            ' Pass 1: Calculate maximum label width for alignment
            Dim yPos As Integer = 0
            Dim maxLabelWidth As Integer = 0
            For Each configVar In selectedList
                Dim lbl As New System.Windows.Forms.Label() With {
                .Text = configVar.DisplayName,
                .AutoSize = True,
                .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
            }
                maxLabelWidth = Math.Max(maxLabelWidth, lbl.PreferredWidth)
            Next

            ' Pass 2: Create and position Label+TextBox pairs
            For Each configVar In selectedList
                ' Create label with DisplayName
                Dim lbl As New System.Windows.Forms.Label() With {
                .Text = configVar.DisplayName,
                .AutoSize = True,
                .Font = New System.Drawing.Font(Me.Font, FontStyle.Regular)
            }
                lbl.Location = New System.Drawing.Point(0, yPos)
                panelConfig.Controls.Add(lbl)
                currentConfigControls.Add(lbl)

                ' Create TextBox with CurrentValue
                Dim txt As New TextBox() With {
                .Width = panelConfig.Width - maxLabelWidth - 30,
                .Text = configVar.CurrentValue
            }
                txt.Location = New System.Drawing.Point(maxLabelWidth + 10, yPos - 2)
                panelConfig.Controls.Add(txt)
                currentConfigControls.Add(txt)

                yPos += lbl.Height + 8
            Next

            ' Append provider note (if exists) at end of field list
            Dim endNote As String = Nothing
            providerNotes.TryGetValue(_activeProvider, endNote)
            If Not String.IsNullOrWhiteSpace(endNote) Then
                Dim noteLabel As New System.Windows.Forms.Label() With {
                .AutoSize = True,
                .MaximumSize = New Size(panelConfig.Width - maxLabelWidth - 30, 0),
                .Text = endNote,
                .ForeColor = SystemColors.GrayText
            }
                noteLabel.Location = New System.Drawing.Point(maxLabelWidth + 10, yPos)
                panelConfig.Controls.Add(noteLabel)
                currentConfigControls.Add(noteLabel)
                yPos += noteLabel.Height + 8
            End If

            panelConfig.Height = yPos + 2
        End Sub

        ''' <summary>Helper method to retrieve provider's configuration variable list by name.</summary>
        ''' <param name="name">Provider name (e.g., "OpenAI", "Google Vertex").</param>
        ''' <returns>List of AppConfigurationVariable instances, or Nothing if provider not found.</returns>
        Private Function GetConfigListByName(name As String) As List(Of AppConfigurationVariable)
            If String.IsNullOrEmpty(name) Then Return Nothing
            Dim list As List(Of AppConfigurationVariable) = Nothing
            If providerConfigs.TryGetValue(name, list) Then
                Return list
            End If
            Return Nothing
        End Function

        ''' <summary>Returns configuration variable list for currently active provider.</summary>
        ''' <returns>List of AppConfigurationVariable instances for active provider, or Nothing if not found.</returns>
        Private Function GetSelectedConfigList() As List(Of AppConfigurationVariable)
            Return GetConfigListByName(_activeProvider)
        End Function

        ''' <summary>OK button click handler. Validates input, maps to ISharedContext, writes INI files, closes wizard.</summary>
        ''' <remarks>
        ''' Execution: SaveCurrentInputToConfig → ValidateAllConfigs → Map VarName to ISharedContext → 
        ''' CreateAppConfig (Word/Excel/Outlook) → Set DialogResult.OK, InitialConfigFailed=False, Close().
        ''' Special case: Sets INI_OAuth2=True for Google Vertex. CInt() converts Timeout/OAuth2ATExpiry to Integer.
        ''' </remarks>
        Private Sub btnOK_Click(sender As Object, e As EventArgs)
            Try
                ' Save inputs from current panel to CurrentValue
                SaveCurrentInputToConfig()

                ' Validate all fields
                If Not ValidateAllConfigs() Then
                    Return
                End If

                ' If validation passed: Get selected provider's variable list
                Dim finalList = GetSelectedConfigList()
                If finalList Is Nothing Then
                    SharedMethods.ShowCustomMessageBox("No AI provider selected.")
                    Return
                End If

                ' Map VarName -> CurrentValue to _context properties
                For Each cv In finalList
                    Select Case cv.VarName
                        Case "INI_APIKey" : _context.INI_APIKey = cv.CurrentValue
                        Case "INI_Temperature" : _context.INI_Temperature = cv.CurrentValue
                        Case "INI_Timeout" : _context.INI_Timeout = CInt(cv.CurrentValue)
                        Case "INI_Model" : _context.INI_Model = cv.CurrentValue
                        Case "INI_Endpoint" : _context.INI_Endpoint = cv.CurrentValue
                        Case "INI_HeaderA" : _context.INI_HeaderA = cv.CurrentValue
                        Case "INI_HeaderB" : _context.INI_HeaderB = cv.CurrentValue
                        Case "INI_APICall" : _context.INI_APICall = cv.CurrentValue
                        Case "INI_Response" : _context.INI_Response = cv.CurrentValue
                        Case "INI_OAuth2ClientMail" : _context.INI_OAuth2ClientMail = cv.CurrentValue
                        Case "INI_OAuth2Scopes" : _context.INI_OAuth2Scopes = cv.CurrentValue
                        Case "INI_OAuth2Endpoint" : _context.INI_OAuth2Endpoint = cv.CurrentValue
                        Case "INI_OAuth2ATExpiry" : _context.INI_OAuth2ATExpiry = CInt(cv.CurrentValue)
                    End Select
                Next

                ' Only Google Vertex requires OAuth2 by default
                If String.Equals(_activeProvider, "Google Vertex", StringComparison.OrdinalIgnoreCase) Then
                    _context.INI_OAuth2 = True
                End If

                _context.INIloaded = False

                Dim providerName As String = _activeProvider

                If chkWord.Checked Then CreateAppConfig("Word", providerName)
                If chkExcel.Checked Then CreateAppConfig("Excel", providerName)
                If chkOutlook.Checked Then CreateAppConfig("Outlook", providerName)

                ' Close wizard
                Me.DialogResult = DialogResult.OK
                _context.InitialConfigFailed = False
                Me.Close()

            Catch ex As System.Exception
                MessageBox.Show("Error in btnOK_Click: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>Cancel button click handler. Closes wizard without saving, sets InitialConfigFailed flag.</summary>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            _context.InitialConfigFailed = True
            Me.Close()
        End Sub

        ''' <summary>Validates all configuration fields for active provider against validation rules.</summary>
        ''' <returns>True if validation succeeded, False if failed (shows error, keeps wizard open).</returns>
        ''' <remarks>
        ''' Rules: NotEmpty, E-Mail (@), Hyperlink (http/https), >0 (positive int), 0.0-2.0 (range), \d+\.\d+ (max value).
        ''' Accepts "." or "," as decimal separator. Validates app checkboxes (host must have at least one checked).
        ''' BUG: "0.0-2.0" rule has early Return True (skips remaining fields).
        ''' </remarks>
        Private Function ValidateAllConfigs() As Boolean
            Dim selectedList = GetSelectedConfigList()

            ' Check if at least one relevant checkbox is checked for current host
            If _context.RDV.StartsWith("Word") AndAlso Not chkWord.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Word' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Outlook") AndAlso Not chkOutlook.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Outlook' checkbox needs to be checked.")
                Return False
            ElseIf _context.RDV.StartsWith("Excel") AndAlso Not chkExcel.Checked Then
                SharedMethods.ShowCustomMessageBox("At least the 'for Excel' checkbox needs to be checked.")
                Return False
            End If

            For Each cv In selectedList
                Dim valRule = cv.ValidationRule
                Dim valValue = cv.CurrentValue

                Debug.WriteLine("Validating: valrule=" & valRule & ", valValue='" & valValue & "'")

                ' NotEmpty validation
                If valRule.Contains("NotEmpty") Then
                    If String.IsNullOrWhiteSpace(valValue) Then
                        SharedMethods.ShowCustomMessageBox("Value For '" & cv.DisplayName & "' cannot be empty.")
                        Return False
                    End If
                End If

                ' E-Mail validation (simple @ check)
                If valRule.Contains("E-Mail") Then
                    If Not valValue.Contains("@") Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid e-mail address.")
                        Return False
                    End If
                End If

                ' Hyperlink validation (http/https protocol check)
                If valRule.Contains("Hyperlink") Then
                    If Not (valValue.StartsWith("http://") OrElse valValue.StartsWith("https://")) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a valid URL (http/https).")
                        Return False
                    End If
                End If

                ' Positive integer validation (>0)
                If valRule.Contains(">0") Then
                    Dim intVal As Integer
                    If Not Integer.TryParse(valValue, intVal) OrElse intVal <= 0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be an integer larger than 0.")
                        Return False
                    End If
                End If

                ' Explicit range validation (0.0-2.0) [backwards compatibility with old field validation rule]
                If valRule.Contains("0.0-2.0") Then
                    Dim dblVal As Double
                    If Not Double.TryParse(valValue, dblVal) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a floating number between 0.0 and 2.0.")
                        Return False
                    End If
                    If dblVal < 0.0 OrElse dblVal > 2.0 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be in [0.0 .. 2.0].", "Validation Error")
                        Return False
                    End If
                    Return True  ' Do not continue with further validation to avoid conflicting with next rule
                End If

                ' Max value validation (regex pattern \d+\.\d+, e.g., "2.0")
                If System.Text.RegularExpressions.Regex.IsMatch(valRule.Trim(), "^\d+\.\d+$") Then
                    Dim maxVal As Double
                    If Not Double.TryParse(valRule.Trim(),
                                               System.Globalization.NumberStyles.Float,
                                               System.Globalization.CultureInfo.InvariantCulture,
                                               maxVal) Then
                        SharedMethods.ShowCustomMessageBox("Internal validation error: cannot parse max value rule '" & valRule & "'.")
                        Return False
                    End If

                    Dim rawValue As String = valValue.Trim()

                    ' Normalize decimal separator: allow either "," or "."
                    ' Replace comma with dot; reject if more than one dot afterwards (invalid format like thousand separators)
                    Dim normalized As String = rawValue.Replace(",", ".")
                    If normalized.Count(Function(c) c = "."c) > 1 Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' is not a valid decimal number. Use one decimal point ('.' or ','). " &
                                                               "(Example: 1.25 or 1,25)")
                        Return False
                    End If

                    Dim dblVal As Double
                    If Not Double.TryParse(normalized,
                                               System.Globalization.NumberStyles.Float,
                                               System.Globalization.CultureInfo.InvariantCulture,
                                               dblVal) Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be a decimal number between 0 and " &
                                                               maxVal.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) &
                                                               " (accepts '.' or ',' as decimal separator).")
                        Return False
                    End If

                    If dblVal < 0.0 OrElse dblVal > maxVal Then
                        SharedMethods.ShowCustomMessageBox("Value for '" & cv.DisplayName & "' must be between 0 and " &
                                                               maxVal.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) &
                                                               ". Entered: " & rawValue)
                        Return False
                    End If
                End If

            Next

            Return True
        End Function

        ''' <summary>Event handler for LinkLabel clicks. Opens documentation URL in default browser.</summary>
        Private Sub LinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
            Try
                Dim link = e.Link.LinkData.ToString()
                System.Diagnostics.Process.Start(link)
            Catch ex As System.Exception
                MessageBox.Show("Could not open link. Error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Writes Red Ink configuration INI file for specified Office application.</summary>
        ''' <param name="App">Office application name ("Word", "Excel", or "Outlook").</param>
        ''' <param name="provider">Provider display name for INI comment.</param>
        ''' <remarks>
        ''' Creates INI at %APPDATA%\redink.ini (Word) or app-specific paths (Excel/Outlook).
        ''' Writes 14 keys. Temperature normalized to dot separator (0.2 not 0,2). UTF-8, CRLF. ~10ms per file.
        ''' </remarks>
        Private Sub CreateAppConfig(App As String, provider As String)
            Try
                ' Define the file path
                Dim filepath = SharedMethods.GetDefaultINIPath(App)

                Debug.WriteLine($"Creating {SharedMethods.AN} configuration file: " & filepath)

                ' Open a StreamWriter to create the file
                Using writer As New System.IO.StreamWriter(filepath)
                    ' Write the header
                    writer.WriteLine($"; {SharedMethods.AN} configuration file (automatically generated)")
                    writer.WriteLine(";")
                    writer.WriteLine($"; Go to {SharedMethods.AN4} on how to find the instructions to manually add or change the configuration settings")

                    ' Write an empty line
                    writer.WriteLine()

                    ' Write provider information
                    writer.WriteLine($"; Minimum configuration for {provider}")

                    ' Write another empty line
                    writer.WriteLine()

                    ' Normalize Temperature to use dot as decimal separator
                    Dim normalizedTemp As String = _context.INI_Temperature
                    If Not String.IsNullOrWhiteSpace(normalizedTemp) Then
                        Dim tempValue As Double
                        ' Parse with current culture, then format with invariant culture (dot separator)
                        If Double.TryParse(normalizedTemp.Replace(","c, "."c),
                                       System.Globalization.NumberStyles.Float,
                                       System.Globalization.CultureInfo.InvariantCulture,
                                       tempValue) Then
                            normalizedTemp = tempValue.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                        End If
                    End If

                    ' Loop through the dictionary and write each configuration value
                    Dim MinimumConfigValues As New Dictionary(Of String, String) From {
                        {"APIKey", _context.INI_APIKey},
                        {"Endpoint", _context.INI_Endpoint},
                        {"HeaderA", _context.INI_HeaderA},
                        {"HeaderB", _context.INI_HeaderB},
                        {"Response", _context.INI_Response},
                        {"APICall", _context.INI_APICall},
                        {"Timeout", _context.INI_Timeout.ToString()},
                        {"Temperature", normalizedTemp},
                        {"Model", _context.INI_Model},
                        {"OAuth2", _context.INI_OAuth2.ToString()},
                        {"OAuth2ClientMail", _context.INI_OAuth2ClientMail},
                        {"OAuth2Scopes", _context.INI_OAuth2Scopes},
                        {"OAuth2Endpoint", _context.INI_OAuth2Endpoint},
                        {"OAuth2ATExpiry", _context.INI_OAuth2ATExpiry.ToString()}
                    }

                    For Each kvp In MinimumConfigValues
                        writer.WriteLine($"{kvp.Key} = {kvp.Value}")
                    Next
                End Using

            Catch ex As System.Exception
                ' Handle errors by showing a custom message box
                SharedMethods.ShowCustomMessageBox($"Error creating configuration file: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace
