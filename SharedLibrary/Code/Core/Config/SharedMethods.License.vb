' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Net
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Main license check function. Returns True if license is valid, False otherwise.
        ''' </summary>
        Public Shared Function LicenseOK(ByVal context As ISharedContext,
                                          ByVal configDict As Dictionary(Of String, String)) As Boolean
            Try
                ' Load license settings from config file and My.Settings
                LoadLicenseSettings(configDict, context)

                ' Check if license checking is disabled
                If LicenseCheckDisabled Then
                    Return True
                End If

                ' Determine if this is GA version or Beta
                Dim isGAVersion As Boolean = AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)

                If isGAVersion Then
                    Return PerformGALicenseCheck(context)
                Else
                    Return PerformBetaLicenseCheck(context)
                End If

            Catch ex As Exception
                ' Fault tolerance - if license check fails, log and allow continuation
                Try
                    ShowCustomMessageBox($"License check encountered an error: {ex.Message}. Continuing with limited functionality.", AN)
                Catch
                End Try
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Loads license settings from config file and My.Settings
        ''' </summary>
        Private Shared Sub LoadLicenseSettings(configDict As Dictionary(Of String, String), Optional context As ISharedContext = Nothing)
            Try
                ' Reset the config flag
                LicenseFromConfig = False

                ' Load LicenseContact from config
                LicenseContact = If(configDict.ContainsKey("LicenseContact"), configDict("LicenseContact"), "")

                ' Load LicenseNoWarning from config
                LicenseNoWarning = ParseBoolean(configDict, "LicenseNoWarning", False)

                ' Determine if this is a Beta version (AppsUrl not pointing to redink.ai)
                Dim isBetaVersion As Boolean = Not AppsUrl.StartsWith(NewHomeURL, StringComparison.OrdinalIgnoreCase)

                ' Check if LicensedTill is "False" or more than 100 years in the future (disable checking)
                If configDict.ContainsKey("LicensedTill") Then
                    Dim configValue = configDict("LicensedTill").Trim()

                    ' Check for explicit disable values
                    If configValue.Equals("False", StringComparison.OrdinalIgnoreCase) OrElse
               configValue.Equals("No", StringComparison.OrdinalIgnoreCase) Then
                        LicenseCheckDisabled = True
                        LicenseStatus = If(configDict.ContainsKey("LicenseStatus"), configDict("LicenseStatus"), "")
                        If configDict.ContainsKey("LicenseUsers") Then Integer.TryParse(configDict("LicenseUsers"), LicenseUsers)
                        Return
                    End If

                    ' Try to parse the date from config
                    Dim configDate As Date
                    If Date.TryParse(configValue, configDate) Then
                        ' Check if more than 100 years in future (disable checking)
                        If configDate > Date.Now.AddYears(LicenseCheckDisabledYears) Then
                            LicenseCheckDisabled = True
                            Return
                        End If
                        ' Config file value takes precedence over everything
                        LicensedTill = configDate
                        LicenseFromConfig = True ' Mark that license came from config
                    Else
                        ' Could not parse date - fall through to defaults
                        LicensedTill = If(isBetaVersion, BetaEndDate, Date.MinValue)
                    End If
                Else
                    ' No LicensedTill in config file
                    If isBetaVersion Then
                        ' Beta version
                        LicensedTill = BetaEndDate
                        LicenseUsers = 1
                        LicenseStatus = "Beta Test License"
                    Else
                        ' GA version: use My.Settings if available, otherwise Date.MinValue
                        ' (which will trigger the license entry form)
                        Try
                            If My.Settings.LicensedTill > Date.MinValue AndAlso
                                My.Settings.LicensedTill < Date.MaxValue Then
                                LicensedTill = My.Settings.LicensedTill

                                ' Check if the license type has a FixedEndDate that is beyond current LicensedTill
                                ' and silently update to the new date (only if context is available)
                                If context IsNot Nothing Then
                                    TryExtendLicenseToFixedEndDate(context)
                                End If
                            Else
                                ' No valid license date - set to MinValue to trigger entry form
                                LicensedTill = Date.MinValue
                            End If
                        Catch
                            LicensedTill = Date.MinValue
                        End Try
                    End If
                End If

                ' Load LicenseStatus - config takes precedence
                If configDict.ContainsKey("LicenseStatus") Then
                    LicenseStatus = configDict("LicenseStatus")
                    ' If config has LicenseStatus, also get LicenseUsers from config
                    If configDict.ContainsKey("LicenseUsers") Then
                        Integer.TryParse(configDict("LicenseUsers"), LicenseUsers)
                    End If
                ElseIf Not isBetaVersion Then
                    ' Use My.Settings
                    Try
                        LicenseStatus = If(String.IsNullOrEmpty(My.Settings.LicenseStatus), "", My.Settings.LicenseStatus)
                        LicenseUsers = If(My.Settings.LicenseUsers > 0, My.Settings.LicenseUsers, 1)
                    Catch
                        LicenseStatus = ""
                        LicenseUsers = 1
                    End Try
                End If

            Catch ex As Exception
                ' Fault tolerance - determine reasonable defaults based on version
                LicenseCheckDisabled = False
                LicenseStatus = ""
                LicenseFromConfig = False
                Dim isBeta = Not AppsUrl.StartsWith("https://redink.ai", StringComparison.OrdinalIgnoreCase)
                LicensedTill = If(isBeta, BetaEndDate, Date.MinValue)
            End Try
        End Sub

        ''' <summary>
        ''' Checks if the current license type (from My.Settings.LicenseStatus) has a FixedEndDate
        ''' that is beyond the current LicensedTill and silently updates if so.
        ''' Only extends if the license type does not allow user-defined end dates.
        ''' </summary>
        Private Shared Sub TryExtendLicenseToFixedEndDate(context As ISharedContext)
            Try
                ' Only proceed if we have a valid LicenseStatus from My.Settings
                Dim storedStatus As String = ""
                Try
                    storedStatus = My.Settings.LicenseStatus
                Catch
                    Return
                End Try

                If String.IsNullOrEmpty(storedStatus) Then Return

                ' Get the version date from context.RDV
                Dim versionDate As Date = ParseVersionDateFromRDV(context.RDV)

                ' Get the license types using the version date
                Dim licenseTypes = GetLicenseTypes(versionDate)

                ' Find the matching license type by name
                Dim matchingType As LicenseTypeInfo = Nothing
                For Each lt In licenseTypes
                    If lt.Name.Equals(storedStatus, StringComparison.OrdinalIgnoreCase) Then
                        matchingType = lt
                        Exit For
                    End If
                Next

                If matchingType Is Nothing Then Return

                ' Only extend if the license type does NOT allow user-defined end dates
                If matchingType.UserDefinedEndDate Then Return

                ' Check if this license type has a FixedEndDate
                If Not matchingType.FixedEndDate.HasValue Then Return

                Dim fixedDate As Date = matchingType.FixedEndDate.Value

                ' Only extend if the fixed date is beyond the current LicensedTill
                If fixedDate > LicensedTill Then
                    ' Silently update the license end date
                    LicensedTill = fixedDate

                    ' Persist the updated date to My.Settings
                    Try
                        My.Settings.LicensedTill = LicensedTill
                        My.Settings.Save()
                    Catch
                        ' Ignore save errors - we still use the updated value in memory
                    End Try
                End If

            Catch
                ' Fault tolerance - ignore any errors in the extension check
            End Try
        End Sub

        ''' <summary>
        ''' Performs license check for GA (General Audience) version
        ''' </summary>
        Private Shared Function PerformGALicenseCheck(context As ISharedContext) As Boolean
            Dim licenseExpired As Boolean = False
            Dim noLicenseConfigured As Boolean = False

            ' If license came from config file with a valid future date, accept it without further checks
            If LicenseFromConfig Then
                If Date.Now > LicensedTill Then
                    ' Check if we're within the grace period
                    If GracePeriodDays > 0 AndAlso Date.Now <= LicensedTill.AddDays(GracePeriodDays) Then
                        ' Within grace period - show warning and allow continuation
                        CheckGracePeriodWarning(context, LicensedTill)
                        Return True
                    End If

                    Dim msg = BuildLicenseMessage(
        $"Your license for {AN} for {context.RDV} has EXPIRED on {LicensedTill:d}." & vbCrLf & vbCrLf &
        "Please contact your administrator to update the license configuration.")
                    ShowCustomMessageBox(msg, $"{AN} License Expired")
                    Return False
                End If

                If Not LicenseNoWarning AndAlso LicensedTill > Date.Now AndAlso LicensedTill < Date.MaxValue Then
                    CheckLicenseExpiryWarnings(context)
                End If
                Return True
            End If

            ' Non-config license path: check My.Settings-based license
            If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue Then
                noLicenseConfigured = True
            ElseIf Date.Now > LicensedTill Then
                ' Check if we're within the grace period
                If GracePeriodDays > 0 AndAlso Date.Now <= LicensedTill.AddDays(GracePeriodDays) Then
                    ' Within grace period - show warning and allow continuation
                    CheckGracePeriodWarning(context, LicensedTill)
                    Return True
                End If
                licenseExpired = True
            End If

            ' Handle expired license (past grace period)
            If licenseExpired Then
                Dim msg = BuildLicenseMessage(
    $"Your license for {AN} for {context.RDV} has EXPIRED on {LicensedTill:d}." & vbCrLf & vbCrLf &
    "Would you like to update your license information now?")

                Dim result = ShowCustomYesNoBox(msg, "Update License", "Cancel", $"{AN} License")
                If result = 1 Then
                    If Not ShowLicenseEntryForm(context) Then
                        Return False
                    End If
                    ' Re-validate after form: check if new license is valid
                    If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue OrElse Date.Now > LicensedTill Then
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

            ' Handle no license configured
            If noLicenseConfigured Then
                Dim msg = BuildLicenseMessage(
    $"No valid license is configured for {AN} for {context.RDV}." & vbCrLf & vbCrLf &
    "Would you like to enter your license information now?")

                Dim result = ShowCustomYesNoBox(msg, "Enter License", "Cancel", $"{AN} License")
                If result = 1 Then
                    If Not ShowLicenseEntryForm(context) Then
                        Return False
                    End If
                    ' Re-validate after form: check if new license is valid
                    If LicensedTill = Date.MinValue OrElse LicensedTill = Date.MaxValue OrElse Date.Now > LicensedTill Then
                        Return False
                    End If
                Else
                    Return False
                End If
            End If

            ' Check for upcoming expiry warnings
            If Not LicenseNoWarning AndAlso LicensedTill > Date.Now AndAlso LicensedTill < Date.MaxValue Then
                CheckLicenseExpiryWarnings(context)
            End If

            Return True
        End Function

        ''' <summary>
        ''' Checks if a grace period warning should be shown and displays it.
        ''' Shows warning every GracePeriodWarningIntervals starts during the grace period.
        ''' Only shows if LicenseNoWarning is False.
        ''' </summary>
        Private Shared Sub CheckGracePeriodWarning(context As ISharedContext, expiredDate As Date)
            Try
                ' Skip warning if LicenseNoWarning is True
                If LicenseNoWarning Then Return

                ' Calculate remaining grace period days
                Dim gracePeriodEnd As Date = expiredDate.AddDays(GracePeriodDays)
                Dim remainingDays As Integer = CInt((gracePeriodEnd.Date - Date.Now.Date).TotalDays)

                ' Check if we should show the warning based on start count
                If ShouldShowGracePeriodWarning() Then
                    Dim msg = BuildLicenseMessage(
                $"Your license for {AN} for {context.RDV} EXPIRED on {expiredDate:d}." & vbCrLf & vbCrLf &
                $"You are currently in a {GracePeriodDays}-day grace period. " &
                $"The add-in will stop working in {remainingDays} day(s) on {gracePeriodEnd:d}." & vbCrLf & vbCrLf &
                If(LicenseFromConfig,
                   "Please contact your administrator to update the license configuration.",
                   "Would you like to update your license information now?"))

                    If LicenseFromConfig Then
                        ShowCustomMessageBox(msg, $"{AN} License Grace Period")
                    Else
                        Dim result = ShowCustomYesNoBox(msg, "Update License", "Later", $"{AN} License Grace Period")
                        If result = 1 Then
                            ShowLicenseEntryForm(context)
                        End If
                    End If

                    RecordGracePeriodWarningShown()
                End If

            Catch
                ' Fault tolerance - ignore warning check errors
            End Try
        End Sub


        ''' <summary>
        ''' Checks if grace period warning should be shown (every GracePeriodWarningIntervals starts)
        ''' </summary>
        Private Shared Function ShouldShowGracePeriodWarning() As Boolean
            Try
                ' Increment start count
                Dim startCount As Integer = 0
                Try
                    startCount = My.Settings.GracePeriodWarningStartcount + 1
                Catch
                    startCount = 1
                End Try

                My.Settings.GracePeriodWarningStartcount = startCount
                My.Settings.Save()

                ' Show warning if start count reaches the interval threshold
                Return startCount >= GracePeriodWarningIntervals

            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Records that grace period warning was shown (resets the counter)
        ''' </summary>
        Private Shared Sub RecordGracePeriodWarningShown()
            Try
                My.Settings.GracePeriodWarningStartCount = 0
                My.Settings.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Performs license check for Beta version
        ''' </summary>
        Private Shared Function PerformBetaLicenseCheck(context As ISharedContext) As Boolean
            ' Check if BetaUpgradeInstructions URL is available
            Dim upgradeAvailable = CheckUrlAvailable(BetaUpgradeInstructions)

            Debug.WriteLine("upgradeAvailable = " & upgradeAvailable)

            If upgradeAvailable And Not Date.Now > BetaEndDate Then
                ' Check if we should warn (every 3rd day)
                If ShouldShowBetaWarning() Then
                    Dim msg = BuildLicenseMessage(
                        $"The beta test for and your copy of {AN} ends on {BetaEndDate:d}." & vbCrLf & vbCrLf &
                        $"To continue using {AN}, upgrade now to the new General Audience or Preview version as per the upgrade instructions at {BetaUpgradeInstructions}." & vbCrLf & vbCrLf &
                        "Would you like to open the upgrade instructions page?")

                    Dim result = ShowCustomYesNoBox(msg, "Open Instructions", "Later", $"{AN} Beta Test")
                    If result = 1 Then
                        Try
                            Process.Start(New ProcessStartInfo(BetaUpgradeInstructions) With {.UseShellExecute = True})
                        Catch ex As Exception
                            ShowCustomMessageBox($"Could not open the upgrade instructions page: {ex.Message}", AN)
                        End Try
                    End If

                    RecordBetaWarningShown()
                End If
            End If

            ' Standard expiry check for beta
            If Date.Now > BetaEndDate Then
                If upgradeAvailable Then
                    Dim msg = BuildLicenseMessage(
                    $"Your beta test license for {AN} for {context.RDV} has EXPIRED on {BetaEndDate:d}. {If(Date.Now > LicensedTill, $"Therefore, your copy of {AN} no longer works.", "")}" & vbCrLf & vbCrLf &
                    $"To continue using {AN}, upgrade to the new General Audience or Preview version at {BetaUpgradeInstructions}." & vbCrLf & vbCrLf &
                    "Would you like to open the upgrade instructions page?")

                    Dim result = ShowCustomYesNoBox(msg, "Open Instructions", "Later", $"{AN} Beta Test")
                    If result = 1 Then
                        Try
                            Process.Start(New ProcessStartInfo(BetaUpgradeInstructions) With {.UseShellExecute = True})
                        Catch ex As Exception
                            ShowCustomMessageBox($"Could not open the upgrade instructions page: {ex.Message}", AN)
                        End Try
                    End If

                    Return Not Date.Now > LicensedTill
                Else
                    Dim msg = BuildLicenseMessage(
                    $"Your beta test license for {AN} for {context.RDV} has EXPIRED on {BetaEndDate:d}. {If(Date.Now > LicensedTill, $"Therefore, your copy of {AN} no longer works.", "")}" & vbCrLf & vbCrLf &
                    $"To continue using {AN}, upgrade to the new General Audience or Preview version at {NewHomeURL}.")
                    ShowCustomMessageBox(msg, $"{AN} Beta Test")

                    Return Not Date.Now > LicensedTill
                End If
            End If

            ' Upcoming expiry warnings for beta
            If Not LicenseNoWarning Then
                CheckLicenseExpiryWarnings(context)
            End If

            Return True
        End Function

        ''' <summary>
        ''' Checks license expiry warnings at 30, 15, 10, 5, 3, 1 days before expiry
        ''' </summary>
        Private Shared Sub CheckLicenseExpiryWarnings(context As ISharedContext)
            Try
                Dim daysUntilExpiry = CInt((LicensedTill.Date - Date.Now.Date).TotalDays)

                For Each warningDay In LicenseWarningDays
                    If daysUntilExpiry = warningDay Then
                        Dim msg = BuildLicenseMessage(
                            $"Your license for {AN} for {context.RDV} will EXPIRE in {daysUntilExpiry} day(s) " &
                            $"on {LicensedTill:d}." & vbCrLf & vbCrLf &
                            If(LicenseFromConfig,
                               "Your license is configured centrally. Contact your administrator to renew.",
                               $"Please update your license at {AN4} or contact your administrator. Updating the license information is possible via 'Settings', then 'About {AN}'." & vbCrLf & vbCrLf & "Would you like to update your license information now?"))

                        ' Only offer to update if license is NOT from config
                        If LicenseFromConfig Then
                            ShowCustomMessageBox(msg, $"{AN} License Warning")
                        Else
                            Dim result = ShowCustomYesNoBox(msg, "Update License", "Later", $"{AN} License Warning")
                            If result = 1 Then
                                ShowLicenseEntryForm(context)
                            End If
                        End If
                        Exit For
                    End If
                Next
            Catch
                ' Fault tolerance - ignore warning check errors
            End Try
        End Sub


        ''' <summary>
        ''' Shows the license entry form for users to select and configure their license
        ''' </summary>
        Public Shared Function ShowLicenseEntryForm(context As ISharedContext) As Boolean
            Try
                Dim versionDate = ParseVersionDateFromRDV(context.RDV)
                Dim licenseTypes = GetLicenseTypes(versionDate)

                Using form As New Form()
                    form.Text = $"{AN} License Configuration"
                    form.FormBorderStyle = FormBorderStyle.FixedDialog
                    form.StartPosition = FormStartPosition.CenterScreen
                    form.MaximizeBox = False
                    form.MinimizeBox = False
                    form.ShowInTaskbar = True
                    form.TopMost = True
                    form.Width = 500

                    ' Set icon
                    Try
                        Dim bmp As New System.Drawing.Bitmap(My.Resources.Red_Ink_Logo)
                        form.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
                    Catch
                    End Try

                    Dim font As New System.Drawing.Font("Segoe UI", 9.0F)
                    form.Font = font

                    Dim yPos = 15
                    Dim leftMargin = 15
                    Dim controlWidth = 450
                    Dim inputControlLeft = 180 ' Aligned position for input controls

                    ' Title label (not bold, 3 extra points before selector)
                    Dim titleLabel As New Label() With {
                        .Text = "Select your license type:",
                        .AutoSize = True,
                        .Location = New System.Drawing.Point(leftMargin, yPos),
                        .Font = New System.Drawing.Font("Segoe UI", 9.0F)
                    }
                    form.Controls.Add(titleLabel)
                    yPos += titleLabel.PreferredHeight + 3

                    ' License type ComboBox (single line, no dropdown list expanding)
                    Dim cboLicenseType As New ComboBox() With {
                        .Location = New System.Drawing.Point(leftMargin, yPos),
                        .Width = controlWidth,
                        .DropDownStyle = ComboBoxStyle.DropDownList
                    }
                    For Each lt In licenseTypes
                        cboLicenseType.Items.Add(lt.Name)
                    Next
                    form.Controls.Add(cboLicenseType)
                    yPos += cboLicenseType.Height + 8

                    ' Calculate max description height needed
                    Dim maxDescHeight = 0
                    Using g = form.CreateGraphics()
                        For Each lt In licenseTypes
                            Dim size = g.MeasureString(lt.Description, font, controlWidth)
                            maxDescHeight = Math.Max(maxDescHeight, CInt(Math.Ceiling(size.Height)))
                        Next
                    End Using

                    ' Description label (flat, no scrollbars, no 3D border, dark grey text)
                    Dim lblDescription As New Label() With {
                        .Location = New System.Drawing.Point(leftMargin, yPos),
                        .Width = controlWidth,
                        .Height = maxDescHeight + 10,
                        .ForeColor = System.Drawing.Color.FromArgb(96, 96, 96),
                        .BackColor = System.Drawing.SystemColors.Control,
                        .BorderStyle = BorderStyle.None
                    }
                    form.Controls.Add(lblDescription)
                    yPos += lblDescription.Height + 8

                    ' License end date label
                    Dim lblEndDate As New Label() With {
                        .Text = "License valid until:",
                        .AutoSize = True,
                        .Location = New System.Drawing.Point(leftMargin, yPos + 3)
                    }
                    form.Controls.Add(lblEndDate)

                    Dim dtpEndDate As New DateTimePicker() With {
                        .Format = DateTimePickerFormat.Short,
                        .Location = New System.Drawing.Point(inputControlLeft, yPos),
                        .Width = 150
                    }
                    form.Controls.Add(dtpEndDate)
                    yPos += dtpEndDate.Height + 10

                    ' Number of users label
                    Dim lblUsers As New Label() With {
                        .Text = "Number of users:",
                        .AutoSize = True,
                        .Location = New System.Drawing.Point(leftMargin, yPos + 3)
                    }
                    form.Controls.Add(lblUsers)

                    Dim nudUsers As New NumericUpDown() With {
                        .Minimum = 1,
                        .Maximum = 10000,
                        .Value = 1,
                        .Location = New System.Drawing.Point(inputControlLeft, yPos),
                        .Width = 80
                    }
                    form.Controls.Add(nudUsers)
                    yPos += nudUsers.Height + 25

                    ' Buttons with proper padding
                    Dim btnSave As New Button() With {
                        .Text = "Save License",
                        .Location = New System.Drawing.Point(leftMargin, yPos),
                        .AutoSize = True,
                        .Padding = New Padding(10, 5, 10, 5)
                    }
                    form.Controls.Add(btnSave)

                    Dim btnCancel As New Button() With {
                        .Text = "Cancel",
                        .Location = New System.Drawing.Point(btnSave.Right + 10, yPos),
                        .AutoSize = True,
                        .Padding = New Padding(10, 5, 10, 5)
                    }
                    form.Controls.Add(btnCancel)

                    ' Set form height based on content plus adequate bottom margin
                    form.ClientSize = New System.Drawing.Size(form.ClientSize.Width, yPos + btnSave.Height + 25)

                    Dim result As Boolean = False

                    ' Update UI based on license type selection
                    AddHandler cboLicenseType.SelectedIndexChanged, Sub(s, e)
                                                                        If cboLicenseType.SelectedIndex >= 0 AndAlso cboLicenseType.SelectedIndex < licenseTypes.Count Then
                                                                            Dim selectedType = licenseTypes(cboLicenseType.SelectedIndex)
                                                                            lblDescription.Text = selectedType.Description

                                                                            ' End date
                                                                            If selectedType.FixedEndDate.HasValue Then
                                                                                dtpEndDate.Value = selectedType.FixedEndDate.Value
                                                                                dtpEndDate.Enabled = False
                                                                            ElseIf selectedType.DefaultEndDate.HasValue Then
                                                                                dtpEndDate.Value = selectedType.DefaultEndDate.Value
                                                                                dtpEndDate.Enabled = True
                                                                            Else
                                                                                dtpEndDate.Value = Date.Now.AddYears(1)
                                                                                dtpEndDate.Enabled = True
                                                                            End If

                                                                            ' Users
                                                                            If selectedType.FixedUsers.HasValue Then
                                                                                nudUsers.Value = selectedType.FixedUsers.Value
                                                                                nudUsers.Enabled = False
                                                                            ElseIf selectedType.DefaultUsers.HasValue Then
                                                                                nudUsers.Value = selectedType.DefaultUsers.Value
                                                                                nudUsers.Enabled = True
                                                                            Else
                                                                                nudUsers.Enabled = True
                                                                            End If
                                                                        End If
                                                                    End Sub

                    ' Save handler
                    AddHandler btnSave.Click, Sub(s, e)
                                                  Try
                                                      ' Validation
                                                      If cboLicenseType.SelectedIndex < 0 Then
                                                          ShowCustomMessageBox("Please select a license type.", AN)
                                                          Return
                                                      End If

                                                      Dim endDate = dtpEndDate.Value.Date
                                                      Dim users = CInt(nudUsers.Value)

                                                      If users <= 0 Then
                                                          ShowCustomMessageBox("Number of users must be at least 1.", AN)
                                                          Return
                                                      End If

                                                      If endDate <= Date.Now.Date Then
                                                          ShowCustomMessageBox("License end date must be in the future.", AN)
                                                          Return
                                                      End If

                                                      If endDate > Date.Now.AddYears(MaxLicenseYearsInFuture) Then
                                                          ShowCustomMessageBox($"License end date cannot be more than {MaxLicenseYearsInFuture} years in the future.", AN)
                                                          Return
                                                      End If

                                                      ' Save to My.Settings
                                                      LicenseStatus = licenseTypes(cboLicenseType.SelectedIndex).Name
                                                      LicenseUsers = users
                                                      LicensedTill = endDate

                                                      My.Settings.LicenseStatus = LicenseStatus
                                                      My.Settings.LicenseUsers = LicenseUsers
                                                      My.Settings.LicensedTill = LicensedTill
                                                      My.Settings.Save()

                                                      result = True
                                                      form.Close()
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox($"Error saving license: {ex.Message}", AN)
                                                  End Try
                                              End Sub

                    AddHandler btnCancel.Click, Sub(s, e)
                                                    form.Close()
                                                End Sub

                    ' Set initial selection if we have a stored license
                    If Not String.IsNullOrEmpty(LicenseStatus) Then
                        For i = 0 To licenseTypes.Count - 1
                            If licenseTypes(i).Name.Equals(LicenseStatus, StringComparison.OrdinalIgnoreCase) Then
                                cboLicenseType.SelectedIndex = i
                                If LicensedTill > Date.MinValue AndAlso LicensedTill < Date.MaxValue Then
                                    Try : dtpEndDate.Value = LicensedTill : Catch : End Try
                                End If
                                nudUsers.Value = Math.Max(1, LicenseUsers)
                                Exit For
                            End If
                        Next
                    End If

                    form.ShowDialog()
                    Return result
                End Using

            Catch ex As Exception
                ShowCustomMessageBox($"Error showing license form: {ex.Message}", AN)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Parses the version date from context.RDV (e.g., "Word (V.101225 Gen2 Beta Test)" -> December 10, 2025)
        ''' </summary>
        Private Shared Function ParseVersionDateFromRDV(rdv As String) As Date
            Try
                ' Find "V." followed by 6 digits (DDMMYY format)
                Dim vIndex = rdv.IndexOf("V.", StringComparison.OrdinalIgnoreCase)
                If vIndex >= 0 AndAlso rdv.Length >= vIndex + 8 Then
                    Dim dateStr = rdv.Substring(vIndex + 2, 6)
                    If dateStr.All(AddressOf Char.IsDigit) Then
                        Dim day = Integer.Parse(dateStr.Substring(0, 2))
                        Dim month = Integer.Parse(dateStr.Substring(2, 2))
                        Dim year = 2000 + Integer.Parse(dateStr.Substring(4, 2))
                        Return New Date(year, month, day)
                    End If
                End If
            Catch
            End Try
            ' Default to current date if parsing fails
            Return Date.Now
        End Function

        ''' <summary>
        ''' Builds a license message with optional contact information appended
        ''' </summary>
        Private Shared Function BuildLicenseMessage(baseMessage As String) As String
            If Not String.IsNullOrEmpty(LicenseContact) Then
                Return baseMessage & vbCrLf & vbCrLf & LicenseContact
            End If
            Return baseMessage
        End Function


        ''' <summary>
        ''' Checks if a URL is available (follows redirects, only accepts 2xx as available)
        ''' </summary>
        Private Shared Function CheckUrlAvailable(url As String) As Boolean
            Try
                ' Enable TLS 1.2 and TLS 1.3 for HTTPS connections
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

                Dim request = DirectCast(WebRequest.Create(url), HttpWebRequest)
                request.Method = "HEAD"
                request.Timeout = 5000
                request.AllowAutoRedirect = True
                request.MaximumAutomaticRedirections = 5
                request.UserAgent = $"{AN}/1.0"

                Try
                    Using response = DirectCast(request.GetResponse(), HttpWebResponse)
                        Dim statusCode = CInt(response.StatusCode)
                        Debug.WriteLine($"{url} returned {response.StatusCode} ({statusCode})")
                        ' Accept only 2xx status codes as available
                        Return statusCode >= 200 AndAlso statusCode < 300
                    End Using
                Catch webEx As WebException
                    ' Check if this is a 4xx/5xx error - if so, return False immediately
                    If webEx.Response IsNot Nothing Then
                        Dim httpResponse = TryCast(webEx.Response, HttpWebResponse)
                        If httpResponse IsNot Nothing Then
                            Dim statusCode = CInt(httpResponse.StatusCode)
                            Debug.WriteLine($"{url} HEAD failed with {httpResponse.StatusCode} ({statusCode})")
                            httpResponse.Close()

                            ' 4xx (client errors including 404) and 5xx (server errors) mean URL is not available
                            If statusCode >= 400 Then
                                Return False
                            End If
                        End If
                    End If

                    ' HEAD might not be supported, retry with GET method
                    Dim getRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
                    getRequest.Method = "GET"
                    getRequest.Timeout = 5000
                    getRequest.AllowAutoRedirect = True
                    getRequest.MaximumAutomaticRedirections = 5
                    getRequest.UserAgent = $"{AN}/1.0"

                    Using response = DirectCast(getRequest.GetResponse(), HttpWebResponse)
                        Dim statusCode = CInt(response.StatusCode)
                        Debug.WriteLine($"{url} GET returned {response.StatusCode} ({statusCode})")
                        ' Accept only 2xx status codes as available
                        Return statusCode >= 200 AndAlso statusCode < 300
                    End Using
                End Try

            Catch ex As WebException
                ' Handle any remaining WebExceptions (including 404, 500, etc.)
                If ex.Response IsNot Nothing Then
                    Dim httpResponse = TryCast(ex.Response, HttpWebResponse)
                    If httpResponse IsNot Nothing Then
                        Dim statusCode = CInt(httpResponse.StatusCode)
                        Debug.WriteLine($"{url} failed with {httpResponse.StatusCode} ({statusCode})")
                        httpResponse.Close()
                        ' Only 2xx means available - everything else (including 404) is not available
                        Return False
                    End If
                End If
                Debug.WriteLine($"{url} failed: {ex.Message}")
                Return False
            Catch ex As Exception
                Debug.WriteLine($"{url} failed: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Checks if beta warning should be shown (every 3 days OR every BetaWarningInterval starts, whichever is sooner)
        ''' </summary>
        Private Shared Function ShouldShowBetaWarning() As Boolean
            Try
                ' Check days since last warning
                Dim lastWarning = My.Settings.LastBetaWarningDate
                Dim daysSinceLastWarning = If(lastWarning = Date.MinValue, Integer.MaxValue, CInt((Date.Now.Date - lastWarning.Date).TotalDays))

                ' Increment and check start count since last warning
                Dim startCount = My.Settings.BetaWarningStartCount + 1
                My.Settings.BetaWarningStartCount = startCount
                My.Settings.Save()

                ' Show warning if either condition is met:
                ' - First time ever (no previous warning date)
                ' - BetaWarningDays or more days have passed since last warning
                ' - BetaWarningInterval or more starts since last warning reset
                Return daysSinceLastWarning >= BetaWarningDays OrElse startCount >= BetaWarningInterval

            Catch
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Records that beta warning was shown (resets both date and counter)
        ''' </summary>
        Private Shared Sub RecordBetaWarningShown()
            Try
                My.Settings.LastBetaWarningDate = Date.Now.Date
                My.Settings.BetaWarningStartCount = 0
                My.Settings.Save()
            Catch
            End Try
        End Sub

    End Class
End Namespace