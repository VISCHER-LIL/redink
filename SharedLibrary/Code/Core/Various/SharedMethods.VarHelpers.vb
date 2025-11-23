' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Management
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Win32


Namespace SharedLibrary
    Partial Public Class SharedMethods


        Public Shared Function GetDefaultINIPath(ByVal key As String) As String

            For Each entry In DefaultINIPaths
                If key.Contains(entry.Key) Then
                    Return ExpandEnvironmentVariables(entry.Value)
                End If
            Next
            Return ExpandEnvironmentVariables(DefaultINIPaths.Values.First())
        End Function


        Public Shared Function RenameFileToBak(filePath As String) As Boolean
            Try
                ' Rename the file to a .bak file
                Dim bakFilePath As String = filePath & ".bak"
                If File.Exists(bakFilePath) Then
                    File.Delete(bakFilePath)
                End If
                File.Move(filePath, bakFilePath)
                Return True
            Catch ex As Exception
                MessageBox.Show($"Error renaming file to .bak: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try

        End Function


        Public Shared Sub WriteToRegistry(ByVal regPath As String, ByVal regValue As String)
            Try
                ' Remove carriage returns from the value
                regValue = RemoveCR(regValue)

                ' Split the registry path into hive and subkey
                Dim hiveName As String = regPath.Split("\"c)(0)
                Dim subKeyPath As String = String.Join("\", regPath.Split("\"c).Skip(1))

                Dim registryHive As RegistryKey

                ' Determine the appropriate registry hive
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        registryHive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        registryHive = Registry.LocalMachine
                    Case Else
                        Throw New ArgumentException("Unsupported registry hive: " & hiveName)
                End Select

                ' Write the value to the registry
                Using subKey As RegistryKey = registryHive.CreateSubKey(subKeyPath, True)
                    If subKey Is Nothing Then
                        Throw New Exception("Unable to open or create the registry key at: " & regPath)
                    End If
                    subKey.SetValue("", regValue, RegistryValueKind.String)
                End Using

                ShowCustomMessageBox($"Written value '{regValue}' to the registry at '{regPath}.'")

            Catch ex As Exception
                MessageBox.Show($"Error: Unable to write to the registry at '{regPath}'. {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Public Shared Function GetFromRegistry(registryPath As String, valueName As String, Optional suppressErrors As Boolean = False) As String
            Try
                ' Split the registry path into hive and subkey
                Dim hiveName As String = registryPath.Split("\"c)(0)
                Dim subKeyPath As String = registryPath.Substring(hiveName.Length + 1)

                ' Determine the registry hive
                Dim hive As RegistryKey = Nothing
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        hive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        hive = Registry.LocalMachine
                    Case "HKEY_CLASSES_ROOT"
                        hive = Registry.ClassesRoot
                    Case "HKEY_USERS"
                        hive = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        hive = Registry.CurrentConfig
                    Case Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - invalid registry hive: " & hiveName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                End Select

                ' Open the subkey and retrieve the value
                Using subKey As RegistryKey = hive.OpenSubKey(subKeyPath)
                    If subKey IsNot Nothing Then
                        Return RemoveCR(subKey.GetValue(valueName, Nothing)?.ToString())
                    Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - Registry key not found: " & subKeyPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                    End If
                End Using

            Catch ex As System.Exception
                If Not suppressErrors Then
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return ""
            End Try
        End Function

        Public Shared Function RemoveCR(ByVal inputtext As String) As String
            If inputtext IsNot Nothing Then
                inputtext = inputtext.Trim()
                inputtext = inputtext.Replace(vbCr, "")
                inputtext = inputtext.Replace(vbLf, "")
                inputtext = inputtext.Replace(vbCrLf, "")
                inputtext = inputtext.Trim()
            Else
                inputtext = ""
            End If
            Return inputtext
        End Function

        Public Shared Function IsEmptyOrBlank(ByVal str As String) As Boolean
            ' Check if the string is empty or consists only of whitespace
            Return String.IsNullOrWhiteSpace(str)
        End Function

        Public Shared Function ExpandEnvironmentVariables(ByVal filePath As String) As String
            ' Start with the input path
            Dim expandedPath As String = Environment.ExpandEnvironmentVariables(filePath)

            Try

                ' Remove any preceding and trailing quotation marks
                expandedPath = expandedPath.Trim(""""c)

                ' Expand known variables using Environment.GetEnvironmentVariable and ensure proper path format
                expandedPath = Regex.Replace(expandedPath, "%APPDATA%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%USERPROFILE%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%WINDIR%", Path.Combine(Environment.GetEnvironmentVariable("WINDIR")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%TEMP%", Path.Combine(Path.GetTempPath()), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%HOMEPATH%", Path.Combine(Environment.GetEnvironmentVariable("HOMEPATH")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%APPSTARTUPPATH%", Path.Combine(System.Windows.Forms.Application.StartupPath), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%DESKTOP%", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), RegexOptions.IgnoreCase)

                ' Clean up any potential double backslashes
                expandedPath = Regex.Replace(expandedPath, "\\{2,}", "\\")

                ' Return the expanded path
                If expandedPath = "" Then Return "" Else Return Path.GetFullPath(expandedPath)

            Catch ex As System.Exception
                ' Return Nothing on failure
                Return ""
            End Try

        End Function



        Public Shared Function DecodeBase64(ByVal base64String As String) As Byte()
            Try
                ' Normalize the input: remove whitespaces and line breaks
                base64String = base64String.Replace(vbCrLf, "").Replace(vbLf, "").Replace(vbCr, "").Replace(" ", "")

                ' Convert URL-safe Base64 to standard Base64 if input is URL-safe
                base64String = base64String.Replace("-", "+").Replace("_", "/")

                ' Add padding
                While (base64String.Length Mod 4) <> 0
                    base64String &= "="
                End While

                ' Decode the Base64 string
                Return System.Convert.FromBase64String(base64String)
            Catch ex As System.Exception
                ' Return Nothing on failure
                Return Nothing
            End Try
        End Function

        Public Shared Function DecodeString(ByVal encodedText As String, ByVal pTerm As String) As String
            ' Remove literal "\n" if present
            encodedText = encodedText.Replace("\n", "")
            ' Also ensure actual newline characters are removed
            encodedText = encodedText.Replace(vbCr, "").Replace(vbLf, "")
            ' Remove spaces if any
            encodedText = encodedText.Replace(" ", "")

            Dim encryptedBytes As Byte() = DecodeBase64(encodedText)
            If encryptedBytes Is Nothing Then
                Return "Error: Invalid Base64 input"
            End If

            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            Dim decryptedBytes(encryptedBytes.Length - 1) As Byte

            For i As Integer = 0 To encryptedBytes.Length - 1
                decryptedBytes(i) = encryptedBytes(i) Xor pTermBytes(i Mod pTermBytes.Length)
            Next

            ' Convert decrypted bytes to string
            ' If UTF8 fails due to unexpected characters, try ASCII or verify the original encoding.
            Try
                Return System.Text.Encoding.UTF8.GetString(decryptedBytes)
            Catch
                Return System.Text.Encoding.ASCII.GetString(decryptedBytes)
            End Try
        End Function

        Public Shared Function CodeString(ByVal inputText As String, ByVal pTerm As String) As String
            Dim inputBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(inputText)
            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            Dim encryptedBytes(inputBytes.Length - 1) As Byte

            Dim inputLength As Integer = inputBytes.Length
            Dim pTermLength As Integer = pTermBytes.Length

            ' Encrypt each byte with XOR operation
            For i As Integer = 0 To inputBytes.Length - 1
                encryptedBytes(i) = inputBytes(i) Xor pTermBytes(i Mod pTermLength)
            Next

            ' Convert encrypted bytes to Base64
            Return System.Convert.ToBase64String(encryptedBytes)
        End Function

        Public Shared Function GetDomain() As String
            Try
                ' Initialize a WMI query to get the Domain property from Win32_ComputerSystem
                Dim searcher As New ManagementObjectSearcher("SELECT Domain FROM Win32_ComputerSystem")
                Dim strDomain As String = String.Empty

                ' Execute the query and retrieve the result
                For Each queryObj As ManagementObject In searcher.Get()
                    If queryObj("Domain") IsNot Nothing Then
                        strDomain = queryObj("Domain").ToString()
                    End If
                Next

                ' If the domain is not retrieved, return an appropriate message
                If String.IsNullOrEmpty(strDomain) Then
                    MessageBox.Show($"Error in GetDomain - unable to determine the domain name or workgroup.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    strDomain = ""
                End If

                Return strDomain
            Catch ex As System.Exception
                MessageBox.Show($"Error in GetDomain - Error retrieving domain or workgroup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        Public Shared Function WrongDomain() As Boolean

            Dim strDomain As String = GetDomain() ' Current domain of the computer
            Dim domainList() As String
            Dim domainFound As Boolean = False

            If Not String.IsNullOrEmpty(alloweddomains) Then
                ' Convert the list of allowed domains into an array
                domainList = alloweddomains.Split(","c)

                ' Check if the current domain is in the allowed list
                For Each domain In domainList
                    If strDomain.Equals(domain.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        domainFound = True
                        Exit For
                    End If
                Next

                ' If the domain is not in the list of allowed domains
                If Not domainFound Then
                    ShowCustomMessageBox($"This copy of {AN} may not be executed in this network environment (which is '{strDomain}'). The domain has to be added to the code by your administrator.")
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function

    End Class

End Namespace
