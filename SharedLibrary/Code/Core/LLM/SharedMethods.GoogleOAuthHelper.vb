' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.Text
Imports System.IO
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Class GoogleOAuthHelper
            ' Public variables
            Public Shared client_email As String = ""
            Public Shared private_key As String = ""
            Public Shared scopes As String = ""
            Public Shared token_uri As String = ""
            Public Shared token_life As Long = 0

            ' Base64Url encoding
            Private Shared Function Base64UrlEncode(input As String) As String
                Return System.Convert.ToBase64String(Encoding.UTF8.GetBytes(input)).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            Private Shared Function Base64UrlEncode(inputBytes As Byte()) As String
                Return System.Convert.ToBase64String(inputBytes).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ' Sign data using BouncyCastle
            Private Shared Function SignData(data As Byte()) As Byte()
                Dim rsaKey As RsaPrivateCrtKeyParameters
                Dim formattedPrivateKey As String = private_key.Replace("\n", Environment.NewLine)

                Using reader As New StringReader(formattedPrivateKey)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    rsaKey = DirectCast(pemReader.ReadObject(), RsaPrivateCrtKeyParameters)
                End Using

                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, rsaKey)
                signer.BlockUpdate(data, 0, data.Length)
                Return signer.GenerateSignature()
            End Function

            ' Generate JWT
            Public Shared Function GenerateJWT() As String
                Dim issuedAt As Long = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
                Dim expiry As Long = issuedAt + 3600 ' 1 hour expiry

                Dim header = New With {.alg = "RS256", .typ = "JWT"}
                Dim payload = New With {
                                        .iss = client_email,
                                        .scope = scopes,
                                        .aud = token_uri,
                                        .exp = expiry,
                                        .iat = issuedAt
                                    }

                Dim headerBase64 = Base64UrlEncode(JsonConvert.SerializeObject(header))
                Dim payloadBase64 = Base64UrlEncode(JsonConvert.SerializeObject(payload))
                Dim unsignedToken = $"{headerBase64}.{payloadBase64}"
                Dim signature = SignData(Encoding.UTF8.GetBytes(unsignedToken))
                Dim signatureBase64 = Base64UrlEncode(signature)

                Return $"{unsignedToken}.{signatureBase64}"
            End Function

            ' Get Access Token
            Public Shared Async Function GetAccessToken() As Task(Of String)
                Dim jwt = GenerateJWT()
                Dim requestBody As String = JsonConvert.SerializeObject(New With {
                                .grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                                .assertion = jwt
                                            })

                Using client As New HttpClient()
                    Dim content = New StringContent(requestBody, Encoding.UTF8, "application/json")
                    Dim response = Await client.PostAsync(token_uri, content)

                    If response.IsSuccessStatusCode Then
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                        Dim tokenData = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(responseBody)
                        Return tokenData("access_token")
                    Else
                        ShowCustomMessageBox($"Error getting access token: {response.ReasonPhrase}")
                        Return ""
                    End If
                End Using
            End Function
        End Class

    End Class

End Namespace