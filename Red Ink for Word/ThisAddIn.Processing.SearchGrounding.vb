' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.SearchGrounding.vb
' Purpose: Collects LLM-driven search terms, performs configured internet searches, and enriches prompts with crawled content.
'
' Workflow:
'  - ConsultInternet: Builds search prompts, optionally requests approval, runs searches, and prepares follow-up prompts.
'  - PerformSearchGrounding: Executes the HTTP search, extracts unique URLs, retrieves qualifying content, and aggregates it.
'  - RetrieveWebsiteContent/CrawlWebsite: Crawl discovered pages within bounded depth, timeout, and error limits while harvesting paragraph text.
'  - CrawlContext/GetAbsoluteUrl: Maintain crawl state and normalize relative links.
'
' External Dependencies:
'  - SharedLibrary.SharedLibrary.SharedMethods for interpolation, UI dialogues, and LLM access.
'  - HtmlAgilityPack for HTML parsing.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports HtmlAgilityPack
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides search-grounding helpers that connect Word instructions with external internet content.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>
    ''' Coordinates LLM-based search-term generation, optional approval, remote search execution, and prompt selection.
    ''' </summary>
    ''' <param name="DoMarkup">Determines whether the markup-specific system prompt is used when selected text exists.</param>
    ''' <returns>True when search preparation completes successfully; otherwise False.</returns>
    Public Async Function ConsultInternet(DoMarkup As Boolean) As Task(Of Boolean)

        Try

            InfoBox.ShowInfoBox("Asking the LLM to determine the necessary searchterms for your instruction ...")

            Dim SysPromptTemp As String
            Dim SearchResults As List(Of String)

            CurrentDate = DateAndTime.Now.ToString("MMMM d, yyyy")

            SysPromptTemp = InterpolateAtRuntime(INI_ISearch_SearchTerm_SP)

            SearchTerms = Await LLM(SysPromptTemp, If(SelectedText = "", "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0)

            If String.IsNullOrWhiteSpace(SearchTerms) Then
                InfoBox.ShowInfoBox("")
                ShowCustomMessageBox("The LLM failed to establish searchterms. Will abort.")
                Return False
            End If

            If INI_ISearch_Approve Then
                InfoBox.ShowInfoBox("")
                Dim approveresult As Integer = ShowCustomYesNoBox("These are the searchterms that the LLM wants to issue to " & INI_ISearch_Name & ": {SearchTerms}", "Approve", "Abort", $"{AN} Internet Search", 5, " = 'Approve'")
                If approveresult = 0 Or approveresult = 2 Then Return False
            End If

            InfoBox.ShowInfoBox($"Now using {INI_ISearch_Name} to search for '{SearchTerms}' ...")

            SearchResults = Await PerformSearchGrounding(SearchTerms, INI_ISearch_URL, INI_ISearch_ResponseMask1, INI_ISearch_ResponseMask2, INI_ISearch_Tries, INI_ISearch_MaxDepth)

            SearchResult = String.Join(Environment.NewLine, SearchResults.Select(Function(result, index) $"<SEARCHRESULT{index + 1}>{result}</SEARCHRESULT{index + 1}>"))

            InfoBox.ShowInfoBox($"Having the LLM execute your instruction using also the {SearchResults.Count} result(s) from the Internet search ...", 3)
            If DoMarkup And Not String.IsNullOrWhiteSpace(SelectedText) Then
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP_Markup)
            Else
                SysPrompt = InterpolateAtRuntime(INI_ISearch_Apply_SP)
            End If

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Error in ConsultInternet: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Executes the configured internet search, extracts unique URLs using response masks, retrieves qualifying content, and returns the collected snippets.
    ''' </summary>
    ''' <param name="SGTerms">Search expression provided to the internet search endpoint.</param>
    ''' <param name="ISearch_URL">Base search URL that precedes the encoded search terms.</param>
    ''' <param name="ISearch_ResponseMask1">Start delimiter used to locate URLs in the response.</param>
    ''' <param name="ISearch_ResponseMask2">End delimiter used to locate URLs in the response.</param>
    ''' <param name="ISearch_Tries">Maximum number of URLs that will be processed.</param>
    ''' <param name="ISearch_MaxDepth">Maximum crawl depth per retrieved URL.</param>
    ''' <returns>List of plain-text contents harvested from the visited URLs.</returns>
    Public Async Function PerformSearchGrounding(SGTerms As String, ISearch_URL As String, ISearch_ResponseMask1 As String, ISearch_ResponseMask2 As String, ISearch_Tries As Integer, ISearch_MaxDepth As Integer) As Task(Of List(Of String))
        Dim results As New List(Of String)
        Using httpClient As New HttpClient()
            Try
                ' Construct the search URL
                Dim searchUrl As String = ISearch_URL & Uri.EscapeDataString(SGTerms)

                InfoBox.ShowInfoBox($"Searching {searchUrl} ...")

                ' Get search results HTML
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30) ' Set to an appropriate value
                Dim searchResponse As String = Await httpClient.GetStringAsync(searchUrl)
                'Debug.WriteLine("Search response: " & Left(searchResponse, 10000))

                InfoBox.ShowInfoBox($"Extracting URLs ...")

                ' Extract URLs using the defined start and mask
                Dim urlPattern As String = Regex.Escape(ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(ISearch_ResponseMask2)
                Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

                Dim extractedUrls As New List(Of String)
                Dim URLList As String = "URLS found so far:" & vbCrLf & vbCrLf
                For Each match As Match In matches
                    Dim rawUrl As String = match.Groups(1).Value
                    Dim decodedUrl As String = WebUtility.UrlDecode(rawUrl.Replace(ISearch_ResponseMask1, ""))

                    ' Check if the decoded URL already exists in the list
                    If Not extractedUrls.Contains(decodedUrl) Then
                        extractedUrls.Add(decodedUrl)
                        URLList += decodedUrl & vbCrLf
                        InfoBox.ShowInfoBox(URLList)
                        'Debug.WriteLine("URL added: " & decodedUrl)
                    Else
                        'Debug.WriteLine("Duplicate URL skipped: " & decodedUrl)
                    End If

                    If extractedUrls.Count >= ISearch_Tries Then Exit For
                Next

                ' Visit each extracted URL and retrieve content
                For Each url In extractedUrls
                    Try
                        Dim content As String = Await RetrieveWebsiteContent(url, ISearch_MaxDepth, httpClient)
                        'Debug.WriteLine("URL {url} provides:" & content)
                        If Not String.IsNullOrWhiteSpace(content) Then
                            If Len(content) > ISearch_MinChars Then
                                results.Add(content)
                                InfoBox.ShowInfoBox($"{url} resulted in: " & Left(content.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, ""), 1000))
                                'Debug.WriteLine("Content=" & content)
                            Else
                                'Debug.WriteLine("Content (not considered)=" & content)
                            End If
                        End If
                    Catch ex As Exception
                        'Debug.WriteLine($"Error retrieving content from URL: {url} - {ex.Message}")
                    End Try
                Next

            Catch ex As HttpRequestException
                'Debug.WriteLine($"HTTP Request Error: {ex.Message}")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (HTTP request error: " & ex.Message & ")")
            Catch ex As TaskCanceledException
                'Debug.WriteLine("Request timed out or was canceled.")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (request timed-out or was canceled: " & ex.Message & ")")
            Catch ex As Exception
                'Debug.WriteLine($"An error occurred: {ex.Message}")
                ShowCustomMessageBox("An error occurred when searching and analyzing the Internet (" & ex.Message & ")")
            Finally
                httpClient.Dispose()
                InfoBox.ShowInfoBox("")
            End Try
        End Using
        Return results
    End Function

    ''' <summary>
    ''' Crawls a base URL up to the specified depth, captures paragraph text, strips HTML, and enforces timeout plus error limits.
    ''' </summary>
    ''' <param name="baseUrl">Starting URL that seeds the crawl.</param>
    ''' <param name="subTries">Maximum link depth to follow.</param>
    ''' <param name="httpClient">HttpClient instance reused for the crawl.</param>
    ''' <returns>Plain-text content (limited to ISearch_MaxChars) harvested from the crawl.</returns>
    Private Async Function RetrieveWebsiteContent(
                        baseUrl As String,
                        subTries As Integer,
                        httpClient As HttpClient
                    ) As Task(Of String)

        ' Use the shared HttpClient instance provided by the caller for the entire crawl
        Dim client As HttpClient = httpClient

        ' Create the shared context object
        Dim context As New CrawlContext With {
                    .VisitedUrls = New HashSet(Of String)(),
                    .ContentBuilder = New StringBuilder(),
                    .ErrorCount = 0,
                    .MaxErrors = ISearch_MaxCrawlErrors
                }

        ' Create one CancellationTokenSource for the entire crawl duration
        Using cts As New CancellationTokenSource(TimeSpan.FromSeconds(INI_ISearch_Timeout))
            ' Call the CrawlWebsite function with the context
            '   - 'subTries' is your maxDepth
            '   - '0' is your currentDepth

            Await CrawlWebsite(
                        currentUrl:=baseUrl,
                        maxDepth:=subTries,
                        currentDepth:=0,
                        httpClient:=client,
                        context:=context,
                        cancellationToken:=cts.Token,
                        timeOutSeconds:=CInt(INI_ISearch_Timeout)
                         )

        End Using

        ' Return plain text with HTML tags removed (up to ISearch_MaxChars)
        Return Left(
                    Regex.Replace(context.ContentBuilder.ToString(), "<.*?>", String.Empty).Trim(),
                    ISearch_MaxChars
                    )
    End Function

    ''' <summary>
    ''' Holds crawl state across recursive invocations so that visited urls, aggregated content, and error limits stay consistent.
    ''' </summary>
    Public Class CrawlContext
        Public Property VisitedUrls As HashSet(Of String)
        Public Property ContentBuilder As StringBuilder
        Public Property ErrorCount As Integer
        Public Property MaxErrors As Integer
    End Class

    ''' <summary>
    ''' Recursively crawls a URL, collecting paragraph text and following links while honoring depth, timeout, and error thresholds.
    ''' </summary>
    ''' <param name="currentUrl">The URL currently being crawled.</param>
    ''' <param name="maxDepth">Maximum depth allowed for recursion.</param>
    ''' <param name="currentDepth">Current recursion depth.</param>
    ''' <param name="httpClient">HttpClient used for fetching page content.</param>
    ''' <param name="context">Shared crawl context for deduplication and aggregation.</param>
    ''' <param name="cancellationToken">Cancellation token controlling crawl timeout.</param>
    ''' <param name="timeOutSeconds">Timeout used when the caller does not supply a token.</param>
    ''' <returns>Empty string because content is accumulated through the shared context.</returns>
    Private Async Function CrawlWebsite(
    currentUrl As String,
    maxDepth As Integer,
    currentDepth As Integer,
    httpClient As HttpClient,
    context As CrawlContext,
    Optional cancellationToken As CancellationToken = Nothing,
    Optional timeOutSeconds As Integer = 10
) As Task(Of String)

        ' If the function has no valid CancellationToken, create one that cancels after 30 seconds
        Dim localCts As CancellationTokenSource = Nothing
        If cancellationToken = CancellationToken.None Then
            localCts = New CancellationTokenSource(TimeSpan.FromSeconds(timeOutSeconds))
            cancellationToken = localCts.Token
        End If

        Dim results As String = ""

        ' If we've already exceeded the max errors, abort quickly
        If context.ErrorCount >= context.MaxErrors Then
            Return results
        End If

        ' Early exit if depth is exceeded or already visited
        If currentDepth > maxDepth OrElse context.VisitedUrls.Contains(currentUrl) Then
            Return results
        End If

        Try
            context.VisitedUrls.Add(currentUrl)

            ' Use the cancellation token to abort if it exceeds the specified time
            Dim response As HttpResponseMessage = Await httpClient.GetAsync(currentUrl, cancellationToken)
            Dim pageHtml As String = Await response.Content.ReadAsStringAsync()

            Dim doc As New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(pageHtml)

            ' Safely extract paragraph text
            Dim pNodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p")
            If pNodes IsNot Nothing Then
                For Each node In pNodes
                    context.ContentBuilder.AppendLine(node.InnerText.Trim())
                Next
            End If

            ' Follow links if depth permits
            If currentDepth < maxDepth Then
                Dim links As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a[@href]")
                If links IsNot Nothing Then
                    For Each link In links
                        Dim hrefValue As String = link.GetAttributeValue("href", "").Trim()
                        Dim absoluteUrl As String = GetAbsoluteUrl(currentUrl, hrefValue)
                        ' You should already have a GetAbsoluteUrl function that resolves relative paths

                        If Not String.IsNullOrEmpty(absoluteUrl) Then
                            Await CrawlWebsite(
                            absoluteUrl,
                            maxDepth,
                            currentDepth + 1,
                            httpClient,
                            context,
                            cancellationToken,
                            timeOutSeconds
                        )

                            ' If error count has now exceeded the limit, stop immediately
                            If context.ErrorCount >= context.MaxErrors Then
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

        Catch ex As System.Threading.Tasks.TaskCanceledException
            ' Decide if a cancellation/timeout should increment errorCount
            context.ErrorCount += 1
            Debug.WriteLine($"Task canceled while crawling URL: {currentUrl} - {ex.Message}")

        Catch ex As System.Exception
            context.ErrorCount += 1
            Debug.WriteLine($"Error crawling URL: {currentUrl} - {ex.Message}")
        Finally
            If localCts IsNot Nothing Then
                localCts.Dispose()
            End If
        End Try

        Return results
    End Function

    ''' <summary>
    ''' Resolves a possibly relative link against the provided base URL and returns the absolute URL string.
    ''' </summary>
    ''' <param name="baseUrl">Page URL that acts as the anchor for relative links.</param>
    ''' <param name="relativeUrl">Relative or absolute href value extracted from a link.</param>
    ''' <returns>Absolute URL when resolvable; otherwise an empty string.</returns>
    Private Function GetAbsoluteUrl(baseUrl As String, relativeUrl As String) As String
        Try
            Dim baseUri As New Uri(baseUrl)
            Dim absoluteUri As New Uri(baseUri, relativeUrl)
            Return absoluteUri.ToString()
        Catch ex As Exception
            ' Invalid relative URL handling
            Return String.Empty
        End Try
    End Function

End Class