' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Slides.vb
' Purpose: Provides PowerPoint presentation manipulation capabilities using
'          OpenXML SDK for slide extraction, creation, and modification.
'
' Architecture:
'  - JSON Extraction: Reads .pptx files and serializes slide metadata, layouts,
'    placeholders, and text content to JSON for AI processing.
'  - Plan Application: Parses AI-generated JSON action plans and executes
'    slide additions with titles, text, bullets, shapes, and SVG icons.
'  - Layout Resolution: Resolves layouts by relId, URI, or name; falls back
'    to cover-like or default layouts when not found.
'  - Template Cloning: Clones layout structure to new slides, copies images,
'    and purges sample text from placeholders.
'  - Text Population: Supports plain text, bulleted lists, and freestanding
'    text boxes with style properties (font, size, bold, italic, color).
'  - Shape Creation: Adds geometric shapes with fill, outline, and optional
'    text content using absolute or percentage-based positioning.
'  - Speaker Notes: Creates or updates notes slide parts with provided content.
'  - Validation: Includes PPTX package validation and OpenXML error checking.
'  - External Dependencies: DocumentFormat.OpenXml, System.Text.Json,
'    SharedLibrary.SharedMethods for UI messaging.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Data
Imports System.Diagnostics
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Validation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


#Region "JSON Extraction"

    ''' <summary>
    ''' Extracts presentation metadata and content from a PPTX file as JSON.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file.</param>
    ''' <returns>JSON string containing slides, layouts, and content; empty string on error.</returns>
    Public Function GetPresentationJson(pptxPath As String) As String
        If Not System.IO.File.Exists(pptxPath) Then
            ShowCustomMessageBox($"File not found: {pptxPath}")
            Return String.Empty
        End If

        If Not IsValidPptxPackage(pptxPath) Then
            ShowCustomMessageBox("PowerPoint file is corrupt or unreadable.")
            Return String.Empty
        End If

        Try
            Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
            DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)

                Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                If presPart Is Nothing OrElse presPart.Presentation Is Nothing Then
                    ShowCustomMessageBox("Invalid or corrupted presentation.")
                    Return String.Empty
                End If

                Dim result As New PresentationJson With {
                    .Title = presDoc.PackageProperties.Title,
                    .Slides = New List(Of SlideJson)(),
                    .Layouts = New List(Of LayoutJson)()
                }

                ' Extract slide dimensions
                If presPart.Presentation.SlideSize IsNot Nothing AndAlso
                   presPart.Presentation.SlideSize.Cx IsNot Nothing AndAlso
                   presPart.Presentation.SlideSize.Cy IsNot Nothing Then
                    result.SlideSize = New SlideSizeJson With {
                        .Width = presPart.Presentation.SlideSize.Cx.Value,
                        .Height = presPart.Presentation.SlideSize.Cy.Value
                    }
                End If

                ' Check for existing slides
                Dim slideIdList = presPart.Presentation.SlideIdList
                Dim hasSlides As Boolean =
                    (slideIdList IsNot Nothing AndAlso slideIdList.ChildElements _
                        .OfType(Of DocumentFormat.OpenXml.Presentation.SlideId)().Any())

                If Not hasSlides Then
                    ' Gather layouts from masters even if there are no slides
                    Try
                        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
                            If sm Is Nothing Then Continue For
                            Dim masterName As System.String = GetMasterName(sm)
                            For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                                If layoutPart Is Nothing OrElse layoutPart.Uri Is Nothing Then Continue For
                                Dim name As System.String = GetLayoutName(layoutPart)
                                Dim layoutUri As System.String = layoutPart.Uri.ToString()
                                Dim relId As System.String = System.String.Empty
                                Try
                                    relId = sm.GetIdOfPart(layoutPart)
                                Catch
                                End Try

                                result.Layouts.Add(New LayoutJson With {
                                    .Name = name,
                                    .LayoutId = layoutUri,
                                    .LayoutRelId = relId
                                })
                            Next
                        Next
                    Catch
                    End Try

                    Return System.Text.Json.JsonSerializer.Serialize(
                        result,
                        New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}
                    )
                End If

                ' Enumerate slides safely
                Try
                    Dim idx As Integer = 0
                    For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In
                        slideIdList.ChildElements.OfType(Of DocumentFormat.OpenXml.Presentation.SlideId)()

                        If sid.RelationshipId Is Nothing Then Continue For
                        Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = Nothing
                        Try
                            sp = TryCast(presPart.GetPartById(sid.RelationshipId),
                                     DocumentFormat.OpenXml.Packaging.SlidePart)
                        Catch
                            Continue For
                        End Try
                        If sp Is Nothing Then Continue For

                        Dim title As String = GetSlideTitle(sp)
                        Dim key As String = If(
                            String.IsNullOrWhiteSpace(title),
                            $"SID-{sid.Id.Value}",
                            $"{SanitizeKey(title)}-{sid.Id.Value}"
                        )

                        Dim layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp.SlideLayoutPart
                        Dim layoutName As String = GetLayoutName(layoutPart)
                        Dim masterName As String = If(
                            layoutPart IsNot Nothing,
                            GetMasterName(layoutPart.SlideMasterPart),
                            String.Empty
                        )

                        Dim placeholders As New List(Of String)
                        Dim content As New List(Of String)

                        If sp.Slide IsNot Nothing AndAlso
                           sp.Slide.CommonSlideData IsNot Nothing AndAlso
                           sp.Slide.CommonSlideData.ShapeTree IsNot Nothing Then

                            CollectPlaceholdersFromShapeTree(sp.Slide.CommonSlideData.ShapeTree, placeholders)
                            CollectTextsFromShapeTree(sp.Slide.CommonSlideData.ShapeTree, content)
                        End If

                        result.Slides.Add(New SlideJson With {
                            .SlideKey = key,
                            .SlideId = sid.Id.Value,
                            .Index = idx,
                            .Title = title,
                            .Layout = layoutName,
                            .Master = masterName,
                            .Placeholders = placeholders,
                            .Content = content
                        })
                        idx += 1
                    Next
                Catch
                    ' Return partial results on error
                    Return System.Text.Json.JsonSerializer.Serialize(
                        result,
                        New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}
                    )
                End Try

                ' Enumerate layouts
                Try
                    For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
                        If sm Is Nothing Then Continue For
                        For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                            If layoutPart Is Nothing OrElse layoutPart.Uri Is Nothing Then Continue For
                            Dim name As String = GetLayoutName(layoutPart)
                            Dim layoutUri As String = layoutPart.Uri.ToString()
                            Dim relId As String = String.Empty
                            Try
                                relId = sm.GetIdOfPart(layoutPart)
                            Catch
                            End Try
                            result.Layouts.Add(New LayoutJson With {
                                .Name = name,
                                .LayoutId = layoutUri,
                                .LayoutRelId = relId
                            })
                        Next
                    Next
                Catch
                End Try

                Return System.Text.Json.JsonSerializer.Serialize(
                    result,
                    New System.Text.Json.JsonSerializerOptions With {.WriteIndented = True}
                )
            End Using

        Catch ex As System.IO.IOException
            ShowCustomMessageBox($"Error opening presentation (I/O): {ex.Message}")
            Return String.Empty
        Catch ex As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ShowCustomMessageBox($"Error processing presentation (OpenXML): {ex.Message}")
            Return String.Empty
        Catch ex As System.Exception
            ShowCustomMessageBox($"Unexpected error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Collects placeholder type names from a shape tree.
    ''' </summary>
    ''' <param name="tree">The shape tree to scan.</param>
    ''' <param name="placeholders">List to populate with placeholder type names.</param>
    Private Shared Sub CollectPlaceholdersFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal placeholders As System.Collections.Generic.List(Of System.String))

        If tree Is Nothing Then Return

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements

            If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(child, DocumentFormat.OpenXml.Presentation.Shape)

                Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    shp.NonVisualShapeProperties

                If nv IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type IsNot Nothing Then

                    placeholders.Add(
                        nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type.Value.ToString())
                End If

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                ' Recursively process children within groups
                Dim grp As DocumentFormat.OpenXml.Presentation.GroupShape =
                    CType(child, DocumentFormat.OpenXml.Presentation.GroupShape)

                If grp.ChildElements.Count > 0 Then
                    CollectPlaceholdersFromGroup(grp, placeholders)
                End If
            End If
            ' Note: GraphicFrame (tables/charts) do not have placeholders
        Next
    End Sub

    ''' <summary>
    ''' Recursively collects placeholder types from a group shape.
    ''' </summary>
    ''' <param name="group">The group shape to scan.</param>
    ''' <param name="placeholders">List to populate with placeholder type names.</param>
    Private Shared Sub CollectPlaceholdersFromGroup(
        ByVal group As DocumentFormat.OpenXml.Presentation.GroupShape,
        ByVal placeholders As System.Collections.Generic.List(Of System.String))

        For Each inner As DocumentFormat.OpenXml.OpenXmlElement In group.ChildElements
            If TypeOf inner Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(inner, DocumentFormat.OpenXml.Presentation.Shape)

                Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    shp.NonVisualShapeProperties

                If nv IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type IsNot Nothing Then

                    placeholders.Add(
                        nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type.Value.ToString())
                End If

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectPlaceholdersFromGroup(
                    CType(inner, DocumentFormat.OpenXml.Presentation.GroupShape), placeholders)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Collects text content from shapes, groups, and tables in a shape tree.
    ''' </summary>
    ''' <param name="tree">The shape tree to scan.</param>
    ''' <param name="content">List to populate with extracted text.</param>
    Private Shared Sub CollectTextsFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal content As System.Collections.Generic.List(Of System.String))

        If tree Is Nothing Then Return

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements

            If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(child, DocumentFormat.OpenXml.Presentation.Shape)

                If shp.TextBody IsNot Nothing Then
                    Dim txt As System.String = ExtractTextFromTextContainer(shp.TextBody)
                    If Not System.String.IsNullOrWhiteSpace(txt) Then content.Add(txt)
                End If

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectTextsFromGroup(
                    CType(child, DocumentFormat.OpenXml.Presentation.GroupShape), content)

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                    CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)

                Dim g As DocumentFormat.OpenXml.Drawing.Graphic = gf.Graphic
                If g IsNot Nothing AndAlso g.GraphicData IsNot Nothing Then
                    Dim gd As DocumentFormat.OpenXml.Drawing.GraphicData = g.GraphicData

                    ' Extract text from tables
                    Dim tbl As DocumentFormat.OpenXml.Drawing.Table =
                        gd.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Table)()
                    If tbl IsNot Nothing Then
                        ExtractTextFromTable(tbl, content)
                    End If
                    ' Note: Charts/SmartArt could be handled here additionally
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extracts concatenated text from all paragraphs in a text container element.
    ''' </summary>
    ''' <param name="container">The container element (TextBody or similar).</param>
    ''' <returns>Concatenated text content trimmed.</returns>
    Private Shared Function ExtractTextFromTextContainer(
        ByVal container As DocumentFormat.OpenXml.OpenXmlElement) As System.String

        If container Is Nothing Then Return System.String.Empty

        Dim parts As New System.Collections.Generic.List(Of System.String)()

        ' Walk all Drawing.Paragraph descendants regardless of the exact TextBody type
        For Each p As DocumentFormat.OpenXml.Drawing.Paragraph In
            container.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()

            Dim runs As New System.Collections.Generic.List(Of System.String)()

            For Each r As DocumentFormat.OpenXml.Drawing.Run In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Run)()
                If r IsNot Nothing AndAlso r.Text IsNot Nothing Then
                    runs.Add(r.Text.Text)
                End If
            Next

            For Each br As DocumentFormat.OpenXml.Drawing.Break In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Break)()
                runs.Add(vbLf)
            Next

            For Each fld As DocumentFormat.OpenXml.Drawing.Field In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Field)()
                If fld IsNot Nothing AndAlso fld.Text IsNot Nothing Then
                    runs.Add(fld.Text.Text)
                End If
            Next

            parts.Add(System.String.Join(System.String.Empty, runs))
        Next

        Return System.String.Join(vbCrLf, parts).Trim()
    End Function

    ''' <summary>
    ''' Recursively collects text content from a group shape.
    ''' </summary>
    ''' <param name="group">The group shape to scan.</param>
    ''' <param name="content">List to populate with extracted text.</param>
    Private Shared Sub CollectTextsFromGroup(
        ByVal group As DocumentFormat.OpenXml.Presentation.GroupShape,
        ByVal content As System.Collections.Generic.List(Of System.String))

        ' Cannot use IsNot comparison on OpenXmlElementList - only check Count
        If group.ChildElements.Count = 0 Then Return

        For Each inner As DocumentFormat.OpenXml.OpenXmlElement In group.ChildElements

            If TypeOf inner Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(inner, DocumentFormat.OpenXml.Presentation.Shape)

                If shp.TextBody IsNot Nothing Then
                    Dim txt As System.String = ExtractTextFromTextContainer(shp.TextBody)
                    If Not System.String.IsNullOrWhiteSpace(txt) Then content.Add(txt)
                End If

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectTextsFromGroup(
                    CType(inner, DocumentFormat.OpenXml.Presentation.GroupShape), content)

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                    CType(inner, DocumentFormat.OpenXml.Presentation.GraphicFrame)

                Dim g As DocumentFormat.OpenXml.Drawing.Graphic = gf.Graphic
                If g IsNot Nothing AndAlso g.GraphicData IsNot Nothing Then
                    Dim gd As DocumentFormat.OpenXml.Drawing.GraphicData = g.GraphicData

                    Dim tbl As DocumentFormat.OpenXml.Drawing.Table =
                        gd.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Table)()
                    If tbl IsNot Nothing Then
                        ExtractTextFromTable(tbl, content)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extracts text from a Drawing.Table element, adding tab-separated rows.
    ''' </summary>
    ''' <param name="table">The table element to extract text from.</param>
    ''' <param name="content">List to populate with row text.</param>
    Private Shared Sub ExtractTextFromTable(
        ByVal table As DocumentFormat.OpenXml.Drawing.Table,
        ByVal content As System.Collections.Generic.List(Of System.String))

        If table Is Nothing Then Return

        For Each row As DocumentFormat.OpenXml.Drawing.TableRow In
            table.Elements(Of DocumentFormat.OpenXml.Drawing.TableRow)()

            Dim rowTexts As New System.Collections.Generic.List(Of System.String)()

            For Each cell As DocumentFormat.OpenXml.Drawing.TableCell In
                row.Elements(Of DocumentFormat.OpenXml.Drawing.TableCell)()

                If cell IsNot Nothing AndAlso cell.TextBody IsNot Nothing Then
                    Dim cellText As System.String = ExtractTextFromTextContainer(cell.TextBody)
                    rowTexts.Add(cellText)
                End If
            Next

            Dim line As System.String = System.String.Join(vbTab, rowTexts)
            If Not System.String.IsNullOrWhiteSpace(line) Then content.Add(line)
        Next
    End Sub


#End Region

    ''' <summary>
    ''' Validates a PPTX file by checking ZIP archive integrity and entry accessibility.
    ''' </summary>
    ''' <param name="path">Full path to the PowerPoint file.</param>
    ''' <returns>True if the file is a valid, readable PPTX package.</returns>
    Private Function IsValidPptxPackage(path As String) As Boolean
        Try
            Using archive As New System.IO.Compression.ZipArchive(System.IO.File.OpenRead(path), IO.Compression.ZipArchiveMode.Read)
                For Each entry In archive.Entries
                    ' Optional: basic sanity check — size not huge, name not empty
                    If String.IsNullOrWhiteSpace(entry.FullName) Then Return False
                    ' Read small text files fully to catch unreadable/corrupt XML
                    If entry.Length > 0 AndAlso entry.Length < 5_000_000 Then
                        Using s = entry.Open()
                            ' Read a few bytes to ensure stream is accessible
                            Dim buffer(255) As Byte
                            Dim read = s.Read(buffer, 0, buffer.Length)
                        End Using
                    End If
                Next
            End Using
            Return True
        Catch
            ' If ZIP open fails or reading any entry fails, it's not valid
            Return False
        End Try
    End Function

#Region "JSON DTOs"

    ''' <summary>
    ''' Represents a single slide's metadata and content for JSON serialization.
    ''' </summary>
    Public Class SlideJson
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String

        <JsonPropertyName("slideId")>
        Public Property SlideId As UInteger

        <JsonPropertyName("index")>
        Public Property Index As Integer

        <JsonPropertyName("title")>
        Public Property Title As String

        <JsonPropertyName("layout")>
        Public Property Layout As String

        <JsonPropertyName("master")>
        Public Property Master As String

        <JsonPropertyName("placeholders")>
        Public Property Placeholders As List(Of String)

        <JsonPropertyName("content")>
        Public Property Content As List(Of String)
    End Class


    ''' <summary>
    ''' Represents a slide layout's metadata for JSON serialization.
    ''' </summary>
    Public Class LayoutJson
        <JsonPropertyName("name")>
        Public Property Name As String

        <JsonPropertyName("layoutId")>
        Public Property LayoutId As String

        <JsonPropertyName("layoutRelId")>
        Public Property LayoutRelId As String
    End Class

    ''' <summary>
    ''' Represents slide dimensions in EMUs for JSON serialization.
    ''' </summary>
    Public Class SlideSizeJson
        <JsonPropertyName("width")>
        Public Property Width As Long

        <JsonPropertyName("height")>
        Public Property Height As Long
    End Class

    ''' <summary>
    ''' Root DTO for presentation JSON export including slides and layouts.
    ''' </summary>
    Public Class PresentationJson
        <JsonPropertyName("title")>
        Public Property Title As String

        <JsonPropertyName("slideSize")>
        Public Property SlideSize As SlideSizeJson

        <JsonPropertyName("slides")>
        Public Property Slides As List(Of SlideJson)

        <JsonPropertyName("layouts")>
        Public Property Layouts As List(Of LayoutJson)
    End Class

    ''' <summary>
    ''' Internal class to track placeholder types present in a layout.
    ''' </summary>
    Private NotInheritable Class LayoutInfo
        Public Property HasTitle As System.Boolean
        Public Property HasCenteredTitle As System.Boolean
        Public Property HasSubTitle As System.Boolean
        Public Property HasBody As System.Boolean
    End Class

    ''' <summary>
    ''' Base class for action plan operations.
    ''' </summary>
    Public MustInherit Class ActionBase
        <JsonPropertyName("op")>
        Public Property Op As String
    End Class

    ''' <summary>
    ''' Represents anchor positioning for slide insertion.
    ''' </summary>
    Public Class Anchor
        <JsonPropertyName("mode")>
        Public Property Mode As String
        <JsonPropertyName("by")>
        Public Property By As AnchorBy
    End Class

    ''' <summary>
    ''' Specifies the slide key for anchor positioning.
    ''' </summary>
    Public Class AnchorBy
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String
    End Class

    ''' <summary>
    ''' Represents an add_slide action with anchor, layout, and elements.
    ''' </summary>
    Public Class AddSlideAction
        Inherits ActionBase
        <JsonPropertyName("anchor")> Public Property Anchor As Anchor
        <JsonPropertyName("layoutRelId")> Public Property LayoutRelId As String
        <JsonPropertyName("elements")> Public Property Elements As List(Of JsonElement)
    End Class


    ''' <summary>
    ''' Index mapping slide keys to IDs and IDs to positions.
    ''' </summary>
    Public Class DeckIndex
        Public Property SlideKeyById As Dictionary(Of String, UInteger)
        Public Property IndexBySlideId As Dictionary(Of UInteger, Integer)
    End Class

#End Region

    ''' <summary>
    ''' Analyzes a layout part to determine which placeholder types are present.
    ''' </summary>
    ''' <param name="lp">The slide layout part to analyze.</param>
    ''' <returns>LayoutInfo with flags for title, subtitle, and body placeholders.</returns>
    Private Function AnalyzeLayoutPlaceholders(lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As LayoutInfo
        Dim li As New LayoutInfo()
        Dim tree = lp?.SlideLayout?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return li

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph = shp?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            If ph.Type Is Nothing Then
                ' in PPT the "subtitle" box is frequently a placeholder with no explicit type but index=1
                If ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then
                    li.HasSubTitle = True
                End If
                Continue For
            End If
            Select Case ph.Type.Value
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title
                    li.HasTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle
                    li.HasCenteredTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                    li.HasSubTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body
                    li.HasBody = True
            End Select
        Next

        Return li
    End Function

    ''' <summary>
    ''' Extracts the title text from a slide's title placeholder.
    ''' </summary>
    ''' <param name="sp">The slide part to search.</param>
    ''' <returns>Title text or empty string if not found.</returns>
    Private Function GetSlideTitle(sp As SlidePart) As String
        If sp Is Nothing OrElse sp.Slide Is Nothing OrElse
       sp.Slide.CommonSlideData Is Nothing OrElse
       sp.Slide.CommonSlideData.ShapeTree Is Nothing Then
            Return String.Empty
        End If

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In
        sp.Slide.CommonSlideData.ShapeTree.ChildElements _
          .OfType(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim nv = shp.NonVisualShapeProperties
            If nv IsNot Nothing AndAlso nv.ApplicationNonVisualDrawingProperties IsNot Nothing Then
                Dim ph = nv.ApplicationNonVisualDrawingProperties.PlaceholderShape
                If ph IsNot Nothing AndAlso
               (ph.Type Is Nothing OrElse
                ph.Type.Value = PlaceholderValues.Title OrElse
                ph.Type.Value = PlaceholderValues.CenteredTitle) Then
                    Return If(shp.TextBody IsNot Nothing, shp.TextBody.InnerText, String.Empty)
                End If
            End If
        Next
        Return String.Empty
    End Function


    ''' <summary>
    ''' Gets the human-readable name of a slide layout.
    ''' </summary>
    ''' <param name="layoutPart">The layout part.</param>
    ''' <returns>Layout name or URI string as fallback.</returns>
    Private Function GetLayoutName(layoutPart As SlideLayoutPart) As String
        If layoutPart Is Nothing Then Return String.Empty
        If layoutPart.SlideLayout IsNot Nothing AndAlso
       layoutPart.SlideLayout.CommonSlideData IsNot Nothing Then
            Dim nm = layoutPart.SlideLayout.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then Return nm
        End If
        Return If(layoutPart.Uri IsNot Nothing, layoutPart.Uri.ToString(), String.Empty)
    End Function

    ''' <summary>
    ''' Gets the human-readable name of a slide master.
    ''' </summary>
    ''' <param name="smPart">The slide master part.</param>
    ''' <returns>Master name or URI string as fallback.</returns>
    Private Function GetMasterName(smPart As SlideMasterPart) As String
        If smPart Is Nothing Then Return String.Empty
        If smPart.SlideMaster IsNot Nothing AndAlso
       smPart.SlideMaster.CommonSlideData IsNot Nothing Then
            Dim nm = smPart.SlideMaster.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then Return nm
        End If
        Return If(smPart.Uri IsNot Nothing, smPart.Uri.ToString(), String.Empty)
    End Function

    ''' <summary>
    ''' Sanitizes a string for use as a slide key by replacing non-alphanumeric characters.
    ''' </summary>
    ''' <param name="s">The string to sanitize.</param>
    ''' <returns>Sanitized string with only letters, digits, and hyphens.</returns>
    Private Function SanitizeKey(s As String) As String
        Return New String(
            s.Select(Function(ch) If(Char.IsLetterOrDigit(ch), ch, "-"c)).ToArray()
        )
    End Function

    ''' <summary>
    ''' Cleans a raw string to extract valid JSON by finding matching braces/brackets.
    ''' </summary>
    ''' <param name="raw">The raw string potentially containing JSON.</param>
    ''' <returns>Cleaned JSON string or trimmed original.</returns>
    Public Function CleanJsonString(raw As String) As String
        If String.IsNullOrEmpty(raw) Then
            Return String.Empty
        End If

        ' Look for object vs. array start
        Dim firstObj = raw.IndexOf("{"c)
        Dim firstArr = raw.IndexOf("["c)
        Dim startIdx As Integer
        Dim openChar As Char
        Dim closeChar As Char

        If firstObj >= 0 AndAlso (firstObj < firstArr OrElse firstArr = -1) Then
            startIdx = firstObj
            openChar = "{"c
            closeChar = "}"c
        ElseIf firstArr >= 0 Then
            startIdx = firstArr
            openChar = "["c
            closeChar = "]"c
        Else
            ' No JSON delimiters found – just return trimmed
            Return raw.Trim()
        End If

        ' Find the last matching closing brace/bracket
        Dim lastIdx = raw.LastIndexOf(closeChar)
        If lastIdx > startIdx Then
            Return raw.Substring(startIdx, lastIdx - startIdx + 1).Trim()
        Else
            ' Malformed or unmatched – fallback to trimming
            Return raw.Trim()
        End If
    End Function


    ''' <summary>
    ''' Applies an AI-generated action plan to modify a PowerPoint presentation.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file to modify.</param>
    ''' <param name="planJson">JSON string containing the action plan with slide operations.</param>
    ''' <returns>True if all actions applied successfully; False on error.</returns>
    Public Function ApplyPlanToPresentation(pptxPath As String, planJson As String) As Boolean
        Try
            ' 1) Check if the file exists
            If Not System.IO.File.Exists(pptxPath) Then
                ShowCustomMessageBox($"Your file '{pptxPath}' was no longer found - aborting.")
                Return False
            End If

            ' 2) Configure JSON serializer options
            Dim opts As New System.Text.Json.JsonSerializerOptions With {
            .PropertyNameCaseInsensitive = True
        }
            opts.Converters.Add(New System.Text.Json.Serialization.JsonStringEnumConverter())

            Dim actions As System.Text.Json.JsonElement.ArrayEnumerator
            Try
                actions = System.Text.Json.JsonDocument.Parse(planJson) _
                      .RootElement _
                      .GetProperty("actions") _
                      .EnumerateArray()
            Catch ex As System.Text.Json.JsonException
                ShowCustomMessageBox("The AI has sent an invalid instruction on how to build the slides: " & ex.Message)
                Return False
            Catch ex As KeyNotFoundException
                ShowCustomMessageBox("An internal error occurred when amending your slidedeck (the AI sent instructions missing the required 'actions' array).")
                Return False
            End Try

            Dim errorMessages As New List(Of String)

            ' 3) Open the presentation
            Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
              DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, True)

                Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                If presPart Is Nothing Then
                    ShowCustomMessageBox("A presentation is missing in the file you have provided; you may have to include at least one slide.")
                    Return False
                End If

                ' Ensure SlideIdList exists
                EnsureSlideIdList(presPart)

                ' 4) Build deck index
                Dim idx As DeckIndex = BuildDeckIndex(presPart)
                Dim currentAnchorKey As String = Nothing

                ' 5) Process actions
                For Each actElem In actions
                    If Not actElem.TryGetProperty("op", Nothing) _
                   OrElse actElem.GetProperty("op").GetString() <> "add_slide" Then
                        Continue For
                    End If

                    Try
                        ' --- 5.1 Anchor (robust) ---
                        Dim anchorId As UInteger = 0UI
                        Dim anchorEl As System.Text.Json.JsonElement
                        Dim byEl As System.Text.Json.JsonElement
                        Dim anchorKey As System.String = Nothing

                        If actElem.TryGetProperty("anchor", anchorEl) Then
                            ' mode is optional; for "at_end" we simply leave anchorId = 0UI
                            If anchorEl.TryGetProperty("by", byEl) AndAlso byEl.TryGetProperty("slideKey", Nothing) Then
                                anchorKey = byEl.GetProperty("slideKey").GetString()
                            End If
                        End If

                        ' Resolve anchorId only if we actually got a key
                        If Not System.String.IsNullOrEmpty(anchorKey) Then
                            If anchorKey <> "lastInserted" Then
                                currentAnchorKey = anchorKey
                            End If

                            Dim keyToUse As System.String = If(anchorKey = "lastInserted", currentAnchorKey, anchorKey)
                            If Not System.String.IsNullOrEmpty(keyToUse) AndAlso idx.SlideKeyById.ContainsKey(keyToUse) Then
                                anchorId = idx.SlideKeyById(keyToUse)
                            End If
                        End If

                        ' --- 5.1b Layout key (guarded) ---
                        Dim layoutRelId As System.String = Nothing
                        Dim tmpEl As System.Text.Json.JsonElement
                        If actElem.TryGetProperty("layoutRelId", tmpEl) Then
                            layoutRelId = tmpEl.GetString()
                        ElseIf actElem.TryGetProperty("layoutKey", tmpEl) AndAlso tmpEl.ValueKind = JsonValueKind.Object Then
                            Dim ridEl As JsonElement
                            If tmpEl.TryGetProperty("relId", ridEl) Then
                                layoutRelId = ridEl.GetString()
                            End If
                        End If

                        ' 5.2 Clone slide
                        Dim newSp As DocumentFormat.OpenXml.Packaging.SlidePart =
                        CloneTemplateSlide(presPart, layoutRelId)
                        Dim newId As UInteger = InsertAfter(presPart, anchorId, newSp)

                        ' 5.3 Populate elements
                        For Each el In actElem.GetProperty("elements").EnumerateArray()
                            Dim t As String = el.GetProperty("type").GetString()
                            Select Case t
                                Case "title"
                                    Dim txt = el.GetProperty("text").GetString()
                                    SetTitle(newSp, txt, el)
                                Case "shape"
                                    AddShape(presPart, newSp, el)
                                Case "svg_icon"
                                    AddSvgIcon(presPart, newSp, el)
                                Case "text"
                                    If el.TryGetProperty("transform", Nothing) Then
                                        CreateFreestandingTextBox(presPart, newSp, el)
                                    Else
                                        ' Don't read placeholder as string anymore - SetText will handle it
                                        Dim txt = el.GetProperty("text").GetString()
                                        SetTextWithPlaceholder(newSp, el.GetProperty("placeholder"), txt, el)
                                    End If
                                Case "bullet_text"
                                    If el.TryGetProperty("transform", Nothing) Then
                                        CreateFreestandingTextBox(presPart, newSp, el)
                                    Else
                                        SetBulletsWithPlaceholder(newSp, el)
                                    End If
                            End Select
                        Next

                        ' 5.4 Speaker notes
                        Dim notesEl As System.Text.Json.JsonElement
                        If actElem.TryGetProperty("notes", notesEl) AndAlso notesEl.ValueKind = JsonValueKind.String Then
                            SetSpeakerNotes(newSp, notesEl.GetString())
                        End If

                        RemoveEmptyBodyPlaceholder(newSp)

                        ' Save intermediate
                        presPart.Presentation.Save()

                        ' 5.5 Rebuild index
                        idx = BuildDeckIndex(presPart)
                        currentAnchorKey = GetSlideKey(newSp, newId)

                    Catch ex As KeyNotFoundException
                        Debug.WriteLine("Could not implement instruction: " & ex.Message)
                        errorMessages.Add("Could not implement instruction: " & ex.Message)

                    Catch ex As Exception
                        Debug.WriteLine("Error creating slides: " & ex.Message)
                        errorMessages.Add("Error creating slides: " & ex.Message)
                    End Try
                Next

                ' 6) Fallback: ensure at least empty notes for every slide
                For Each sid In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                    Dim spPart As DocumentFormat.OpenXml.Packaging.SlidePart =
                        CType(presPart.GetPartById(sid.RelationshipId),
                              DocumentFormat.OpenXml.Packaging.SlidePart)

                    If spPart.NotesSlidePart Is Nothing Then
                        ' Create empty notes
                        SetSpeakerNotes(spPart, String.Empty)
                    End If
                Next

                ' 7) Final save
                presPart.Presentation.Save()
            End Using

            If errorMessages IsNot Nothing AndAlso errorMessages.Count > 0 Then
                Dim allErrors As String = String.Join(vbCrLf, errorMessages)
                ShowCustomMessageBox("Several errors occurred during applying the AI's instruction to your slidedeck (it may still have worked partially):" & vbCrLf & vbCrLf & allErrors)
                Return False
            End If

            Return True

        Catch oxEx As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ShowCustomMessageBox("A PowerPoint file error occurred: " & oxEx.Message)
            Return False
        Catch ex As Exception
            ShowCustomMessageBox("An unexpected error occurred when amending your slidedeck: " & ex.Message)
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Validates a PPTX file for OpenXML compliance errors.
    ''' </summary>
    ''' <param name="path">Full path to the PowerPoint file to validate.</param>
    ''' <returns>Error description string if validation fails; empty string if no errors found.</returns>
    ''' <remarks>Returns only the first error found for debugging purposes.</remarks>
    Function ValidatePptx(path As String) As String

        Dim ErrorString As String = ""

        Using doc As PresentationDocument = PresentationDocument.Open(path, False)

            Dim validator As New OpenXmlValidator()
            Dim errors = validator.Validate(doc)

            If Not errors.Any() Then
                Debug.WriteLine("No formal OpenXML-Errors found.")
                Return ""
            End If

            For Each err As ValidationErrorInfo In errors
                Debug.WriteLine("----------")
                Debug.WriteLine($"Part : {err.Part.Uri}")
                Debug.WriteLine($"XPath: {err.Path.XPath}")
                Debug.WriteLine($"Info : {err.Description}")
                ErrorString = $"Part: {err.Part.Uri}; XPath: {err.Path.XPath}; Info: {err.Description}"
                ' Stop after first error – sufficient for debugging
                Exit For
            Next

        End Using

        Return ErrorString

    End Function


    ' Overload: Takes PresentationPart instead of path

    ''' <summary>
    ''' Builds a deck index from a presentation part for slide lookups.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>DeckIndex with slide key and ID mappings.</returns>    
    Public Function BuildDeckIndex(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DeckIndex
        Dim idx As New DeckIndex With {
        .SlideKeyById = New Dictionary(Of String, UInteger)(),
        .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
    }
        Dim i As Integer = 0
        For Each sid In presPart.Presentation.SlideIdList.Elements(
                            Of DocumentFormat.OpenXml.Presentation.SlideId)()
            idx.IndexBySlideId(sid.Id.Value) = i
            Dim sp = CType(presPart.GetPartById(sid.RelationshipId),
                       DocumentFormat.OpenXml.Packaging.SlidePart)
            Dim key = GetSlideKey(sp, sid.Id.Value)
            idx.SlideKeyById(key) = sid.Id.Value
            i += 1
        Next
        Return idx
    End Function

    ''' <summary>
    ''' Builds a deck index from a PPTX file path.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file.</param>
    ''' <returns>DeckIndex with slide key and ID mappings.</returns>
    Public Function BuildDeckIndex(pptxPath As String) As DeckIndex
        Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
              DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)
            Dim presPart = presDoc.PresentationPart
            Dim idx As New DeckIndex With {
              .SlideKeyById = New Dictionary(Of String, UInteger)(),
              .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
            }
            Dim i As Integer = 0
            For Each sid As DocumentFormat.OpenXml.Presentation.SlideId _
                In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                idx.IndexBySlideId(sid.Id.Value) = i
                Dim sp = CType(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                Dim key = GetSlideKey(sp, sid.Id.Value)
                idx.SlideKeyById(key) = sid.Id.Value
                i += 1
            Next
            Return idx
        End Using
    End Function


    ''' <summary>
    ''' Clones a template slide from the specified layout.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="layoutRelId">Layout relationship ID, URI, or name.</param>
    ''' <returns>A new slide part cloned from the layout.</returns>
    Private Function CloneTemplateSlide(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    layoutRelId As System.String
) As DocumentFormat.OpenXml.Packaging.SlidePart

        Dim targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart =
        ResolveLayout(presPart, layoutRelId)

        If targetLayout Is Nothing Then
            ' final fallback to a decent default (prefer Title+Subtitle without Body if possible)
            targetLayout = PickCoverLikeLayout(presPart)
            If targetLayout Is Nothing Then
                targetLayout = PickDefaultLayout(presPart)
            End If
        End If

        Dim newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart =
        presPart.AddNewPart(Of DocumentFormat.OpenXml.Packaging.SlidePart)()

        Dim newSlide As New DocumentFormat.OpenXml.Presentation.Slide()

        If targetLayout.SlideLayout.CommonSlideData IsNot Nothing Then
            newSlide.CommonSlideData = CType(
            targetLayout.SlideLayout.CommonSlideData.CloneNode(True),
            DocumentFormat.OpenXml.Presentation.CommonSlideData)
        End If

        If targetLayout.SlideLayout.ColorMapOverride IsNot Nothing Then
            newSlide.ColorMapOverride = CType(
            targetLayout.SlideLayout.ColorMapOverride.CloneNode(True),
            DocumentFormat.OpenXml.Presentation.ColorMapOverride)
        End If

        PurgeLayoutSampleText(newSlide)

        newSlidePart.Slide = newSlide

        CopyLayoutImagesToSlide(targetLayout, newSlidePart)

        newSlidePart.AddPart(targetLayout)

        newSlidePart.Slide.Save()
        Return newSlidePart
    End Function

    ''' <summary>
    ''' Resolves a layout by relationship ID, URI, or name.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="requested">The layout identifier to resolve.</param>
    ''' <returns>The matching SlideLayoutPart or Nothing if not found.</returns>
    Private Function ResolveLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    requested As System.String
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        If System.String.IsNullOrWhiteSpace(requested) Then Return Nothing

        Dim req = requested.Trim()

        ' 1) Try as exact relId
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim rid As System.String = System.String.Empty
                Try : rid = sm.GetIdOfPart(lp) : Catch : End Try
                If Not System.String.IsNullOrEmpty(rid) AndAlso
               System.String.Equals(rid, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        ' 2) Try by URI string
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim u = lp.Uri?.ToString()
                If Not System.String.IsNullOrEmpty(u) AndAlso
               System.String.Equals(u, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        ' 3) Try by human-readable layout name (e.g., "Title Slide" / "Titel")
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim name As System.String = GetLayoutName(lp)
                If Not System.String.IsNullOrEmpty(name) AndAlso
               System.String.Equals(name, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Picks a cover-like layout (Title + Subtitle, preferably without Body).
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>A suitable SlideLayoutPart or Nothing.</returns>
    Private Function PickCoverLikeLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim li = AnalyzeLayoutPlaceholders(lp)
                ' Typical title slide: Title + Subtitle; often NO Body placeholder
                If (li.HasTitle OrElse li.HasCenteredTitle) AndAlso li.HasSubTitle AndAlso Not li.HasBody Then
                    Return lp
                End If
            Next
        Next

        ' next best: Title + Subtitle, even if Body exists
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim li = AnalyzeLayoutPlaceholders(lp)
                If (li.HasTitle OrElse li.HasCenteredTitle) AndAlso li.HasSubTitle Then
                    Return lp
                End If
            Next
        Next

        Return Nothing
    End Function


    ''' <summary>
    ''' Picks a default layout with Title and Body placeholders.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>A suitable SlideLayoutPart.</returns>
    ''' <exception cref="System.Exception">Thrown when no layout is available.</exception>
    Private Function PickDefaultLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        Dim firstMaster As DocumentFormat.OpenXml.Packaging.SlideMasterPart =
        presPart.SlideMasterParts.FirstOrDefault()
        If firstMaster Is Nothing Then
            Throw New System.Exception("No SlideMasterPart found in the presentation.")
        End If

        ' Prefer a layout that has both Title and Body placeholders
        For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In firstMaster.SlideLayoutParts
            Dim hasTitle As Boolean = False
            Dim hasBody As Boolean = False

            If lp.SlideLayout IsNot Nothing AndAlso
           lp.SlideLayout.CommonSlideData IsNot Nothing AndAlso
           lp.SlideLayout.CommonSlideData.ShapeTree IsNot Nothing Then

                Dim shapes =
                lp.SlideLayout.CommonSlideData.ShapeTree.
                    Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

                For Each sh As DocumentFormat.OpenXml.Presentation.Shape In shapes
                    Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    sh.NonVisualShapeProperties
                    If nv Is Nothing Then Continue For

                    Dim app As DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties =
                    nv.ApplicationNonVisualDrawingProperties
                    If app Is Nothing Then Continue For

                    Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                    app.PlaceholderShape
                    If ph Is Nothing OrElse ph.Type Is Nothing Then Continue For

                    Dim t As DocumentFormat.OpenXml.Presentation.PlaceholderValues = ph.Type.Value
                    If t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse
                   t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle Then
                        hasTitle = True
                    ElseIf t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then
                        hasBody = True
                    End If

                    If hasTitle AndAlso hasBody Then
                        Return lp
                    End If
                Next
            End If
        Next

        ' Fallback: first available layout
        Dim anyLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart =
        firstMaster.SlideLayoutParts.FirstOrDefault()
        If anyLayout Is Nothing Then
            Throw New System.Exception("No SlideLayoutPart available to create a new slide.")
        End If
        Return anyLayout
    End Function



    ''' <summary>
    ''' Copies image parts from a layout to a slide and rewrites embed IDs.
    ''' </summary>
    ''' <param name="layoutPart">The source layout part.</param>
    ''' <param name="slidePart">The target slide part.</param>
    Private Sub CopyLayoutImagesToSlide(
        layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
        slidePart As DocumentFormat.OpenXml.Packaging.SlidePart)

        ' 1) oldRelId → newRelId
        Dim idMap As New System.Collections.Generic.Dictionary(Of String, String)(
        System.StringComparer.OrdinalIgnoreCase)

        ' 2) clone ONLY ImageParts — everything else stays in the layout
        For Each img In layoutPart.ImageParts
            Dim oldId As String = layoutPart.GetIdOfPart(img)

            ' create a fresh ImagePart in the slide
            Dim newImg = slidePart.AddImagePart(img.ContentType)
            ' copy the binary
            Using src = img.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read),
              dst = newImg.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                src.CopyTo(dst)
            End Using

            idMap(oldId) = slidePart.GetIdOfPart(newImg)
        Next

        ' 3) rewrite every <a:blip embed="…">
        For Each blip In slidePart.Slide.
             Descendants(Of DocumentFormat.OpenXml.Drawing.Blip)()
            Dim oldId = blip.Embed?.Value
            If oldId IsNot Nothing AndAlso idMap.ContainsKey(oldId) Then
                blip.Embed.Value = idMap(oldId)
            End If
        Next
    End Sub


    ''' <summary>
    ''' Clears sample text from title and body placeholders in a cloned slide.
    ''' </summary>
    ''' <param name="sld">The slide to purge.</param>
    Private Sub PurgeLayoutSampleText(sld As DocumentFormat.OpenXml.Presentation.Slide)

        ' only Title / CenteredTitle / Body placeholders get wiped
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape _
        In sld.CommonSlideData.ShapeTree.
               Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                 ApplicationNonVisualDrawingProperties?.
                 PlaceholderShape
            If ph Is Nothing Then Continue For

            Dim t As DocumentFormat.OpenXml.Presentation.PlaceholderValues? = Nothing
            If ph.Type IsNot Nothing Then t = ph.Type.Value

            If t Is Nothing _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object Then

                ' wipe existing content
                shp.TextBody?.Remove()

                ' insert minimal, valid skeleton
                shp.Append(New DocumentFormat.OpenXml.Presentation.TextBody(
                    New DocumentFormat.OpenXml.Drawing.BodyProperties(),
                    New DocumentFormat.OpenXml.Drawing.ListStyle(),
                    New DocumentFormat.OpenXml.Drawing.Paragraph(
                        New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties())))
            End If
        Next
    End Sub


    ''' <summary>
    ''' Ensures the presentation has a SlideIdList element.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    Private Sub EnsureSlideIdList(presPart As DocumentFormat.OpenXml.Packaging.PresentationPart)
        Dim pres = presPart.Presentation
        If pres.SlideIdList IsNot Nothing Then Exit Sub

        Dim sldIdList As New DocumentFormat.OpenXml.Presentation.SlideIdList()

        ' Find the correct insertion index:
        ' Order (simplified): SlideMasterIdList?, NotesMasterIdList?, HandoutMasterIdList?, SlideIdList?, SlideSize?, NotesSize? ...
        Dim children = pres.ChildElements
        Dim insertIndex As Integer = children.Count ' default to end, then adjust

        ' Prefer to insert BEFORE SlideSize or NotesSize if present
        For i As Integer = 0 To children.Count - 1
            If TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.SlideSize _
        OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.NotesSize Then
                insertIndex = i
                Exit For
            End If
        Next

        ' If we didn't find sizes, place right after SlideMasterIdList / NotesMasterIdList / HandoutMasterIdList if any
        If insertIndex = children.Count Then
            Dim afterIndex As Integer = -1
            For i As Integer = 0 To children.Count - 1
                If TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.SlideMasterIdList _
            OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.NotesMasterIdList _
            OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.HandoutMasterIdList Then
                    afterIndex = i
                End If
            Next
            insertIndex = If(afterIndex >= 0, afterIndex + 1, 0)
        End If

        pres.InsertAt(sldIdList, insertIndex)
        pres.Save()
    End Sub


    ''' <summary>
    ''' Inserts a new slide after the specified anchor slide or at the end.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="anchorSlideId">The slide ID to insert after; 0 to append at end.</param>
    ''' <param name="newSlidePart">The new slide part to insert.</param>
    ''' <returns>The new slide ID assigned.</returns>
    Private Function InsertAfter(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    anchorSlideId As UInteger,
    newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart
) As UInteger

        Dim slideList = presPart.Presentation.SlideIdList
        Dim relId = presPart.GetIdOfPart(newSlidePart)

        ' Find existing SlideId elements
        Dim existing = slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
        Dim newId As UInteger

        If existing.Any() Then
            newId = existing.Max(Function(s) s.Id.Value) + 1UI
        Else
            newId = 256UI   ' First slide
        End If

        Dim newSlide = New DocumentFormat.OpenXml.Presentation.SlideId() With {
      .Id = newId,
      .RelationshipId = relId
    }

        ' If anchorSlideId = 0, always append at the end
        If anchorSlideId = 0UI Then
            slideList.Append(newSlide)
        Else
            ' Otherwise insert after the specified anchor
            Dim anchor = existing.FirstOrDefault(Function(s) s.Id.Value = anchorSlideId)
            If anchor Is Nothing Then
                slideList.Append(newSlide)
            Else
                anchor.InsertAfterSelf(newSlide)
            End If
        End If

        Return newId
    End Function


    ''' <summary>
    ''' Sets the title text in a slide's title placeholder shape.
    ''' </summary>
    ''' <param name="sp">The slide part containing the title shape.</param>
    ''' <param name="text">The title text to insert.</param>
    ''' <param name="el">JSON element containing style properties.</param>
    ''' <remarks>
    ''' Search priority: 1) Explicit Title/CenteredTitle placeholder, 2) Placeholder with index=0, 
    ''' 3) Shape name containing "title" (case-insensitive).
    ''' </remarks>
    Private Sub SetTitle(
        sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        text As System.String,
        el As System.Text.Json.JsonElement)

        Dim shapes = sp.Slide.CommonSlideData.ShapeTree.
                 Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        ' --- 1) explicit Title / CenteredTitle --------------------------------
        titleShape = shapes.FirstOrDefault(Function(shp)
                                               Dim ph = shp.NonVisualShapeProperties? _
                 .ApplicationNonVisualDrawingProperties? _
                 .PlaceholderShape
                                               Return ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
               (ph.Type.Value =
                DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse
                ph.Type.Value =
                DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle)
                                           End Function)

        ' --- 2) implicit: placeholder with ph:index = 0 -----------------------
        If titleShape Is Nothing Then
            titleShape = shapes.FirstOrDefault(Function(shp)
                                                   Dim ph = shp.NonVisualShapeProperties? _
                     .ApplicationNonVisualDrawingProperties? _
                     .PlaceholderShape
                                                   Return ph IsNot Nothing AndAlso ph.Index IsNot Nothing AndAlso
                   ph.Index.Value = 0UI
                                               End Function)
        End If

        ' --- 3) last fallback: shape name contains "title" --------------------
        If titleShape Is Nothing Then
            titleShape = shapes.FirstOrDefault(Function(shp)
                                                   Dim nm = shp.NonVisualShapeProperties? _
                     .NonVisualDrawingProperties?.Name?.Value
                                                   Return Not System.String.IsNullOrWhiteSpace(nm) AndAlso
                   nm.IndexOf("title",
                              System.StringComparison.OrdinalIgnoreCase) >= 0
                                               End Function)
        End If

        If titleShape Is Nothing Then Return   ' nothing suitable found

        ' --- 4) build a fresh TextBody ----------------------------------------
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle())

        tb.Append(BuildParagraph(text, el))   ' your existing helper

        titleShape.TextBody = tb
        sp.Slide.Save()
    End Sub



    ''' <summary>
    ''' Sets bulleted text content in a body placeholder shape.
    ''' </summary>
    ''' <param name="sp">The slide part containing the shape.</param>
    ''' <param name="el">JSON element containing bullets array, optional placeholder name, and style properties.</param>
    ''' <remarks>
    ''' Searches for target shape in order: Body/Object placeholders, placeholder by name, first non-title shape.
    ''' Each bullet can specify text and optional nesting level (0-8).
    ''' </remarks>
    Private Sub SetBullets(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement
)
        ' 1) Read optional placeholder name from JSON
        Dim placeholderName As String = Nothing
        Dim tmpEl As System.Text.Json.JsonElement
        If el.TryGetProperty("placeholder", tmpEl) Then
            placeholderName = tmpEl.GetString()
        End If

        ' 2) Search all shapes on the slide
        Dim allShapes = sp.Slide.CommonSlideData.ShapeTree.
                Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
        Dim bodyShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        ' 2a) Prefer Body or Object placeholder boxes (and exclude footer/date/slide number)
        For Each shp In allShapes
            Dim ph = shp.NonVisualShapeProperties?.
         ApplicationNonVisualDrawingProperties?.
         PlaceholderShape
            If IsBodyLikePlaceholder(ph) Then
                bodyShape = shp
                Exit For
            End If
        Next

        ' 2b) Fallback: JSON placeholder in shape name
        If bodyShape Is Nothing AndAlso Not String.IsNullOrEmpty(placeholderName) Then
            For Each shp In allShapes
                Dim nvProps = shp.NonVisualShapeProperties? _
                      .NonVisualDrawingProperties
                Dim shpName As String = If(nvProps?.Name?.Value, "")
                If shpName.IndexOf(placeholderName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    bodyShape = shp
                    Exit For
                End If
            Next
        End If

        ' 2c) Fallback: first non-title shape (and also not footer/date/slide number)
        If bodyShape Is Nothing Then
            For Each shp In allShapes
                Dim ph = shp.NonVisualShapeProperties?.
             ApplicationNonVisualDrawingProperties?.
             PlaceholderShape
                Dim typ = If(ph?.Type IsNot Nothing, ph.Type.Value, Nothing)

                Dim isTitle = (typ = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse
                   typ = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle)
                Dim isFooterLike = (typ = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer OrElse
                        typ = DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime OrElse
                        typ = DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber)

                If Not isTitle AndAlso Not isFooterLike Then
                    bodyShape = shp
                    Exit For
                End If
            Next
        End If

        ' Abort if no body shape was found
        If bodyShape Is Nothing Then
            Return
        End If

        ' 3) Create new TextBody with ListStyle
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
        tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        tb.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())

        ' 4) Read bullets from JSON and append as nested paragraphs
        For Each bElem As System.Text.Json.JsonElement In el.GetProperty("bullets").EnumerateArray()
            ' 4a) Extract text and level
            Dim text As String
            Dim level As Integer = 0
            If bElem.ValueKind = System.Text.Json.JsonValueKind.Object Then
                If bElem.TryGetProperty("text", tmpEl) Then
                    text = tmpEl.GetString()
                Else
                    Continue For
                End If
                If bElem.TryGetProperty("level", tmpEl) Then
                    level = tmpEl.GetInt32()
                End If
            Else
                text = bElem.GetString()
            End If

            ' 4b) Create RunProperties from el.style
            Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
            Dim styleEl As System.Text.Json.JsonElement
            If el.TryGetProperty("style", styleEl) Then
                Dim tmp As System.Text.Json.JsonElement
                If styleEl.TryGetProperty("fontFamily", tmp) Then
                    rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
                End If
                If styleEl.TryGetProperty("fontSize", tmp) Then
                    rp.FontSize = CUInt(tmp.GetInt32() * 100)
                End If
                If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
                If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            End If

            ' 4c) Set ParagraphProperties with level
            Dim actualLevel = System.Math.Max(0, System.Math.Min(8, level))
            Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                .Level = CByte(actualLevel)
            }
            ' 4d) Build Run and Paragraph
            Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
            Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
            para.Append(pPr)
            para.Append(runElem)

            tb.Append(para)
        Next

        ' 5) Assign TextBody to shape and save
        bodyShape.TextBody = tb
        sp.Slide.Save()
    End Sub


    ''' <summary>
    ''' Builds a single Drawing.Paragraph with optional paragraph properties.
    ''' </summary>
    ''' <param name="text">The text content.</param>
    ''' <param name="el">JSON element containing style properties.</param>
    ''' <param name="pPr">Optional paragraph properties to apply.</param>
    ''' <returns>A fully configured Paragraph element.</returns>
    Private Function BuildParagraph(
      text As String,
      el As JsonElement
    ) As DocumentFormat.OpenXml.Drawing.Paragraph

        ' 1) RunProperties and Style from JSON
        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            End If
            If styleEl.TryGetProperty("fontSize", tmp) Then
                rp.FontSize = CUInt(tmp.GetInt32() * 100)
            End If
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            If styleEl.TryGetProperty("color", tmp) Then
                Dim hex = tmp.GetString().TrimStart("#"c)

            End If
        End If

        ' 2) Create Run + Paragraph 
        Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
        Dim para = New DocumentFormat.OpenXml.Drawing.Paragraph()
        para.Append(runElem)
        Return para
    End Function


    Private Function BuildParagraph(text As String, el As JsonElement, Optional pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = Nothing) As DocumentFormat.OpenXml.Drawing.Paragraph
        Dim rp = New DocumentFormat.OpenXml.Drawing.RunProperties()
        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            Dim tmp As JsonElement
            If styleEl.TryGetProperty("fontFamily", tmp) Then rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
            If styleEl.TryGetProperty("fontSize", tmp) Then rp.FontSize = CInt(tmp.GetInt32() * 100)
            If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
            If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
            If styleEl.TryGetProperty("color", tmp) Then
                rp.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = tmp.GetString().TrimStart("#"c)}))
            End If
        End If

        Dim run = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
        Dim para = New DocumentFormat.OpenXml.Drawing.Paragraph()
        If pPr IsNot Nothing Then
            para.Append(pPr.CloneNode(True))
        End If
        para.Append(run)
        Return para
    End Function


    ''' <summary>
    ''' Generates a unique slide key combining sanitized title and slide ID.
    ''' </summary>
    ''' <param name="sp">The slide part to extract title from.</param>
    ''' <param name="slideId">The slide's unique ID.</param>
    ''' <returns>Slide key in format "{SanitizedTitle}-{slideId}" or "SID-{slideId}" if no title.</returns>
    Private Function GetSlideKey(
        sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        slideId As UInteger
      ) As String
        Dim title = GetSlideTitle(sp)
        If String.IsNullOrWhiteSpace(title) Then
            Return $"SID-{slideId}"
        Else
            Return $"{SanitizeKey(title)}-{slideId}"
        End If
    End Function

    ''' <summary>
    ''' Creates or updates speaker notes for a slide with the specified text content.
    ''' </summary>
    ''' <param name="sp">The slide part to add or update notes for.</param>
    ''' <param name="notesText">The notes text content.</param>
    ''' <remarks>
    ''' Creates a new NotesSlidePart if none exists. Removes existing shapes and inserts
    ''' a new body placeholder at index 2 (after header elements).
    ''' </remarks>
    Private Sub SetSpeakerNotes(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    notesText As String)

        Dim notesPart As DocumentFormat.OpenXml.Packaging.NotesSlidePart = sp.NotesSlidePart
        If notesPart Is Nothing Then
            notesPart = sp.AddNewPart(Of DocumentFormat.OpenXml.Packaging.NotesSlidePart)()
            notesPart.NotesSlide = New DocumentFormat.OpenXml.Presentation.NotesSlide(
            New DocumentFormat.OpenXml.Presentation.CommonSlideData(
                New DocumentFormat.OpenXml.Presentation.ShapeTree(
                    New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1UI, .Name = ""},
                        New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                    New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())),
            New DocumentFormat.OpenXml.Presentation.ColorMapOverride(
                New DocumentFormat.OpenXml.Drawing.MasterColorMapping()))
        End If

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree =
        notesPart.NotesSlide.CommonSlideData.ShapeTree

        ' ----- only remove Shapes/Pics -----
        For Each n In tree.ChildElements.OfType(Of DocumentFormat.OpenXml.OpenXmlElement)().ToList()
            If TypeOf n Is DocumentFormat.OpenXml.Presentation.Shape _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.Picture _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                n.Remove()
            End If
        Next

        ' ----- new Body-Shape -----
        Dim nvSpPr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 2UI, .Name = "NotesBody"},
        New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
            New DocumentFormat.OpenXml.Drawing.ShapeLocks() With {.NoGrouping = True}),
        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties(
            New DocumentFormat.OpenXml.Presentation.PlaceholderShape() With {
                .Type = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                .Index = 1UI}))
        Dim shapePr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        Dim noteShape As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, shapePr)

        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle())
        Dim run As New DocumentFormat.OpenXml.Drawing.Run(
        New DocumentFormat.OpenXml.Drawing.RunProperties(),
        New DocumentFormat.OpenXml.Drawing.Text(notesText))
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph(run) With {
        .ParagraphProperties = New DocumentFormat.OpenXml.Drawing.ParagraphProperties()}
        tb.Append(para)
        noteShape.Append(tb)

        ' insert after header
        If tree.ChildElements.Count >= 2 Then
            tree.InsertAt(noteShape, 2)
        Else
            tree.Append(noteShape)
        End If

        notesPart.NotesSlide.Save()
    End Sub



    ''' <summary>
    ''' Creates an OpenXML Fill element from a JSON definition.
    ''' </summary>
    ''' <param name="fillJson">JSON element containing fill type and color properties.</param>
    ''' <returns>SolidFill element with specified color, or NoFill as fallback.</returns>

    Private Function CreateFill(fillJson As JsonElement) As DocumentFormat.OpenXml.OpenXmlElement
        Dim fillType As String = ""
        If fillJson.TryGetProperty("type", Nothing) Then fillType = fillJson.GetProperty("type").GetString()

        Select Case fillType.ToLower()
            Case "solid"
                If fillJson.TryGetProperty("color", Nothing) Then
                    Dim colorHex = fillJson.GetProperty("color").GetString().TrimStart("#"c)
                    Return New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorHex})
                End If
        End Select
        Return New DocumentFormat.OpenXml.Drawing.NoFill() ' Fallback
    End Function

    ''' <summary>
    ''' Creates an OpenXML Outline element from a JSON definition.
    ''' </summary>
    ''' <param name="outlineJson">JSON element containing width, color, and dashType properties.</param>
    ''' <returns>Configured Outline element with width in EMUs, color, and dash style.</returns>
    ''' <remarks>Width is parsed culture-invariantly; 1 point = 12700 EMUs.</remarks>
    Private Function CreateOutline(outlineJson As JsonElement) As DocumentFormat.OpenXml.Drawing.Outline
        Dim outline As New DocumentFormat.OpenXml.Drawing.Outline()

        Dim widthJson As JsonElement
        If outlineJson.TryGetProperty("width", widthJson) Then
            ' [FIX] This safely parses numbers like "1" or "1.5" regardless of system language.
            Dim widthValue As Double
            If Double.TryParse(widthJson.GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, widthValue) Then
                outline.Width = CInt(widthValue * 12700) ' 1 point = 12700 EMUs
            End If
        End If

        Dim colorJson As JsonElement
        If outlineJson.TryGetProperty("color", colorJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorJson.GetString().TrimStart("#"c)}))
        End If

        Dim dashJson As JsonElement
        If outlineJson.TryGetProperty("dashType", dashJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.PresetDash With {.Val = JsonDashNameToEnumValue(dashJson.GetString())})
        End If

        Return outline
    End Function



    ''' <summary>
    ''' Converts relative percentage-based coordinates from JSON into absolute EMU coordinates.
    ''' </summary>
    ''' <param name="presPart">The presentation part, to get the master slide dimensions.</param>
    ''' <param name="transformJson">The JSON "transform" object.</param>
    ''' <returns>A fully calculated Transform2D object with absolute EMUs.</returns>
    Private Function ConvertRelativeToAbsoluteTransform(presPart As DocumentFormat.OpenXml.Packaging.PresentationPart, transformJson As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Drawing.Transform2D
        ' Get the master slide dimensions in EMUs
        Dim slideWidthEmu = presPart.Presentation.SlideSize.Cx.Value
        Dim slideHeightEmu = presPart.Presentation.SlideSize.Cy.Value

        ' Safely parse the relative percentage values from JSON
        Dim relX, relY, relW, relH As Double
        Double.TryParse(transformJson.GetProperty("x").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relX)
        Double.TryParse(transformJson.GetProperty("y").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relY)
        Double.TryParse(transformJson.GetProperty("width").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relW)
        Double.TryParse(transformJson.GetProperty("height").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relH)

        ' Calculate the absolute EMU values
        Dim absX = CLng(slideWidthEmu * relX)
        Dim absY = CLng(slideHeightEmu * relY)
        Dim absCx = CLng(slideWidthEmu * relW)
        Dim absCy = CLng(slideHeightEmu * relH)

        Return New DocumentFormat.OpenXml.Drawing.Transform2D(
        New DocumentFormat.OpenXml.Drawing.Offset With {.X = absX, .Y = absY},
        New DocumentFormat.OpenXml.Drawing.Extents With {.Cx = absCx, .Cy = absCy}
    )
    End Function


    ''' <summary>
    ''' Converts a JSON shape name string to the corresponding OpenXML ShapeTypeValues enum.
    ''' </summary>
    ''' <param name="jsonName">Case-insensitive shape name (e.g., "rectangle", "circle").</param>
    ''' <returns>Matching ShapeTypeValues enum; defaults to Rectangle if unknown.</returns>
    Private Function JsonShapeNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.ShapeTypeValues
        Select Case jsonName.ToLower()
            Case "rectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
            Case "oval", "ellipse", "circle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse
            Case "line" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line
            Case "rightarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RightArrow
            Case "leftarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.LeftArrow
            Case "triangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Triangle ' Corrected from IsoscelesTriangle
            Case "roundedrectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle
            Case "flowchartprocess" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartProcess
            Case "flowchartdecision" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartDecision
            Case "flowchartterminator" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartTerminator
            Case "chevron" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Chevron
            Case "pentagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Pentagon
            Case "hexagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Hexagon
            Case "plus" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Plus
            Case "blockarc" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.BlockArc
            Case Else : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle ' Fallback
        End Select
    End Function


    ''' <summary>
    ''' Converts a JSON dash style name to the corresponding OpenXML PresetLineDashValues enum.
    ''' </summary>
    ''' <param name="jsonName">Case-insensitive dash name (e.g., "solid", "dashed", "dotted").</param>
    ''' <returns>Matching PresetLineDashValues enum; defaults to Solid if unknown.</returns>
    Private Function JsonDashNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.PresetLineDashValues
        Select Case jsonName.ToLower()
            Case "solid"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid
            Case "dot", "dotted"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot
            Case "dash", "dashed"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash
            Case "longdash"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDash
            Case "dashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.DashDot
            Case "longdashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDashDot
            Case Else
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid ' Fallback
        End Select
    End Function


    ''' <summary>
    ''' Builds a Drawing.Paragraph with text, nesting level, and optional bullet formatting.
    ''' </summary>
    ''' <param name="text">The text content.</param>
    ''' <param name="level">Nesting level for bullet indentation.</param>
    ''' <param name="el">JSON element containing style properties (font, size, bold, color, align).</param>
    ''' <param name="isBulleted">True to use default bullets; False to apply NoBullet.</param>
    ''' <returns>Configured Paragraph element with style and bullet properties.</returns>
    Private Function BuildStyledParagraph(text As String, level As Integer, el As System.Text.Json.JsonElement, isBulleted As Boolean) As DocumentFormat.OpenXml.Drawing.Paragraph
        Dim pPr = New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {.Level = level}
        Dim rp = New DocumentFormat.OpenXml.Drawing.RunProperties()

        Dim styleEl As JsonElement
        If el.TryGetProperty("style", styleEl) Then
            If styleEl.TryGetProperty("font", Nothing) Then rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = styleEl.GetProperty("font").GetString()})
            If styleEl.TryGetProperty("size", Nothing) Then rp.FontSize = CInt(styleEl.GetProperty("size").GetInt32() * 100)
            If styleEl.TryGetProperty("bold", Nothing) AndAlso styleEl.GetProperty("bold").GetBoolean() Then rp.Bold = True
            'If styleEl.TryGetProperty("color", Nothing) Then rp.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = styleEl.GetProperty("color").GetString().TrimStart("#"c)}))
            If styleEl.TryGetProperty("align", Nothing) Then
                Select Case styleEl.GetProperty("align").GetString().ToLower()
                    Case "center" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center
                    Case "right" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right
                    Case Else : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left
                End Select
            End If
        End If

        If Not isBulleted Then
            pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        End If

        Return New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text)))
    End Function

    ''' <summary>
    ''' Creates a freestanding text box shape with text or bulleted content at specified coordinates.
    ''' </summary>
    ''' <param name="presPart">The presentation part for slide dimensions.</param>
    ''' <param name="sp">The slide part to add the text box to.</param>
    ''' <param name="el">JSON element containing transform, text/bullets, and style properties.</param>
    ''' <remarks>
    ''' Supports both "text" and "bullet_text" types. Transform can be in EMUs or percentages.
    ''' Bullets use 0.5cm hanging indent with character bullet (•).
    ''' </remarks>
    Private Sub CreateFreestandingTextBox(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) Find next available shape ID
        Dim maxId As UInteger = 0
        For Each nonVisPr As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties _
        In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nonVisPr.Id.Value > maxId Then maxId = nonVisPr.Id.Value
        Next
        Dim newId As UInteger = maxId + 1

        ' 2) Locate the transform JSON
        Dim tf As System.Text.Json.JsonElement
        If Not el.TryGetProperty("transform", tf) Then
            If el.TryGetProperty("style", tf) AndAlso tf.TryGetProperty("transform", tf) Then
                ' nested under style
            Else
                Return
            End If
        End If

        ' 3) Compute absolute EMU coordinates
        Dim rawX As Double
        Double.TryParse(tf.GetProperty("x").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, rawX)

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D
        If rawX > 1 Then
            ' already EMU
            Dim ofs As New DocumentFormat.OpenXml.Drawing.Offset() With {
            .X = CLng(tf.GetProperty("x").GetInt64()),
            .Y = CLng(tf.GetProperty("y").GetInt64())
        }
            Dim ext As New DocumentFormat.OpenXml.Drawing.Extents() With {
            .Cx = CLng(tf.GetProperty("width").GetInt64()),
            .Cy = CLng(tf.GetProperty("height").GetInt64())
        }
            xfrm = New DocumentFormat.OpenXml.Drawing.Transform2D(ofs, ext)
        Else
            ' percent → EMU
            xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf)
        End If

        ' 4) Build the textbox shape
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = xfrm}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle})
        spPr.Append(New DocumentFormat.OpenXml.Drawing.NoFill())

        Dim nvDr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties() With {.TextBox = True}
        nvDr.AppendChild(New DocumentFormat.OpenXml.Drawing.ShapeLocks())

        Dim nvProps As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {
            .Id = newId,
            .Name = "TextBox " & newId
        },
        nvDr,
        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
    )

        Dim shp As New DocumentFormat.OpenXml.Presentation.Shape(nvProps, spPr)

        ' 5) Populate text or bullets
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle()
    )

        Select Case el.GetProperty("type").GetString()
            Case "text"
                tb.Append(BuildParagraph(el.GetProperty("text").GetString(), el))
            Case "bullet_text"
                For Each b In el.GetProperty("bullets").EnumerateArray()
                    Dim txt As String = If(
                        b.ValueKind = JsonValueKind.Object,
                        b.GetProperty("text").GetString(),
                        b.GetString())

                    Dim lvl As Integer = 0
                    Dim tmp As System.Text.Json.JsonElement
                    If b.ValueKind = JsonValueKind.Object AndAlso b.TryGetProperty("level", tmp) Then
                        lvl = tmp.GetInt32()
                    End If

                    ' Classic 0.5-cm hanging indent: bullet at 0, text at 0.5 cm
                    Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                        .Level = CByte(System.Math.Max(0, System.Math.Min(8, lvl))),
                        .LeftMargin = 457200,     ' 0.5 cm
                        .Indent = -457200,    ' hanging indent equals left margin
                        .Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left}

                    pPr.Append(New DocumentFormat.OpenXml.Drawing.BulletFont() With {.Typeface = "Arial"})
                    pPr.Append(New DocumentFormat.OpenXml.Drawing.CharacterBullet() With {.Char = "•"c})

                    ' Run: *no* extra tab needed – indent handles spacing
                    Dim run = New DocumentFormat.OpenXml.Drawing.Run(
                  New DocumentFormat.OpenXml.Drawing.RunProperties(),
                  New DocumentFormat.OpenXml.Drawing.Text(txt))

                    tb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, run))
                Next


        End Select

        shp.Append(tb)
        tree.Append(shp)
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Adds a geometric shape to a slide with optional fill, outline, and text content.
    ''' </summary>
    ''' <param name="presPart">The presentation part for slide dimensions.</param>
    ''' <param name="sp">The slide part to add the shape to.</param>
    ''' <param name="el">JSON element containing shapeType, transform, fill, outline, and optional text properties.</param>
    ''' <remarks>
    ''' Transform coordinates can be in EMUs (>1) or percentages (≤1).
    ''' Supports various shape types via JsonShapeNameToEnumValue mapping.
    ''' </remarks>
    Private Sub AddShape(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) Determine next available shape ID
        Dim maxId As UInteger = 0UI
        For Each nvPr As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties _
        In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nvPr.Id.Value > maxId Then maxId = nvPr.Id.Value
        Next
        Dim newId As UInteger = maxId + 1UI

        ' 2) Extract transform from JSON
        Dim transformJson = el.GetProperty("transform")

        ' Check raw value: ≤1 = percentage, >1 = already EMUs
        Dim rawX As Double
        If Not Double.TryParse(transformJson.GetProperty("x").GetRawText(),
                       Globalization.NumberStyles.Any,
                       Globalization.CultureInfo.InvariantCulture,
                       rawX) Then
            rawX = 0.0
        End If

        Dim absoluteTransform As DocumentFormat.OpenXml.Drawing.Transform2D

        If rawX <= 1.0 Then
            ' Percentage values → convert to EMU
            absoluteTransform = ConvertRelativeToAbsoluteTransform(presPart, transformJson)
        Else
            ' Use direct EMU values
            Dim ofs As New DocumentFormat.OpenXml.Drawing.Offset() With {
            .X = CLng(transformJson.GetProperty("x").GetInt64()),
        .Y = CLng(transformJson.GetProperty("y").GetInt64())
    }
            Dim ext As New DocumentFormat.OpenXml.Drawing.Extents() With {
        .Cx = CLng(transformJson.GetProperty("width").GetInt64()),
        .Cy = CLng(transformJson.GetProperty("height").GetInt64())
    }
            absoluteTransform = New DocumentFormat.OpenXml.Drawing.Transform2D(ofs, ext)
        End If

        ' 3) ShapeProperties
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = absoluteTransform}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = JsonShapeNameToEnumValue(el.GetProperty("shapeType").GetString())})
        If el.TryGetProperty("fill", Nothing) Then spPr.Append(CreateFill(el.GetProperty("fill")))
        If el.TryGetProperty("outline", Nothing) Then spPr.Append(CreateOutline(el.GetProperty("outline")))

        ' 4) nvSpPr (only set TextBox if text content follows)
        Dim nvSpDr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties()
        If el.TryGetProperty("text", Nothing) Then
            nvSpDr.TextBox = True
            nvSpDr.AppendChild(New DocumentFormat.OpenXml.Drawing.ShapeLocks())
        End If
        Dim nvSpPr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = newId, .Name = $"Shape {newId}"},
    nvSpDr,
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
)
        ' 5) Create new Shape
        Dim shp As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, spPr)

        ' 6) Optional text content
        If el.TryGetProperty("text", Nothing) Then
            Dim tb = New DocumentFormat.OpenXml.Presentation.TextBody(
            New DocumentFormat.OpenXml.Drawing.BodyProperties(),
            New DocumentFormat.OpenXml.Drawing.ListStyle()
        )
            tb.Append(BuildStyledParagraph(el.GetProperty("text").GetString(), 0, el, False))
            shp.Append(tb)
        End If

        ' 7) Insert and save
        tree.Append(shp)
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Inserts an SVG icon from the JSON at the given location.
    ''' Uses a standard <p:pic> with <a:blip>; this is the same recipe
    ''' PowerPoint 2019+ generates and shows on Office 2016 (Oct-2018) too.
    ''' </summary>
    Private Sub AddSvgIcon(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) unique ID on slide
        Dim newId As UInteger =
    tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)().
        Select(Function(nv) nv.Id.Value).DefaultIfEmpty(0).Max() + 1UI

        ' 2) build Transform2D (percent → EMU if needed)
        Dim tf = el.GetProperty("transform")
        Dim rawX As Double
        Double.TryParse(tf.GetProperty("x").GetRawText(),
        Globalization.NumberStyles.Any,
        Globalization.CultureInfo.InvariantCulture,
        rawX)

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D
        If rawX > 1 Then
            xfrm = New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset With {
                .X = CLng(tf.GetProperty("x").GetInt64()),
                .Y = CLng(tf.GetProperty("y").GetInt64())},
            New DocumentFormat.OpenXml.Drawing.Extents With {
                .Cx = CLng(tf.GetProperty("width").GetInt64()),
                .Cy = CLng(tf.GetProperty("height").GetInt64())})
        Else
            ' Assuming ConvertRelativeToAbsoluteTransform exists and returns a Transform2D
            ' xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf) 
            xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf)
        End If

        ' 3) embed SVG file
        Dim svgPart = sp.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Svg)
        Using ms As New IO.MemoryStream(
        System.Text.Encoding.UTF8.GetBytes(el.GetProperty("svg").GetString()))
            svgPart.FeedData(ms)
        End Using
        Dim relId As String = sp.GetIdOfPart(svgPart)

        ' 4) build <p:pic>
        Dim nvPic As New DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {
        .Id = newId, .Name = "Icon " & newId},
    New DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties(
        New DocumentFormat.OpenXml.Drawing.PictureLocks() With {.NoChangeAspect = True}),
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim blipFill As New DocumentFormat.OpenXml.Presentation.BlipFill(
    New DocumentFormat.OpenXml.Drawing.Blip() With {
        .Embed = relId,
        .CompressionState =
            DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print},
    New DocumentFormat.OpenXml.Drawing.Stretch(
        New DocumentFormat.OpenXml.Drawing.FillRectangle()))

        ' Define the rectangle geometry
        Dim prstGeom As New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle}

        ' Create the ShapeProperties object
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()

        ' Append the transform and geometry as child elements
        spPr.Append(xfrm)
        spPr.Append(prstGeom)

        ' Create the final picture by combining all the parts
        Dim pic As New DocumentFormat.OpenXml.Presentation.Picture(nvPic, blipFill, spPr)

        ' 5) append & save
        tree.Append(pic)
        sp.Slide.Save()
    End Sub


    ''' <summary>
    ''' Removes empty body placeholder shapes from a slide.
    ''' </summary>
    ''' <param name="sp">The slide part to process.</param>
    ''' <remarks>
    ''' A body placeholder is considered empty if it has no TextBody or contains only whitespace.
    ''' Only removes the first empty body placeholder found.
    ''' </remarks>
    Private Sub RemoveEmptyBodyPlaceholder(sp As DocumentFormat.OpenXml.Packaging.SlidePart)
        Dim shpToRemove As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        For Each shp In sp.Slide.CommonSlideData.ShapeTree.
                         Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                        ApplicationNonVisualDrawingProperties?.
                        PlaceholderShape
            If ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
               ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then

                ' empty = only one paragraph with no text or whitespace
                Dim empty As Boolean =
                    (shp.TextBody Is Nothing) OrElse
                    Not shp.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().
                        Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text))

                If empty Then shpToRemove = shp
                Exit For
            End If
        Next

        If shpToRemove IsNot Nothing Then
            shpToRemove.Remove()
            sp.Slide.Save()
        End If
    End Sub

    ''' <summary>
    ''' Determines whether a placeholder represents a body-like content area.
    ''' </summary>
    ''' <param name="ph">The placeholder shape to check.</param>
    ''' <returns>True if the placeholder is Body or Object type; False otherwise.</returns>
    ''' <remarks>
    ''' Excludes footer, date/time, slide number, title, and subtitle placeholders.
    ''' Implicit placeholders (no type) are treated as body-like only if index >= 2.
    ''' </remarks>
    Private Function IsBodyLikePlaceholder(ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape) As Boolean
        If ph Is Nothing Then Return False
        If ph.Type Is Nothing Then
            ' Implicit placeholder: treat as body-like only if index is typical for content (not header/footer indices)
            ' Common patterns: Title = index 0, Subtitle = index 1; Footer/Date/SlideNumber often have explicit types.
            ' With no type, be conservative: don't auto-accept implicit unless index >= 2.
            If ph.Index IsNot Nothing Then
                Return ph.Index.Value >= 2UI
            End If
            Return False
        End If

        Select Case ph.Type.Value
            Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object
                Return True
            Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                Return False
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Sets plain text in a shape identified by placeholder element.
    ''' </summary>
    ''' <param name="sp">The slide part containing the shape.</param>
    ''' <param name="placeholderEl">JSON element identifying the placeholder.</param>
    ''' <param name="text">The text to insert.</param>
    ''' <param name="el">JSON element containing style properties.</param>
    Private Sub SetTextWithPlaceholder(
    sp As SlidePart,
    placeholderEl As JsonElement,
    text As String,
    el As JsonElement)

        Dim targetShape = FindShapeByPlaceholderElement(sp, placeholderEl)
        If targetShape Is Nothing Then Return

        ' Build TextBody without bullets
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
        tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())

        ' Build paragraph with styles from el
        Dim para = BuildParagraph(text, el)
        Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties()
        pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        para.ParagraphProperties = pPr

        tb.Append(para)
        targetShape.TextBody = tb
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Sets bulleted text in a shape identified by placeholder element.
    ''' </summary>
    ''' <param name="sp">The slide part containing the shape.</param>
    ''' <param name="el">JSON element containing bullets array and style properties.</param>
    Private Sub SetBulletsWithPlaceholder(
    sp As SlidePart,
    el As JsonElement)

        Dim placeholderEl As JsonElement
        If el.TryGetProperty("placeholder", placeholderEl) Then
            Dim targetShape = FindShapeByPlaceholderElement(sp, placeholderEl)
            If targetShape Is Nothing Then Return

            ' Create TextBody with proper structure
            Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody()
            tb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
            tb.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())

            ' Process bullets - this is the missing part!
            Dim tmpEl As JsonElement
            For Each bElem In el.GetProperty("bullets").EnumerateArray()
                ' Extract text and level
                Dim text As String
                Dim level As Integer = 0

                If bElem.ValueKind = JsonValueKind.Object Then
                    If bElem.TryGetProperty("text", tmpEl) Then
                        text = tmpEl.GetString()
                    Else
                        Continue For
                    End If
                    If bElem.TryGetProperty("level", tmpEl) Then
                        level = tmpEl.GetInt32()
                    End If
                Else
                    text = bElem.GetString()
                End If

                ' Build RunProperties from style
                Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
                Dim styleEl As JsonElement
                If el.TryGetProperty("style", styleEl) Then
                    Dim tmp As JsonElement
                    If styleEl.TryGetProperty("fontFamily", tmp) Then
                        rp.Append(New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = tmp.GetString()})
                    End If
                    If styleEl.TryGetProperty("fontSize", tmp) Then
                        rp.FontSize = CUInt(tmp.GetInt32() * 100)
                    End If
                    If styleEl.TryGetProperty("bold", tmp) AndAlso tmp.GetBoolean() Then rp.Bold = True
                    If styleEl.TryGetProperty("italic", tmp) AndAlso tmp.GetBoolean() Then rp.Italic = True
                End If

                ' Create ParagraphProperties with level
                Dim actualLevel = System.Math.Max(0, System.Math.Min(8, level))
                Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                .Level = CByte(actualLevel)
            }

                ' Build Run and Paragraph
                Dim runElem = New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text))
                Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
                para.Append(pPr)
                para.Append(runElem)

                tb.Append(para)
            Next

            ' If no bullets were added, add an empty paragraph to keep valid structure
            If Not tb.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().Any() Then
                tb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(
                New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()))
            End If

            targetShape.TextBody = tb
            sp.Slide.Save()
        Else
            ' Fallback to original SetBullets
            SetBullets(sp, el)
        End If
    End Sub

    ''' <summary>
    ''' Locates a shape by placeholder specification from JSON.
    ''' </summary>
    ''' <param name="sp">The slide part to search.</param>
    ''' <param name="placeholderEl">JSON element specifying placeholder by type or name.</param>
    ''' <returns>The matching Shape or Nothing if not found.</returns>
    ''' <remarks>
    ''' Supports JSON object format like {"type": "Body"} or string format for shape name matching.
    ''' Falls back to first body-like placeholder if no match found.
    ''' </remarks>
    Private Function FindShapeByPlaceholderElement(
        sp As SlidePart,
        placeholderEl As JsonElement) As DocumentFormat.OpenXml.Presentation.Shape

        Dim allShapes = sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

        ' Handle object placeholder like { "type": "Body" }
        If placeholderEl.ValueKind = JsonValueKind.Object Then
            Dim typeEl As JsonElement
            If placeholderEl.TryGetProperty("type", typeEl) Then
                Dim typeStr = typeEl.GetString()

                For Each shp In allShapes
                    Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph Is Nothing Then Continue For

                    Select Case typeStr.ToLower()
                        Case "body"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = PlaceholderValues.Body Then Return shp
                        Case "object"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = PlaceholderValues.Object Then Return shp
                        Case "subtitle"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = PlaceholderValues.SubTitle Then Return shp
                            If ph.Type Is Nothing AndAlso ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then Return shp
                        Case "title"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = PlaceholderValues.Title Then Return shp
                        Case "centeredtitle"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = PlaceholderValues.CenteredTitle Then Return shp
                    End Select
                Next
            End If
        ElseIf placeholderEl.ValueKind = JsonValueKind.String Then
            ' Handle string placeholder (shape name match)
            Dim nameToFind = placeholderEl.GetString()
            For Each shp In allShapes
                Dim nm = shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value
                If Not String.IsNullOrEmpty(nm) AndAlso nm.IndexOf(nameToFind, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return shp
                End If
            Next
        End If

        ' Fallback to body-like placeholder
        For Each shp In allShapes
            Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If IsBodyLikePlaceholder(ph) Then Return shp
        Next

        Return Nothing
    End Function


End Class
