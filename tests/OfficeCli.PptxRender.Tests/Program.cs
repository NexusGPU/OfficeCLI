// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli;
using OfficeCli.Handlers;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

var path = Path.Combine(Path.GetTempPath(), $"officecli-pptx-render-{Guid.NewGuid():N}.pptx");
try
{
    CreateFixture(path);
    var explicitNoAutofitPath = FindShapePathByText(path, "explicit no autofit");
    var compactSingleLinePath = FindShapePathByText(path, "compact single line");
    var trailingEmptyParagraphsPath = FindShapePathByText(path, "trailing empty paragraphs");
    using (var handler = new PowerPointHandler(path, editable: false))
    {
        var html = handler.ViewAsHtml(startSlide: 1, endSlide: 1);

        Assert(html.Contains("rgba(0,0,0,0.85)", StringComparison.Ordinal),
            "preset color with alpha should render as rgba");
        Assert(html.Contains("rgba(0,0,0,0.5)", StringComparison.Ordinal),
            "system color with alpha should render as rgba");
        Assert(html.Contains("font-size:9pt", StringComparison.Ordinal),
            "non-placeholder shape should inherit font size from its local list style");
        Assert(html.Contains("color:#FFFFFF", StringComparison.Ordinal),
            "non-placeholder shape should inherit color from its local list style");
        Assert(html.Contains("font-family:'OPPOSans M'", StringComparison.Ordinal),
            "non-placeholder shape should inherit typeface from its local list style");
        Assert(html.Contains("font-size:17pt", StringComparison.Ordinal),
            "placeholder should inherit font size from the master body text style");
        Assert(html.Contains("color:#123456", StringComparison.Ordinal),
            "placeholder should inherit color from the matching master placeholder");
        Assert(html.Contains("font-family:'Master Placeholder'", StringComparison.Ordinal),
            "placeholder should inherit typeface from the matching master placeholder");
        Assert(html.Contains("text-align:center", StringComparison.Ordinal),
            "placeholder should inherit alignment from the matching layout placeholder");
        Assert(html.Contains("padding-left:20pt", StringComparison.Ordinal),
            "placeholder should continue past a partial layout style to the master placeholder margin");
        Assert(html.Contains("padding-right:10pt", StringComparison.Ordinal),
            "placeholder should continue to the presentation default text style for missing properties");
        Assert(html.Contains("margin-bottom:6pt", StringComparison.Ordinal),
            "placeholder should merge paragraph spacing from the master body text style");
        var inheritanceText = html.IndexOf("master inheritance cascade", StringComparison.Ordinal);
        Assert(inheritanceText >= 0, "master inheritance fixture text should render");
        var inheritanceContext = html[Math.Max(0, inheritanceText - 1200)..inheritanceText];
        Assert(inheritanceContext.Contains("font-weight:bold", StringComparison.Ordinal),
            "placeholder should inherit bold from the matching layout placeholder");
        Assert(inheritanceContext.Contains("font-style:italic", StringComparison.Ordinal),
            "placeholder should continue to the master placeholder for missing italic style");
        Assert(html.Contains("1/1/2024", StringComparison.Ordinal),
            "date-axis serial categories should use the axis number format");
        var renderedDateLabels = CountOccurrences(html, "/2024");
        Assert(renderedDateLabels > 0 && renderedDateLabels < 12,
            "date axes should adaptively thin labels when no explicit interval is present");
        Assert(!html.Contains("rotate(-1000", StringComparison.Ordinal),
            "out-of-range WPS axis rotation should not be emitted as an SVG rotation");
        Assert(html.Contains(
                "clip-path:polygon(0 0,93.75% 0,100% 50%,93.75% 100%,0 100%,6.25% 50%)",
                StringComparison.Ordinal),
            "wide chevrons should derive notch depth from the short side");
        var opaqueGradientStop = html.IndexOf("#FFFFFF 30%", StringComparison.Ordinal);
        var transparentGradientStop = html.IndexOf("rgba(255,255,255,0) 77%", StringComparison.Ordinal);
        Assert(opaqueGradientStop >= 0 && transparentGradientStop > opaqueGradientStop,
            "gradient stops should be ordered by position before emitting CSS");
        Assert(html.Contains(
                "<path d=\"M0 100000 L50000 0 L100000 50000\" fill=\"none\" stroke=\"",
                StringComparison.Ordinal),
            "shape-level no-fill custom geometry should render its stroke as SVG");
        Assert(CountOccurrences(html, "stroke=\"#806040\"") == 1,
            "dashed custom geometry should not receive a duplicate rectangle outline");
        Assert(html.Contains(
                "markerUnits=\"userSpaceOnUse\" markerWidth=\"6\" markerHeight=\"6\"",
                StringComparison.Ordinal),
            "thin connector arrowheads should retain a visible minimum size");
        Assert(!html.Contains("-webkit-text-stroke", StringComparison.Ordinal),
            "a no-fill run outline should not paint a default black glyph stroke");
        Assert(handler.CheckShapeTextOverflow(explicitNoAutofitPath) == null,
            "explicit no-autofit should be treated as intentional overflow");
        Assert(handler.CheckShapeTextOverflow(compactSingleLinePath) == null,
            "single-line text should not require inter-line leading");
        Assert(handler.CheckShapeTextOverflow(trailingEmptyParagraphsPath) == null,
            "trailing empty paragraphs should not create phantom overflow lines");
    }
    Console.WriteLine("PPTX render regression tests passed.");
}
finally
{
    if (File.Exists(path)) File.Delete(path);
}

static void CreateFixture(string path)
{
    BlankDocCreator.Create(path);
    using (var builder = new PowerPointHandler(path, editable: true))
    {
        builder.Add("/", "slide", position: null, properties: new Dictionary<string, string>());
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "preset alpha" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "system alpha", ["y"] = "3cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "local list style", ["y"] = "6cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "explicit no autofit", ["y"] = "9cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "compact single line", ["y"] = "10cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "trailing empty paragraphs", ["y"] = "11cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "wide chevron", ["y"] = "12cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "reverse gradient stops", ["y"] = "13cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "stroke-only custom path", ["y"] = "14cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "dashed custom path", ["y"] = "15cm" });
        builder.Add("/slide[1]", "connector", position: null,
            properties: new Dictionary<string, string>
            {
                ["x"] = "1cm",
                ["y"] = "16cm",
                ["width"] = "6cm",
                ["height"] = "0cm",
                ["lineColor"] = "705030",
                ["lineWidth"] = "0.25pt",
                ["tailEnd"] = "triangle"
            });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "no-fill text outline", ["y"] = "17cm" });
        builder.Add("/slide[1]", "textbox", position: null,
            properties: new Dictionary<string, string> { ["text"] = "master inheritance cascade", ["y"] = "18cm" });
        builder.Add("/slide[1]", "chart", position: null,
            properties: new Dictionary<string, string>
            {
                ["type"] = "line",
                ["categories"] = string.Join(',', Enumerable.Range(1, 12).Select(month =>
                    new DateTime(2024, month, 1).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture))),
                ["data"] = $"NAV:{string.Join(',', Enumerable.Range(1, 12))}",
                ["x"] = "12cm",
                ["y"] = "1cm",
                ["width"] = "8cm",
                ["height"] = "5cm"
            });
        builder.Save();
    }

    using (var document = PresentationDocument.Open(path, isEditable: true))
    {
        var slide = document.PresentationPart!.SlideParts.Single().Slide!;
        var shapes = slide.CommonSlideData!.ShapeTree!.Elements<Shape>().ToList();
        SetRunColor(shapes.Single(shape => shape.InnerText.Contains("preset", StringComparison.OrdinalIgnoreCase)),
            new Drawing.PresetColor { Val = Drawing.PresetColorValues.Black }, 85000);
        SetRunColor(shapes.Single(shape => shape.InnerText.Contains("system", StringComparison.OrdinalIgnoreCase)),
            new Drawing.SystemColor { Val = Drawing.SystemColorValues.WindowText, LastColor = "000000" }, 50000);
        SetLocalListStyle(shapes.Single(shape => shape.InnerText.Contains("local list style", StringComparison.OrdinalIgnoreCase)));
        ConfigureOverflowFixture(
            shapes.Single(shape => shape.InnerText.Contains("explicit no autofit", StringComparison.OrdinalIgnoreCase)),
            fontSize: 1200,
            heightPt: 12,
            lineSpacingPercent: null,
            explicitNoAutofit: true,
            trailingEmptyParagraphs: 0);
        ConfigureOverflowFixture(
            shapes.Single(shape => shape.InnerText.Contains("compact single line", StringComparison.OrdinalIgnoreCase)),
            fontSize: 1400,
            heightPt: 12.96,
            lineSpacingPercent: 95000,
            explicitNoAutofit: false,
            trailingEmptyParagraphs: 0);
        ConfigureOverflowFixture(
            shapes.Single(shape => shape.InnerText.Contains("trailing empty paragraphs", StringComparison.OrdinalIgnoreCase)),
            fontSize: 1000,
            heightPt: 12,
            lineSpacingPercent: null,
            explicitNoAutofit: false,
            trailingEmptyParagraphs: 2);
        ConfigureWideChevronFixture(
            shapes.Single(shape => shape.InnerText.Contains("wide chevron", StringComparison.OrdinalIgnoreCase)));
        ConfigureReverseGradientFixture(
            shapes.Single(shape => shape.InnerText.Contains("reverse gradient stops", StringComparison.OrdinalIgnoreCase)));
        ConfigureStrokeOnlyCustomPathFixture(
            shapes.Single(shape => shape.InnerText.Contains("stroke-only custom path", StringComparison.OrdinalIgnoreCase)));
        ConfigureDashedCustomPathFixture(
            shapes.Single(shape => shape.InnerText.Contains("dashed custom path", StringComparison.OrdinalIgnoreCase)));
        ConfigureNoFillTextOutlineFixture(
            shapes.Single(shape => shape.InnerText.Contains("no-fill text outline", StringComparison.OrdinalIgnoreCase)));
        ConfigureMasterInheritanceFixture(
            document,
            shapes.Single(shape => shape.InnerText.Contains("master inheritance cascade", StringComparison.OrdinalIgnoreCase)));
        ConfigureWpsDateAxis(document);
        slide.Save();
    }

    using var checker = new PowerPointHandler(path, editable: false);
    var validationErrors = checker.Validate();
    Assert(validationErrors.Count == 0,
        $"generated PPTX fixture should pass schema validation: {string.Join(" | ", validationErrors)}");
}

static void ConfigureMasterInheritanceFixture(PresentationDocument document, Shape slideShape)
{
    const uint placeholderIndex = 77;
    var slidePart = document.PresentationPart!.SlideParts.Single();
    var layoutPart = slidePart.SlideLayoutPart!;
    var masterPart = layoutPart.SlideMasterPart!;

    SetPlaceholder(slideShape, placeholderIndex);
    slideShape.TextBody!.GetFirstChild<Drawing.ListStyle>()?.Remove();
    slideShape.TextBody.InsertAfter(new Drawing.ListStyle(),
        slideShape.TextBody.GetFirstChild<Drawing.BodyProperties>());
    slideShape.TextBody.AppendChild(new Drawing.Paragraph(
        new Drawing.Run(new Drawing.Text("master cascade second paragraph"))));

    var layoutShape = (Shape)slideShape.CloneNode(true);
    SetShapeIdentity(layoutShape, 9001U, "Master cascade layout placeholder");
    ClearShapeText(layoutShape);
    SetPlaceholder(layoutShape, placeholderIndex);
    SetListStyle(layoutShape, new Drawing.Level1ParagraphProperties(
        new Drawing.DefaultRunProperties { Bold = true })
    {
        Alignment = Drawing.TextAlignmentTypeValues.Center,
    });
    layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!.AppendChild(layoutShape);

    var masterShape = (Shape)slideShape.CloneNode(true);
    SetShapeIdentity(masterShape, 9002U, "Master cascade master placeholder");
    ClearShapeText(masterShape);
    SetPlaceholder(masterShape, placeholderIndex);
    var masterPlaceholderRunProperties = new Drawing.DefaultRunProperties
    {
        Italic = true,
    };
    masterPlaceholderRunProperties.Append(
        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "123456" }),
        new Drawing.LatinFont { Typeface = "Master Placeholder" },
        new Drawing.EastAsianFont { Typeface = "Master Placeholder" },
        new Drawing.ComplexScriptFont { Typeface = "Master Placeholder" });
    SetListStyle(masterShape, new Drawing.Level1ParagraphProperties(masterPlaceholderRunProperties)
    {
        LeftMargin = 254000,
    });
    masterPart.SlideMaster!.CommonSlideData!.ShapeTree!.AppendChild(masterShape);

    var textStyles = masterPart.SlideMaster.TextStyles;
    if (textStyles == null)
    {
        textStyles = new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle());
        masterPart.SlideMaster.TextStyles = textStyles;
    }
    var bodyStyle = textStyles.BodyStyle ?? textStyles.AppendChild(new BodyStyle());
    var masterBodyLevel = bodyStyle.GetFirstChild<Drawing.Level1ParagraphProperties>();
    if (masterBodyLevel == null)
    {
        masterBodyLevel = new Drawing.Level1ParagraphProperties();
        bodyStyle.AppendChild(masterBodyLevel);
    }
    var bodyRunProperties = masterBodyLevel.GetFirstChild<Drawing.DefaultRunProperties>();
    if (bodyRunProperties == null)
    {
        bodyRunProperties = new Drawing.DefaultRunProperties();
        masterBodyLevel.AppendChild(bodyRunProperties);
    }
    bodyRunProperties.FontSize = 1700;
    masterBodyLevel.RemoveAllChildren<Drawing.SpaceAfter>();
    masterBodyLevel.InsertBefore(
        new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = 600 }),
        bodyRunProperties);

    var presentation = document.PresentationPart.Presentation!;
    var defaultTextStyle = presentation.DefaultTextStyle;
    if (defaultTextStyle == null)
    {
        defaultTextStyle = new DefaultTextStyle();
        presentation.DefaultTextStyle = defaultTextStyle;
    }
    var presentationLevel = defaultTextStyle.GetFirstChild<Drawing.Level1ParagraphProperties>();
    if (presentationLevel == null)
    {
        presentationLevel = new Drawing.Level1ParagraphProperties();
        defaultTextStyle.AppendChild(presentationLevel);
    }
    presentationLevel.RightMargin = 127000;

    layoutPart.SlideLayout.Save();
    masterPart.SlideMaster.Save();
    presentation.Save();
}

static void SetListStyle(Shape shape, Drawing.Level1ParagraphProperties level)
{
    shape.TextBody!.GetFirstChild<Drawing.ListStyle>()?.Remove();
    shape.TextBody.InsertAfter(new Drawing.ListStyle(level),
        shape.TextBody.GetFirstChild<Drawing.BodyProperties>());
}

static void SetPlaceholder(Shape shape, uint index)
{
    var applicationProperties = shape.NonVisualShapeProperties!
        .ApplicationNonVisualDrawingProperties!;
    applicationProperties.RemoveAllChildren<PlaceholderShape>();
    applicationProperties.AppendChild(new PlaceholderShape
    {
        Index = index,
        Type = PlaceholderValues.Body,
    });
}

static void SetShapeIdentity(Shape shape, uint id, string name)
{
    var properties = shape.NonVisualShapeProperties!.NonVisualDrawingProperties!;
    properties.Id = id;
    properties.Name = name;
}

static void ClearShapeText(Shape shape)
{
    foreach (var text in shape.Descendants<Drawing.Text>()) text.Text = string.Empty;
}

static void ConfigureNoFillTextOutlineFixture(Shape shape)
{
    var runProperties = shape.Descendants<Drawing.Run>().Single().RunProperties!;
    runProperties.AddChild(new Drawing.Outline(new Drawing.NoFill()), true);
}

static void ConfigureDashedCustomPathFixture(Shape shape)
{
    ConfigureStrokeOnlyCustomPathFixture(shape);
    shape.ShapeProperties!.AppendChild(new Drawing.Outline(
        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "806040" }),
        new Drawing.PresetDash { Val = Drawing.PresetLineDashValues.Dash })
    { Width = 12700 });
}

static void ConfigureStrokeOnlyCustomPathFixture(Shape shape)
{
    var shapeProperties = shape.ShapeProperties!;
    shapeProperties.RemoveAllChildren<Drawing.PresetGeometry>();
    shapeProperties.RemoveAllChildren<Drawing.SolidFill>();
    shapeProperties.RemoveAllChildren<Drawing.Outline>();
    shapeProperties.AppendChild(new Drawing.CustomGeometry(
        new Drawing.AdjustValueList(),
        new Drawing.ShapeGuideList(),
        new Drawing.AdjustHandleList(),
        new Drawing.ConnectionSiteList(),
        new Drawing.Rectangle { Left = "l", Top = "t", Right = "r", Bottom = "b" },
        new Drawing.PathList(
            new Drawing.Path(
                new Drawing.MoveTo(new Drawing.Point { X = "0", Y = "100000" }),
                new Drawing.LineTo(new Drawing.Point { X = "50000", Y = "0" }),
                new Drawing.LineTo(new Drawing.Point { X = "100000", Y = "50000" }))
            { Width = 100000, Height = 100000 })));
    shapeProperties.AppendChild(new Drawing.NoFill());
    shape.ShapeStyle = new ShapeStyle(
        "<p:style xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef>" +
        "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
        "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>" +
        "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef>" +
        "</p:style>");
}

static void ConfigureWideChevronFixture(Shape shape)
{
    var shapeProperties = shape.ShapeProperties!;
    shapeProperties.GetFirstChild<Drawing.PresetGeometry>()?.Remove();
    shapeProperties.AddChild(new Drawing.PresetGeometry(
        new Drawing.AdjustValueList())
    { Preset = Drawing.ShapeTypeValues.Chevron }, true);
    shapeProperties.Transform2D!.Extents!.Cx = 2880000;
    shapeProperties.Transform2D.Extents.Cy = 360000;
}

static void ConfigureReverseGradientFixture(Shape shape)
{
    var shapeProperties = shape.ShapeProperties!;
    shapeProperties.RemoveAllChildren<Drawing.SolidFill>();
    shapeProperties.RemoveAllChildren<Drawing.GradientFill>();
    shapeProperties.RemoveAllChildren<Drawing.NoFill>();

    var transparentWhite = new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.Background1 };
    transparentWhite.AppendChild(new Drawing.Alpha { Val = 0 });
    var gradient = new Drawing.GradientFill(
        new Drawing.GradientStopList(
            new Drawing.GradientStop(transparentWhite) { Position = 77000 },
            new Drawing.GradientStop(
                new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.Background1 })
            { Position = 30000 }),
        new Drawing.LinearGradientFill { Angle = 5400000, Scaled = true });
    shapeProperties.AddChild(gradient, true);
}

static void ConfigureOverflowFixture(
    Shape shape,
    int fontSize,
    double heightPt,
    int? lineSpacingPercent,
    bool explicitNoAutofit,
    int trailingEmptyParagraphs)
{
    var bodyProperties = shape.TextBody!.GetFirstChild<Drawing.BodyProperties>()!;
    bodyProperties.LeftInset = 0;
    bodyProperties.RightInset = 0;
    bodyProperties.TopInset = 0;
    bodyProperties.BottomInset = 0;
    bodyProperties.RemoveAllChildren<Drawing.NormalAutoFit>();
    bodyProperties.RemoveAllChildren<Drawing.ShapeAutoFit>();
    bodyProperties.RemoveAllChildren<Drawing.NoAutoFit>();
    if (explicitNoAutofit)
        bodyProperties.AppendChild(new Drawing.NoAutoFit());

    var paragraph = shape.TextBody.Elements<Drawing.Paragraph>().Single();
    var paragraphProperties = paragraph.ParagraphProperties ??
        paragraph.PrependChild(new Drawing.ParagraphProperties());
    paragraphProperties.RemoveAllChildren<Drawing.LineSpacing>();
    if (lineSpacingPercent.HasValue)
        paragraphProperties.AppendChild(new Drawing.LineSpacing(
            new Drawing.SpacingPercent { Val = lineSpacingPercent.Value }));
    foreach (var run in paragraph.Elements<Drawing.Run>())
    {
        run.RunProperties ??= new Drawing.RunProperties();
        run.RunProperties.FontSize = fontSize;
    }

    for (var i = 0; i < trailingEmptyParagraphs; i++)
        shape.TextBody.AppendChild(new Drawing.Paragraph());

    shape.ShapeProperties!.Transform2D!.Extents!.Cy =
        (long)Math.Round(heightPt * 12700);
}

static string FindShapePathByText(string path, string text)
{
    using var document = PresentationDocument.Open(path, isEditable: false);
    var shapes = document.PresentationPart!.SlideParts.Single().Slide!
        .CommonSlideData!.ShapeTree!.Elements<Shape>().ToList();
    var index = shapes.FindIndex(shape => shape.InnerText.Contains(text, StringComparison.OrdinalIgnoreCase));
    if (index < 0) throw new InvalidOperationException($"fixture shape not found: {text}");
    return $"/slide[1]/shape[{index + 1}]";
}

static void ConfigureWpsDateAxis(PresentationDocument document)
{
    var chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
    var plotArea = chartPart.ChartSpace!.Descendants<Charts.PlotArea>().Single();
    var categoryAxis = plotArea.GetFirstChild<Charts.CategoryAxis>()!;
    var dateAxis = new Charts.DateAxis();
    foreach (var child in categoryAxis.ChildElements)
    {
        if (child.LocalName is "lblAlgn" or "noMultiLvlLbl") continue;
        dateAxis.Append(child.CloneNode(true));
    }
    var numberingFormat = new Charts.NumberingFormat { FormatCode = "m/d/yyyy", SourceLinked = false };
    dateAxis.InsertAfter(numberingFormat, dateAxis.GetFirstChild<Charts.AxisPosition>());
    var textProperties = dateAxis.GetFirstChild<Charts.TextProperties>();
    if (textProperties == null)
    {
        textProperties = new Charts.TextProperties(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.ParagraphProperties(
                new Drawing.DefaultRunProperties { FontSize = 800 })));
        dateAxis.InsertBefore(textProperties, dateAxis.GetFirstChild<Charts.CrossingAxis>());
    }
    textProperties.GetFirstChild<Drawing.BodyProperties>()!.Rotation = -60000000;
    plotArea.InsertBefore(dateAxis, categoryAxis);
    categoryAxis.Remove();
    chartPart.ChartSpace.Save();
}

static void SetLocalListStyle(Shape shape)
{
    var run = shape.Descendants<Drawing.Run>().Single();
    run.RunProperties = new Drawing.RunProperties { Language = "zh-CN" };

    var defaultRunProperties = new Drawing.DefaultRunProperties { FontSize = 900 };
    defaultRunProperties.Append(
        new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.Background1 }),
        new Drawing.LatinFont { Typeface = "OPPOSans M" },
        new Drawing.EastAsianFont { Typeface = "OPPOSans M" },
        new Drawing.ComplexScriptFont { Typeface = "OPPOSans M" });
    var level = new Drawing.Level1ParagraphProperties(defaultRunProperties);
    var listStyle = new Drawing.ListStyle(level);

    shape.TextBody!.GetFirstChild<Drawing.ListStyle>()?.Remove();
    shape.TextBody.InsertAfter(listStyle, shape.TextBody.GetFirstChild<Drawing.BodyProperties>());
}

static void SetRunColor(Shape shape, OpenXmlElement color, int alpha)
{
    var run = shape.Descendants<Drawing.Run>().First();
    var runProperties = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
    runProperties.RemoveAllChildren<Drawing.SolidFill>();
    var solidFill = new Drawing.SolidFill(color);
    color.AppendChild(new Drawing.Alpha { Val = alpha });
    if (!runProperties.AddChild(solidFill, throwOnError: false))
        runProperties.AppendChild(solidFill);
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}

static int CountOccurrences(string value, string needle)
{
    var count = 0;
    var index = 0;
    while ((index = value.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
    {
        count++;
        index += needle.Length;
    }
    return count;
}
