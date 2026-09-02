// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli;
using OfficeCli.Handlers;
using Drawing = DocumentFormat.OpenXml.Drawing;

var path = Path.Combine(Path.GetTempPath(), $"officecli-pptx-render-{Guid.NewGuid():N}.pptx");
try
{
    CreateFixture(path);
    using (var handler = new PowerPointHandler(path, editable: false))
    {
        var html = handler.ViewAsHtml(startSlide: 1, endSlide: 1);

        Assert(html.Contains("rgba(0,0,0,0.85)", StringComparison.Ordinal),
            "preset color with alpha should render as rgba");
        Assert(html.Contains("rgba(0,0,0,0.5)", StringComparison.Ordinal),
            "system color with alpha should render as rgba");
    }
    Console.WriteLine("PPTX preset/system color alpha render tests passed.");
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
        slide.Save();
    }

    using var checker = new PowerPointHandler(path, editable: false);
    Assert(checker.Validate().Count == 0, "generated PPTX fixture should pass schema validation");
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
