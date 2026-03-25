using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxTools.Tests;

internal static class TestPptxHelper
{
    private static readonly byte[] SampleImageBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    public static void CreateMinimalPresentation(string filePath, string? titleText = "Test Slide") =>
        CreatePresentation(filePath,
        [
            new TestSlideDefinition
            {
                TitleText = titleText
            }
        ]);

    public static void CreatePresentation(string filePath, IReadOnlyList<TestSlideDefinition> slides)
    {
        using var doc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);

        var presentationPart = doc.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Title,
        };
        slideLayoutPart.SlideLayout.CommonSlideData!.Name = "Title Slide";

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(
                new SlideLayoutId
                {
                    Id = 2049,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                }));

        slideLayoutPart.AddPart(slideMasterPart);

        var slideIdList = new SlideIdList();
        uint nextSlideId = 256;

        foreach (var slideDefinition in slides)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            slidePart.Slide = BuildSlide(slidePart, slideDefinition);

            if (!string.IsNullOrWhiteSpace(slideDefinition.SpeakerNotesText))
                AddSpeakerNotes(slidePart, slideDefinition.SpeakerNotesText!);

            slideIdList.Append(new SlideId
            {
                Id = nextSlideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        presentationPart.Presentation = new Presentation(
            slideIdList,
            new SlideSize { Cx = (int)Emu.Inches10, Cy = (int)Emu.Inches7_5, Type = SlideSizeValues.Screen4x3 },
            new NotesSize { Cx = (int)Emu.Inches7_5, Cy = (int)Emu.Inches10 });

        var slideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        presentationPart.Presentation.InsertAt(slideMasterIdList, 0);
        presentationPart.Presentation.Save();
    }

    private static Slide BuildSlide(SlidePart slidePart, TestSlideDefinition slideDefinition)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        uint nextShapeId = 2;

        if (!string.IsNullOrWhiteSpace(slideDefinition.TitleText))
        {
            shapeTree.Append(CreateTextShape(
                nextShapeId++,
                "Title 1",
                [new TestParagraphDefinition { Text = slideDefinition.TitleText }],
                PlaceholderValues.CenteredTitle,
                Emu.HalfInch,
                Emu.Inches0_3,
                Emu.Inches9,
                Emu.ThreeQuartersInch));
        }

        long currentY = Emu.Inches1_5;
        foreach (var textShape in slideDefinition.TextShapes)
        {
            var paragraphDefinitions = textShape.ParagraphDefinitions.Count > 0
                ? textShape.ParagraphDefinitions
                : textShape.Paragraphs.Select(paragraph => new TestParagraphDefinition { Text = paragraph }).ToList();

            long height = textShape.Height ?? Math.Max(Emu.ThreeQuartersInch, Emu.ThreeEighthsInch * Math.Max(1, paragraphDefinitions.Count));
            shapeTree.Append(CreateTextShape(
                nextShapeId,
                string.IsNullOrWhiteSpace(textShape.Name) ? $"Text {nextShapeId}" : textShape.Name!,
                paragraphDefinitions,
                textShape.PlaceholderType,
                textShape.X ?? Emu.OneInch,
                textShape.Y ?? currentY,
                textShape.Width ?? Emu.Inches8,
                height));
            nextShapeId++;

            if (textShape.Y is null)
                currentY += height + Emu.QuarterInch;
        }

        foreach (var table in slideDefinition.Tables)
        {
            long height = table.Height ?? Emu.Inches1_5;
            shapeTree.Append(CreateTable(
                nextShapeId,
                string.IsNullOrWhiteSpace(table.Name) ? $"Table {nextShapeId}" : table.Name!,
                table,
                table.X ?? Emu.OneInch,
                table.Y ?? currentY,
                table.Width ?? Emu.Inches8,
                height));
            nextShapeId++;

            if (table.Y is null)
                currentY += height + Emu.QuarterInch;
        }

        var slide = new Slide(
            new CommonSlideData(shapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        if (slideDefinition.IncludeImage)
        {
            var imagePart = slidePart.AddImagePart(ImagePartType.Png);
            using var imageStream = new MemoryStream(SampleImageBytes);
            imagePart.FeedData(imageStream);

            shapeTree.Append(CreatePicture(
                nextShapeId++,
                slidePart.GetIdOfPart(imagePart),
                Emu.OneInch,
                currentY,
                Emu.Inches4,
                Emu.Inches3));
        }

        // Charts must be added after the Slide is assigned to the SlidePart so
        // AddNewPart<ChartPart>() can generate a relationship on the right part.
        foreach (var chart in slideDefinition.Charts)
        {
            long chartHeight = chart.Height ?? Emu.Inches3;
            var chartPart = slidePart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = BuildChartSpace(chart);
            chartPart.ChartSpace.Save();

            var relId = slidePart.GetIdOfPart(chartPart);
            var chartShapeId = nextShapeId++;
            shapeTree.Append(CreateChartFrame(
                chartShapeId,
                string.IsNullOrWhiteSpace(chart.Name) ? $"Chart {chartShapeId}" : chart.Name!,
                relId,
                chart.X ?? Emu.OneInch,
                chart.Y ?? currentY,
                chart.Width ?? Emu.Inches8,
                chartHeight));
        }

        return slide;
    }

    private static void AddSpeakerNotes(SlidePart slidePart, string speakerNotesText)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        notesSlidePart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    CreateSpeakerNotesShape(speakerNotesText))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        notesSlidePart.AddPart(slidePart);
        notesSlidePart.NotesSlide.Save();
    }

    private static Shape CreateSpeakerNotesShape(string speakerNotesText) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 2" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1U })),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(new A.Text(speakerNotesText)),
                    new A.EndParagraphRunProperties())));

    private static Shape CreateTextShape(
        uint shapeId,
        string name,
        IReadOnlyList<TestParagraphDefinition> paragraphs,
        PlaceholderValues? placeholderType,
        long x,
        long y,
        long width,
        long height)
    {
        var applicationProperties = placeholderType is null
            ? new ApplicationNonVisualDrawingProperties()
            : new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = placeholderType });

        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in paragraphs.DefaultIfEmpty(new TestParagraphDefinition()))
            textBody.Append(CreateParagraph(paragraph));

        return new Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                applicationProperties),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height })),
            textBody);
    }

    private static A.Paragraph CreateParagraph(TestParagraphDefinition paragraph)
    {
        var pptParagraph = new A.Paragraph();
        if (paragraph.IsBullet || paragraph.IsNumbered || paragraph.Level > 0)
        {
            var properties = new A.ParagraphProperties();
            if (paragraph.Level > 0)
                properties.Level = paragraph.Level;

            properties.Append(new A.CharacterBullet { Char = "•" });
            pptParagraph.Append(properties);
        }

        pptParagraph.Append(new A.Run(new A.Text(paragraph.Text ?? string.Empty)));
        pptParagraph.Append(new A.EndParagraphRunProperties());
        return pptParagraph;
    }

    private static P.GraphicFrame CreateTable(uint shapeId, string name, TestTableDefinition table, long x, long y, long width, long height)
    {
        var rows = table.Rows.Count == 0
            ? new List<List<string>> { new() { string.Empty } }
            : table.Rows.Select(row => row.Count == 0 ? new List<string> { string.Empty } : row.ToList()).ToList();
        var columnCount = rows.Max(row => row.Count);
        var rowHeight = Math.Max(Emu.ThreeEighthsInch, height / rows.Count);
        var columnWidth = Math.Max(1L, width / columnCount);

        var drawingTable = new A.Table(new A.TableProperties { FirstRow = true, BandRow = true });
        var tableGrid = drawingTable.AppendChild(new A.TableGrid());
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            tableGrid.Append(new A.GridColumn { Width = columnWidth });

        foreach (var row in rows)
        {
            var normalizedRow = row.Concat(Enumerable.Repeat(string.Empty, columnCount - row.Count)).ToList();
            var tableRow = new A.TableRow { Height = rowHeight };
            foreach (var cellText in normalizedRow)
            {
                tableRow.Append(new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(new A.Text(cellText ?? string.Empty)),
                            new A.EndParagraphRunProperties())),
                    new A.TableCellProperties()));
            }

            drawingTable.Append(tableRow);
        }

        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = width, Cy = height }),
            new A.Graphic(
                new A.GraphicData(drawingTable)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                }));
    }

    internal static Picture CreatePicture(uint shapeId, string relationshipId, long x, long y, long width, long height, string? name = null) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name ?? $"Picture {shapeId}" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

    private static P.GraphicFrame CreateChartFrame(uint shapeId, string name, string relId, long x, long y, long width, long height) =>
        new(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = width, Cy = height }),
            new A.Graphic(
                new A.GraphicData(new C.ChartReference { Id = relId })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }));

    private static C.ChartSpace BuildChartSpace(TestChartDefinition chart)
    {
        var plotArea = new C.PlotArea();

        var seriesList = chart.Series.Count == 0
            ? [new TestSeriesDefinition { Name = "Series 1", Values = [1.0, 2.0, 3.0] }]
            : chart.Series;

        var categories = chart.Categories.Count > 0
            ? chart.Categories.ToArray()
            : seriesList[0].Values.Select((_, i) => $"Category {i + 1}").ToArray();

        switch (chart.ChartType.ToUpperInvariant())
        {
            case "LINE":
                var lineChart = new C.LineChart(
                    new C.Grouping { Val = C.GroupingValues.Standard });
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    lineChart.Append(BuildLineChartSeries((uint)idx, ser, categories));
                plotArea.Append(lineChart);
                break;

            case "PIE":
                var pieChart = new C.PieChart();
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    pieChart.Append(BuildPieChartSeries((uint)idx, ser, categories));
                plotArea.Append(pieChart);
                break;

            case "BAR":
                var barChartH = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Bar },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered });
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    barChartH.Append(BuildBarChartSeries((uint)idx, ser, categories));
                plotArea.Append(barChartH);
                break;

            case "AREA":
                var areaChart = new C.AreaChart(new C.Grouping { Val = C.GroupingValues.Standard });
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    areaChart.Append(BuildAreaChartSeries((uint)idx, ser, categories));
                plotArea.Append(areaChart);
                break;

            case "DOUGHNUT":
                var doughnutChart = new C.DoughnutChart();
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    doughnutChart.Append(BuildPieChartSeries((uint)idx, ser, categories));
                plotArea.Append(doughnutChart);
                break;

            case "SCATTER":
                var scatterChart = new C.ScatterChart(new C.ScatterStyle { Val = C.ScatterStyleValues.Line });
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    scatterChart.Append(BuildScatterChartSeries((uint)idx, ser));
                plotArea.Append(scatterChart);
                break;

            default: // Column
                var colChart = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered });
                foreach (var (ser, idx) in seriesList.Select((s, i) => (s, i)))
                    colChart.Append(BuildBarChartSeries((uint)idx, ser, categories));
                plotArea.Append(colChart);
                break;
        }

        return new C.ChartSpace(new C.Chart(plotArea));
    }

    private static C.BarChartSeries BuildBarChartSeries(uint index, TestSeriesDefinition ser, string[] categories) =>
        new(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(ser.Name),
            BuildCategoryAxisData(categories),
            BuildValues(ser.Values));

    private static C.LineChartSeries BuildLineChartSeries(uint index, TestSeriesDefinition ser, string[] categories) =>
        new(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(ser.Name),
            BuildCategoryAxisData(categories),
            BuildValues(ser.Values));

    private static C.PieChartSeries BuildPieChartSeries(uint index, TestSeriesDefinition ser, string[] categories) =>
        new(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(ser.Name),
            BuildCategoryAxisData(categories),
            BuildValues(ser.Values));

    private static C.AreaChartSeries BuildAreaChartSeries(uint index, TestSeriesDefinition ser, string[] categories) =>
        new(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(ser.Name),
            BuildCategoryAxisData(categories),
            BuildValues(ser.Values));

    private static C.ScatterChartSeries BuildScatterChartSeries(uint index, TestSeriesDefinition ser)
    {
        // X values are generated as sequential integers (1, 2, 3, …) because the
        // test helper does not accept explicit X data — use ser.Values for Y.
        var xCache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)ser.Values.Count });
        for (uint i = 0; i < ser.Values.Count; i++)
            xCache.Append(new C.NumericPoint { Index = i, NumericValue = new C.NumericValue((i + 1).ToString()) });

        var yCache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)ser.Values.Count });
        for (uint i = 0; i < ser.Values.Count; i++)
            yCache.Append(new C.NumericPoint
            {
                Index = i,
                NumericValue = new C.NumericValue(ser.Values[(int)i].ToString(System.Globalization.CultureInfo.InvariantCulture))
            });

        return new C.ScatterChartSeries(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(ser.Name),
            new C.XValues(new C.NumberReference(new C.Formula(string.Empty), xCache)),
            new C.YValues(new C.NumberReference(new C.Formula(string.Empty), yCache)));
    }

    private static C.SeriesText BuildSeriesText(string? name) =>
        new(new C.StringReference(
            new C.Formula(string.Empty),
            new C.StringCache(
                new C.PointCount { Val = 1U },
                new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue(name ?? string.Empty) })));

    private static C.CategoryAxisData BuildCategoryAxisData(string[] categories)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)categories.Length });
        for (uint i = 0; i < categories.Length; i++)
            cache.Append(new C.StringPoint { Index = i, NumericValue = new C.NumericValue(categories[i]) });
        return new C.CategoryAxisData(new C.StringReference(new C.Formula(string.Empty), cache));
    }

    private static C.Values BuildValues(IReadOnlyList<double> values)
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Count });
        for (uint i = 0; i < values.Count; i++)
            cache.Append(new C.NumericPoint
            {
                Index = i,
                NumericValue = new C.NumericValue(values[(int)i].ToString(System.Globalization.CultureInfo.InvariantCulture))
            });
        return new C.Values(new C.NumberReference(new C.Formula(string.Empty), cache));
    }
}

public sealed class TestSlideDefinition
{
    public string? TitleText { get; init; }

    public string? SpeakerNotesText { get; init; }

    public IReadOnlyList<TestTextShapeDefinition> TextShapes { get; init; } = [];

    public IReadOnlyList<TestTableDefinition> Tables { get; init; } = [];

    public IReadOnlyList<TestChartDefinition> Charts { get; init; } = [];

    public bool IncludeImage { get; init; }
}

public sealed class TestTextShapeDefinition
{
    public string? Name { get; init; }

    public IReadOnlyList<string> Paragraphs { get; init; } = [];

    public IReadOnlyList<TestParagraphDefinition> ParagraphDefinitions { get; init; } = [];

    public PlaceholderValues? PlaceholderType { get; init; }

    public long? X { get; init; }

    public long? Y { get; init; }

    public long? Width { get; init; }

    public long? Height { get; init; }
}

public sealed class TestParagraphDefinition
{
    public string? Text { get; init; }

    public int Level { get; init; }

    public bool IsBullet { get; init; }

    public bool IsNumbered { get; init; }
}

public sealed class TestTableDefinition
{
    public string? Name { get; init; }

    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = [];

    public long? X { get; init; }

    public long? Y { get; init; }

    public long? Width { get; init; }

    public long? Height { get; init; }
}

public sealed class TestChartDefinition
{
    /// <summary>Shape name for the chart GraphicFrame. Defaults to "Chart {id}".</summary>
    public string? Name { get; init; }

    /// <summary>Chart type: Column (default), Bar, Line, Pie.</summary>
    public string ChartType { get; init; } = "Column";

    /// <summary>Category labels shared across all series. When empty, generated as "Category 1", "Category 2", …</summary>
    public IReadOnlyList<string> Categories { get; init; } = [];

    /// <summary>One or more series to include in the chart.</summary>
    public IReadOnlyList<TestSeriesDefinition> Series { get; init; } = [];

    public long? X { get; init; }

    public long? Y { get; init; }

    public long? Width { get; init; }

    public long? Height { get; init; }
}

public sealed class TestSeriesDefinition
{
    public string? Name { get; init; }

    public IReadOnlyList<double> Values { get; init; } = [1.0, 2.0, 3.0];
}
