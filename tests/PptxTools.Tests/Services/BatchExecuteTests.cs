using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using PptxTools.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTools.Tests.Services;

[Trait("Category", "Unit")]
public class BatchExecuteTests : PptxTestBase
{
    private static readonly byte[] MinimalPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nZxQAAAAASUVORK5CYII=");

    // ────────────────────────────────────────────────────────
    // UpdateText operations (backward compatibility)
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_UpdateText_WorksSameAsLegacyBatchUpdate()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Updated subtitle"),
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateText, NewText: "5.0M ARR"),
            new BatchOperation(3, "Risk Body", BatchOperationType.UpdateText, NewText: "New risk item")
        ]);

        Assert.Equal(3, result.TotalOperations);
        Assert.Equal(3, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.False(result.RolledBack);
        Assert.All(result.Results, r => Assert.True(r.Success));
        Assert.All(result.Results, r => Assert.Equal(BatchOperationType.UpdateText, r.Type));

        Assert.Equal("Updated subtitle", FindShapeText(path, 0, "Executive Subtitle"));
        Assert.Equal("5.0M ARR", FindShapeText(path, 1, "Revenue Value"));
    }

    [Fact]
    public void BatchExecute_UpdateText_MixedSuccessAndFailure()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Good update"),
            new BatchOperation(1, "NonExistent Shape", BatchOperationType.UpdateText, NewText: "Will fail"),
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateText, NewText: "6.0M ARR")
        ]);

        Assert.Equal(3, result.TotalOperations);
        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.True(result.Results[0].Success);
        Assert.False(result.Results[1].Success);
        Assert.NotNull(result.Results[1].Error);
        Assert.True(result.Results[2].Success);
    }

    // ────────────────────────────────────────────────────────
    // UpdateTableCell operations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_UpdateTableCell_SingleCell()
    {
        var path = CreateDeckWithTable("KPI Table", [["Metric", "Value"], ["ARR", "3.2M"]]);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "KPI Table", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 1, CellValue: "4.8M")
        ]);

        Assert.Equal(1, result.TotalOperations);
        Assert.Equal(1, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.True(result.Results[0].Success);
        Assert.Contains("[1,1]", result.Results[0].Detail);

        var slide = Service.GetSlideContent(path, 0);
        var table = Assert.Single(slide.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("4.8M", table.TableRows![1][1]);
    }

    [Fact]
    public void BatchExecute_UpdateTableCell_MultipleCellsSameTable()
    {
        var path = CreateDeckWithTable("Revenue", [["Region", "Q1", "Q2"], ["NA", "3.0M", "3.2M"], ["EMEA", "1.2M", "1.4M"]]);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Revenue", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 1, CellValue: "3.5M"),
            new BatchOperation(1, "Revenue", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 2, CellValue: "3.8M"),
            new BatchOperation(1, "Revenue", BatchOperationType.UpdateTableCell,
                TableRow: 2, TableColumn: 1, CellValue: "1.5M")
        ]);

        Assert.Equal(3, result.TotalOperations);
        Assert.Equal(3, result.SuccessCount);
        Assert.All(result.Results, r => Assert.True(r.Success));

        var slide = Service.GetSlideContent(path, 0);
        var table = Assert.Single(slide.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("3.5M", table.TableRows![1][1]);
        Assert.Equal("3.8M", table.TableRows[1][2]);
        Assert.Equal("1.5M", table.TableRows[2][1]);
    }

    [Theory]
    [InlineData(-1, 0, "out of range")]
    [InlineData(99, 0, "out of range")]
    [InlineData(0, -1, "out of range")]
    [InlineData(0, 99, "out of range")]
    public void BatchExecute_UpdateTableCell_InvalidRowOrColumn_FailsGracefully(int row, int col, string expectedError)
    {
        var path = CreateDeckWithTable("T1", [["A", "B"], ["1", "2"]]);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "T1", BatchOperationType.UpdateTableCell,
                TableRow: row, TableColumn: col, CellValue: "X")
        ]);

        Assert.Equal(1, result.TotalOperations);
        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains(expectedError, result.Results[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BatchExecute_UpdateTableCell_TableNotFound_FailsGracefully()
    {
        var path = CreateDeckWithTable("Actual Table", [["A"], ["1"]]);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Missing Table", BatchOperationType.UpdateTableCell,
                TableRow: 0, TableColumn: 0, CellValue: "X")
        ]);

        Assert.Equal(1, result.TotalOperations);
        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("Actual Table", result.Results[0].Error);
    }

    [Fact]
    public void BatchExecute_UpdateTableCell_MissingRowColumn_FailsGracefully()
    {
        var path = CreateDeckWithTable("T", [["A"], ["1"]]);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "T", BatchOperationType.UpdateTableCell, CellValue: "X")
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("required", result.Results[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // UpdateShapeProperties operations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_UpdateShapeProperties_Position()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateShapeProperties,
                X: Emu.Inches2, Y: Emu.Inches3)
        ]);

        Assert.Equal(1, result.SuccessCount);
        Assert.True(result.Results[0].Success);

        var (x, y, _, _) = ReadShapeTransform(path, 1, "Revenue Value");
        Assert.Equal(Emu.Inches2, x);
        Assert.Equal(Emu.Inches3, y);
    }

    [Fact]
    public void BatchExecute_UpdateShapeProperties_Size()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateShapeProperties,
                Width: Emu.Inches5, Height: Emu.Inches1_5)
        ]);

        Assert.Equal(1, result.SuccessCount);
        Assert.True(result.Results[0].Success);

        var (_, _, w, h) = ReadShapeTransform(path, 1, "Revenue Value");
        Assert.Equal(Emu.Inches5, w);
        Assert.Equal(Emu.Inches1_5, h);
    }

    [Fact]
    public void BatchExecute_UpdateShapeProperties_Rotation()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateShapeProperties,
                Rotation: 5_400_000) // 90 degrees
        ]);

        Assert.Equal(1, result.SuccessCount);
        Assert.True(result.Results[0].Success);

        var rotation = ReadShapeRotation(path, 1, "Revenue Value");
        Assert.Equal(5_400_000, rotation);
    }

    [Fact]
    public void BatchExecute_UpdateShapeProperties_PartialUpdate_OnlyX()
    {
        var path = CreateMetricDeck();
        var (origX, origY, _, _) = ReadShapeTransform(path, 1, "Revenue Value");

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateShapeProperties,
                X: Emu.Inches4)
        ]);

        Assert.True(result.Results[0].Success);

        var (newX, newY, _, _) = ReadShapeTransform(path, 1, "Revenue Value");
        Assert.Equal(Emu.Inches4, newX);
        Assert.Equal(origY, newY); // Y unchanged
    }

    [Fact]
    public void BatchExecute_UpdateShapeProperties_ShapeNotFound_FailsGracefully()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Ghost Shape", BatchOperationType.UpdateShapeProperties,
                X: 100)
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("Ghost Shape", result.Results[0].Error);
    }

    [Fact]
    public void BatchExecute_UpdateShapeProperties_NoProperties_FailsGracefully()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateShapeProperties)
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("At least one", result.Results[0].Error);
    }

    // ────────────────────────────────────────────────────────
    // ReplaceImage operations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_ReplaceImage_HappyPath()
    {
        var path = CreateDeckWithPicture("Logo");
        var imagePath = CreateTempPng();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Logo", BatchOperationType.ReplaceImage, ImagePath: imagePath)
        ]);

        Assert.Equal(1, result.SuccessCount);
        Assert.True(result.Results[0].Success);
        Assert.Contains("Logo", result.Results[0].Detail);
    }

    [Fact]
    public void BatchExecute_ReplaceImage_FileNotFound_FailsGracefully()
    {
        var path = CreateDeckWithPicture("Logo");

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Logo", BatchOperationType.ReplaceImage,
                ImagePath: @"C:\nonexistent\image.png")
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("not found", result.Results[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BatchExecute_ReplaceImage_TargetNotPicture_FailsGracefully()
    {
        var path = CreateMetricDeck();
        var imagePath = CreateTempPng();

        // "Executive Subtitle" is a text shape, not a picture
        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.ReplaceImage, ImagePath: imagePath)
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.NotNull(result.Results[0].Error);
    }

    [Fact]
    public void BatchExecute_ReplaceImage_MissingImagePath_FailsGracefully()
    {
        var path = CreateDeckWithPicture("Logo");

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Logo", BatchOperationType.ReplaceImage)
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("ImagePath", result.Results[0].Error);
    }

    // ────────────────────────────────────────────────────────
    // Mixed operations
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_MixedOperations_AllSucceed()
    {
        var path = CreateMixedDeck();
        var imagePath = CreateTempPng();
        var baselineErrors = ValidatePresentation(path);

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateText, NewText: "New subtitle"),
            new BatchOperation(1, "Data Table", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 1, CellValue: "99.9%"),
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateShapeProperties,
                X: Emu.Inches2),
            new BatchOperation(1, "Slide Image", BatchOperationType.ReplaceImage, ImagePath: imagePath)
        ]);

        Assert.Equal(4, result.TotalOperations);
        Assert.Equal(4, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.False(result.RolledBack);
        Assert.All(result.Results, r => Assert.True(r.Success));

        // Verify each change persisted
        Assert.Equal("New subtitle", FindShapeText(path, 0, "Subtitle"));

        var slide = Service.GetSlideContent(path, 0);
        var table = Assert.Single(slide.Shapes, s => s.ShapeType == "Table");
        Assert.Equal("99.9%", table.TableRows![1][1]);

        Assert.Equal(baselineErrors, ValidatePresentation(path));
    }

    [Fact]
    public void BatchExecute_MixedOperations_PartialFailure()
    {
        var path = CreateMixedDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateText, NewText: "Updated"),
            new BatchOperation(1, "Missing Table", BatchOperationType.UpdateTableCell,
                TableRow: 0, TableColumn: 0, CellValue: "X"),
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateShapeProperties,
                Width: Emu.Inches5)
        ]);

        Assert.Equal(3, result.TotalOperations);
        Assert.Equal(2, result.SuccessCount);
        Assert.Equal(1, result.FailureCount);
        Assert.True(result.Results[0].Success);
        Assert.False(result.Results[1].Success);
        Assert.True(result.Results[2].Success);
    }

    [Fact]
    public void BatchExecute_MixedOperations_CorrectTypesReported()
    {
        var path = CreateMixedDeck();
        var imagePath = CreateTempPng();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateText, NewText: "T"),
            new BatchOperation(1, "Data Table", BatchOperationType.UpdateTableCell,
                TableRow: 0, TableColumn: 0, CellValue: "V"),
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateShapeProperties, X: 100),
            new BatchOperation(1, "Slide Image", BatchOperationType.ReplaceImage, ImagePath: imagePath)
        ]);

        Assert.Equal(BatchOperationType.UpdateText, result.Results[0].Type);
        Assert.Equal(BatchOperationType.UpdateTableCell, result.Results[1].Type);
        Assert.Equal(BatchOperationType.UpdateShapeProperties, result.Results[2].Type);
        Assert.Equal(BatchOperationType.ReplaceImage, result.Results[3].Type);
    }

    [Fact]
    public void BatchExecute_MixedOperations_PresentationCompatible()
    {
        var path = CreateMixedDeck();
        var imagePath = CreateTempPng();
        var baselineErrors = ValidatePresentation(path);

        Service.BatchExecute(path,
        [
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateText, NewText: "Compatible check"),
            new BatchOperation(1, "Data Table", BatchOperationType.UpdateTableCell,
                TableRow: 1, TableColumn: 0, CellValue: "Updated"),
            new BatchOperation(1, "Subtitle", BatchOperationType.UpdateShapeProperties,
                Width: Emu.Inches4, Height: Emu.ThreeQuartersInch)
        ]);

        AssertPresentationCompatible(path);
        Assert.Equal(baselineErrors, ValidatePresentation(path));
    }

    // ────────────────────────────────────────────────────────
    // Atomic/Transaction semantics
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_Atomic_AllSucceed_ChangesPersisted()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Atomic update"),
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateText, NewText: "9.9M ARR")
        ], atomic: true);

        Assert.Equal(2, result.SuccessCount);
        Assert.False(result.RolledBack);

        Assert.Equal("Atomic update", FindShapeText(path, 0, "Executive Subtitle"));
        Assert.Equal("9.9M ARR", FindShapeText(path, 1, "Revenue Value"));
    }

    [Fact]
    public void BatchExecute_Atomic_OneFails_FileReverted()
    {
        var path = CreateMetricDeck();
        var originalText = FindShapeText(path, 0, "Executive Subtitle");

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Should be rolled back"),
            new BatchOperation(1, "NonExistent Shape", BatchOperationType.UpdateText, NewText: "Will fail")
        ], atomic: true);

        Assert.Equal(1, result.FailureCount);
        Assert.True(result.RolledBack);

        // Original text should be restored
        Assert.Equal(originalText, FindShapeText(path, 0, "Executive Subtitle"));
    }

    [Fact]
    public void BatchExecute_NonAtomic_OneFails_PartialChangesPersisted()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Partial persisted"),
            new BatchOperation(1, "NonExistent Shape", BatchOperationType.UpdateText, NewText: "Will fail")
        ], atomic: false);

        Assert.Equal(1, result.FailureCount);
        Assert.False(result.RolledBack);

        // Successful change should persist
        Assert.Equal("Partial persisted", FindShapeText(path, 0, "Executive Subtitle"));
    }

    [Fact]
    public void BatchExecute_Atomic_BackupFileCleanedUp()
    {
        var path = CreateMetricDeck();
        var backupPath = path + ".bak";

        Service.BatchExecute(path,
        [
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "Test")
        ], atomic: true);

        Assert.False(File.Exists(backupPath), "Backup file should be cleaned up after atomic operation");
    }

    // ────────────────────────────────────────────────────────
    // Edge cases
    // ────────────────────────────────────────────────────────

    [Fact]
    public void BatchExecute_EmptyOperations_ReturnsZeroCounts()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path, []);

        Assert.Equal(0, result.TotalOperations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.False(result.RolledBack);
        Assert.Empty(result.Results);
    }

    [Fact]
    public void BatchExecute_AllOperationsTargetSameSlide()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateText, NewText: "7.0M"),
            new BatchOperation(2, "Gross Margin", BatchOperationType.UpdateText, NewText: "75%"),
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateShapeProperties,
                X: Emu.Inches2)
        ]);

        Assert.Equal(3, result.TotalOperations);
        Assert.Equal(3, result.SuccessCount);

        Assert.Equal("7.0M", FindShapeText(path, 1, "Revenue Value"));
        Assert.Equal("75%", FindShapeText(path, 1, "Gross Margin"));
    }

    [Fact]
    public void BatchExecute_NonExistentSlide_FailsGracefully()
    {
        var path = CreateMinimalPptx();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(99, "SomeShape", BatchOperationType.UpdateText, NewText: "X")
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.NotNull(result.Results[0].Error);
    }

    [Fact]
    public void BatchExecute_NullOperations_ReturnsZeroCounts()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path, null!);

        Assert.Equal(0, result.TotalOperations);
        Assert.Equal(0, result.SuccessCount);
        Assert.Equal(0, result.FailureCount);
        Assert.Empty(result.Results);
    }

    [Fact]
    public void BatchExecute_ResultsInRequestOrder()
    {
        var path = CreateMetricDeck();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(3, "Risk Body", BatchOperationType.UpdateText, NewText: "Third"),
            new BatchOperation(1, "Executive Subtitle", BatchOperationType.UpdateText, NewText: "First"),
            new BatchOperation(2, "Revenue Value", BatchOperationType.UpdateText, NewText: "Second")
        ]);

        Assert.Equal(3, result.Results[0].SlideNumber);
        Assert.Equal("Risk Body", result.Results[0].ShapeName);
        Assert.Equal(1, result.Results[1].SlideNumber);
        Assert.Equal("Executive Subtitle", result.Results[1].ShapeName);
        Assert.Equal(2, result.Results[2].SlideNumber);
        Assert.Equal("Revenue Value", result.Results[2].ShapeName);
    }

    [Fact]
    public void BatchExecute_UpdateTableCell_SlideWithNoTables_FailsGracefully()
    {
        var path = CreateMinimalPptx();

        var result = Service.BatchExecute(path,
        [
            new BatchOperation(1, "SomeTable", BatchOperationType.UpdateTableCell,
                TableRow: 0, TableColumn: 0, CellValue: "X")
        ]);

        Assert.Equal(0, result.SuccessCount);
        Assert.False(result.Results[0].Success);
        Assert.Contains("no tables", result.Results[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    // ────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────

    private string CreateMetricDeck() =>
        CreatePptxWithSlides(
            new TestSlideDefinition
            {
                TitleText = "FY26 Metrics",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Executive Subtitle",
                        PlaceholderType = PlaceholderValues.SubTitle,
                        Paragraphs = ["Board dashboard"]
                    }
                ]
            },
            new TestSlideDefinition
            {
                TitleText = "Revenue Dashboard",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Revenue Value",
                        Paragraphs = ["3.2M ARR"],
                        X = Emu.OneInch,
                        Y = Emu.Inches2,
                        Width = Emu.Inches3,
                        Height = Emu.ThreeQuartersInch
                    },
                    new TestTextShapeDefinition
                    {
                        Name = "Gross Margin",
                        Paragraphs = ["62%"],
                        X = Emu.Inches4,
                        Y = Emu.Inches2,
                        Width = Emu.Inches2,
                        Height = Emu.ThreeQuartersInch
                    }
                ],
                IncludeImage = true
            },
            new TestSlideDefinition
            {
                TitleText = "Execution Risks",
                TextShapes =
                [
                    new TestTextShapeDefinition
                    {
                        Name = "Risk Body",
                        PlaceholderType = PlaceholderValues.Body,
                        Paragraphs = ["Support EMEA renewals"]
                    }
                ]
            });

    private string CreateDeckWithTable(string tableName, IReadOnlyList<IReadOnlyList<string>> rows) =>
        CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Table Slide",
            Tables =
            [
                new TestTableDefinition
                {
                    Name = tableName,
                    Rows = rows
                }
            ]
        });

    private string CreateDeckWithPicture(string pictureName)
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Image Slide",
            IncludeImage = true
        });

        // Rename the picture shape created by TestPptxHelper
        using var doc = PresentationDocument.Open(path, true);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var picture = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().First();
        picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name = pictureName;
        slidePart.Slide.Save();

        return path;
    }

    private string CreateMixedDeck()
    {
        var path = CreatePptxWithSlides(new TestSlideDefinition
        {
            TitleText = "Mixed Slide",
            TextShapes =
            [
                new TestTextShapeDefinition
                {
                    Name = "Subtitle",
                    Paragraphs = ["Original subtitle"],
                    X = Emu.OneInch,
                    Y = Emu.Inches1_5,
                    Width = Emu.Inches8,
                    Height = Emu.ThreeQuartersInch
                }
            ],
            Tables =
            [
                new TestTableDefinition
                {
                    Name = "Data Table",
                    Rows = [["Metric", "Value"], ["Uptime", "99.5%"]]
                }
            ],
            IncludeImage = true
        });

        // Name the picture "Slide Image"
        using var doc = PresentationDocument.Open(path, true);
        var slidePart = (SlidePart)doc.PresentationPart!.GetPartById(
            doc.PresentationPart.Presentation.SlideIdList!.Elements<SlideId>().First().RelationshipId!.Value!);
        var picture = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Picture>().FirstOrDefault();
        if (picture is not null)
        {
            picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name = "Slide Image";
            slidePart.Slide.Save();
        }

        return path;
    }

    private string CreateTempPng()
    {
        var imagePath = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".png");
        File.WriteAllBytes(imagePath, MinimalPng);
        TrackTempFile(imagePath);
        return imagePath;
    }

    private string FindShapeText(string path, int slideIndex, string shapeName)
    {
        var slide = Service.GetSlideContent(path, slideIndex);
        var shape = Assert.Single(slide.Shapes, s => s.Name == shapeName);
        return shape.Text;
    }

    private static (long X, long Y, long Width, long Height) ReadShapeTransform(string path, int slideIndex, string shapeName)
    {
        using var doc = PresentationDocument.Open(path, false);
        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideIds[slideIndex].RelationshipId!.Value!);

        foreach (var shape in slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (string.Equals(name, shapeName, StringComparison.OrdinalIgnoreCase))
            {
                var xfrm = shape.ShapeProperties?.Transform2D;
                return (
                    xfrm?.Offset?.X?.Value ?? 0,
                    xfrm?.Offset?.Y?.Value ?? 0,
                    xfrm?.Extents?.Cx?.Value ?? 0,
                    xfrm?.Extents?.Cy?.Value ?? 0);
            }
        }

        throw new InvalidOperationException($"Shape '{shapeName}' not found on slide {slideIndex}");
    }

    private static long ReadShapeRotation(string path, int slideIndex, string shapeName)
    {
        using var doc = PresentationDocument.Open(path, false);
        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideIds[slideIndex].RelationshipId!.Value!);

        foreach (var shape in slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>())
        {
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (string.Equals(name, shapeName, StringComparison.OrdinalIgnoreCase))
                return shape.ShapeProperties?.Transform2D?.Rotation?.Value ?? 0;
        }

        throw new InvalidOperationException($"Shape '{shapeName}' not found on slide {slideIndex}");
    }

    private static void AssertPresentationCompatible(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var presentation = Assert.IsType<Presentation>(presentationPart.Presentation);
        var slideIdList = Assert.IsType<SlideIdList>(presentation.SlideIdList);
        var slideIds = slideIdList.Elements<SlideId>().ToList();
        Assert.NotEmpty(slideIds);

        foreach (var slideId in slideIds)
        {
            var slidePart = Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!));
            var slide = Assert.IsType<Slide>(slidePart.Slide);
            Assert.NotNull(slide.CommonSlideData?.ShapeTree);
        }
    }

    private static List<string> ValidatePresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(document)
            .Select(error => $"{error.Path?.XPath ?? "<unknown>"}: {error.Description}")
            .ToList();
    }
}
