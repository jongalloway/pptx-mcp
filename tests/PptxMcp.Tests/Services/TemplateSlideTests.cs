using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using PptxMcp.Models;

namespace PptxMcp.Tests.Services;

public class TemplateSlideTests : IDisposable
{
    private readonly PresentationService _service = new();
    private readonly List<string> _tempFiles = [];

    public void Dispose()
    {
        foreach (var file in _tempFiles)
            if (File.Exists(file)) File.Delete(file);
    }

    [Fact]
    public void AddSlideFromLayout_PopulatesRequestedPlaceholders()
    {
        var path = CreateTemplateDeck();

        var result = _service.AddSlideFromLayout(path, TemplateDeckHelper.TitleBodyLayoutName, new Dictionary<string, string>
        {
            ["Title"] = "Executive Summary",
            ["Body:1"] = "Revenue up 18%",
            ["Body:2"] = "Pipeline healthy"
        });

        Assert.True(result.Success);
        Assert.Equal(2, result.SlideNumber);
        Assert.Equal(TemplateDeckHelper.TitleBodyLayoutName, result.LayoutName);
        Assert.Equal(3, result.PlaceholdersPopulated);

        var slides = _service.GetSlides(path);
        Assert.Equal("Executive Summary", slides[1].Title);

        var addedSlide = _service.GetSlideContent(path, 1);
        Assert.Equal("Revenue up 18%", addedSlide.Shapes.Single(shape => shape.PlaceholderIndex == 1).Text);
        Assert.Equal("Pipeline healthy", addedSlide.Shapes.Single(shape => shape.PlaceholderIndex == 2).Text);
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void AddSlideFromLayout_WithoutPlaceholderValues_CreatesSlideWithLayoutRelationship()
    {
        var path = CreateTemplateDeck();
        var baselineErrors = ValidatePresentation(path);

        var result = _service.AddSlideFromLayout(path, TemplateDeckHelper.PictureCaptionLayoutName);

        Assert.True(result.Success);
        Assert.Equal(2, result.SlideNumber);
        Assert.Equal(TemplateDeckHelper.PictureCaptionLayoutName, result.LayoutName);
        Assert.Equal(0, result.PlaceholdersPopulated);

        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var slideIdList = Assert.IsType<SlideIdList>(presentationPart.Presentation.SlideIdList);
        var addedSlidePart = Assert.IsType<SlidePart>(presentationPart.GetPartById(slideIdList.Elements<SlideId>().Last().RelationshipId!.Value!));
        Assert.Equal(TemplateDeckHelper.PictureCaptionLayoutName, addedSlidePart.SlideLayoutPart?.SlideLayout.CommonSlideData?.Name?.Value);
        Assert.Equal(baselineErrors, ValidatePresentation(path));
    }

    [Fact]
    public void AddSlideFromLayout_ThrowsMeaningfulError_WhenLayoutIsMissing()
    {
        var path = CreateTemplateDeck();

        var exception = Assert.Throws<InvalidOperationException>(() => _service.AddSlideFromLayout(path, "Missing Layout"));

        Assert.Contains("Missing Layout", exception.Message);
        Assert.Contains(TemplateDeckHelper.TitleBodyLayoutName, exception.Message);
        Assert.Contains(TemplateDeckHelper.PictureCaptionLayoutName, exception.Message);
    }

    [Fact]
    public void AddSlideFromLayout_RejectsPicturePlaceholderTextOverrides()
    {
        var path = CreateTemplateDeck();

        var exception = Assert.Throws<InvalidOperationException>(() => _service.AddSlideFromLayout(path, TemplateDeckHelper.PictureCaptionLayoutName, new Dictionary<string, string>
        {
            ["Picture:1"] = "not-an-image"
        }));

        Assert.Contains("not text-capable", exception.Message);
        Assert.Contains("Picture:1", exception.Message);
    }

    [Fact]
    public void DuplicateSlide_WithOverrides_ClonesSlideAndKeepsSourceUntouched()
    {
        var path = CreateTemplateDeck();

        var result = _service.DuplicateSlide(path, 1, new Dictionary<string, string>
        {
            ["Title"] = "Duplicated Review",
            ["Body:2"] = "Action owners assigned"
        });

        Assert.True(result.Success);
        Assert.Equal(2, result.NewSlideNumber);
        Assert.Equal(5, result.ShapesCopied);
        Assert.Equal(2, result.OverridesApplied);

        var slides = _service.GetSlides(path);
        Assert.Equal("Quarterly Business Review", slides[0].Title);
        Assert.Equal("Duplicated Review", slides[1].Title);

        var duplicatedSlide = _service.GetSlideContent(path, 1);
        Assert.Equal("Action owners assigned", duplicatedSlide.Shapes.Single(shape => shape.PlaceholderIndex == 2).Text);

        var originalSlide = _service.GetSlideContent(path, 0);
        Assert.Equal("Follow-up items", originalSlide.Shapes.Single(shape => shape.PlaceholderIndex == 2).Text);
        AssertPresentationCompatible(path);
    }

    [Fact]
    public void DuplicateSlide_WithoutOverrides_ClonesImagePartIndependently()
    {
        var path = CreateTemplateDeck();
        var baselineErrors = ValidatePresentation(path);

        var result = _service.DuplicateSlide(path, 1);

        Assert.True(result.Success);
        Assert.Equal(2, result.NewSlideNumber);
        Assert.Equal(0, result.OverridesApplied);

        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var slideIdList = Assert.IsType<SlideIdList>(presentationPart.Presentation.SlideIdList);
        var slideParts = slideIdList.Elements<SlideId>()
            .Select(slideId => Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!)))
            .ToList();

        var sourceImagePart = Assert.Single(slideParts[0].ImageParts);
        var duplicatedImagePart = Assert.Single(slideParts[1].ImageParts);
        Assert.NotEqual(sourceImagePart.Uri, duplicatedImagePart.Uri);
        Assert.Equal(sourceImagePart.ContentType, duplicatedImagePart.ContentType);
        Assert.Equal(baselineErrors, ValidatePresentation(path));
    }

    private string CreateTemplateDeck()
    {
        var path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".pptx");
        _tempFiles.Add(path);
        TemplateDeckHelper.CreateTemplatePresentation(path);
        return path;
    }

    private static void AssertPresentationCompatible(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = Assert.IsType<PresentationPart>(document.PresentationPart);
        var slideIdList = Assert.IsType<SlideIdList>(presentationPart.Presentation.SlideIdList);
        Assert.All(slideIdList.Elements<SlideId>(), slideId =>
        {
            var slidePart = Assert.IsType<SlidePart>(presentationPart.GetPartById(slideId.RelationshipId!.Value!));
            Assert.NotNull(slidePart.Slide);
            Assert.NotNull(slidePart.SlideLayoutPart);
        });
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
