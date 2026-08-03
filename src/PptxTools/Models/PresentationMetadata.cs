namespace PptxTools.Models;

/// <summary>Presentation-level metadata extracted from package properties.</summary>
/// <param name="Title">Document title.</param>
/// <param name="Creator">Author / creator of the presentation.</param>
/// <param name="Created">Date the presentation was created.</param>
/// <param name="Modified">Date the presentation was last modified.</param>
/// <param name="Subject">Subject field.</param>
/// <param name="Keywords">Keywords / tags.</param>
/// <param name="Description">Description or comments.</param>
/// <param name="LastModifiedBy">Identity of the last person who saved changes.</param>
/// <param name="Category">Category field.</param>
/// <param name="SlideCount">Total number of slides in the presentation.</param>
public record PresentationMetadata(
    string? Title,
    string? Creator,
    DateTime? Created,
    DateTime? Modified,
    string? Subject,
    string? Keywords,
    string? Description,
    string? LastModifiedBy,
    string? Category,
    int SlideCount);
