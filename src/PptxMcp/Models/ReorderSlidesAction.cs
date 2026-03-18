namespace PptxMcp.Models;

/// <summary>Actions for the consolidated pptx_reorder_slides tool.</summary>
public enum ReorderSlidesAction
{
    /// <summary>Move a single slide to a new position.</summary>
    Move,

    /// <summary>Reorder all slides by providing the complete new sequence as a 1-based array.</summary>
    Reorder
}
