namespace PptxMcp.Models;

/// <summary>Actions for the consolidated pptx_manage_layouts tool.</summary>
public enum ManageLayoutsAction
{
    /// <summary>Find unused slide layouts and masters with estimated space savings.</summary>
    Find,

    /// <summary>Remove unused slide layouts and orphaned masters from the presentation.</summary>
    Remove
}
