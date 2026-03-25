namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_manage_slides tool.</summary>
public enum ManageSlidesAction
{
    /// <summary>Add a blank slide, optionally specifying a layout name.</summary>
    Add,

    /// <summary>Create a slide from a named layout with optional placeholder population.</summary>
    AddFromLayout,

    /// <summary>Duplicate an existing slide with optional placeholder overrides.</summary>
    Duplicate
}
