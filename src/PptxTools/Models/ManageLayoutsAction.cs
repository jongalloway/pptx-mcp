namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_manage_layouts tool.</summary>
public enum ManageLayoutsAction
{
    /// <summary>Find unused slide layouts and masters with estimated space savings.</summary>
    Find,

    /// <summary>Remove unused slide layouts and orphaned masters from the presentation.</summary>
    Remove,

    /// <summary>Set the semantic type attribute on a named slide layout.</summary>
    SetType,

    /// <summary>Modify the type/index of an existing placeholder on a named slide layout.</summary>
    ModifyPlaceholder,

    /// <summary>Add a new placeholder shape to a named slide layout.</summary>
    AddPlaceholder
}
