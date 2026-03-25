namespace PptxTools.Models;

/// <summary>Actions for the consolidated pptx_chart_data tool.</summary>
public enum ChartDataAction
{
    /// <summary>Read chart data (series names, categories, and values) from an existing chart shape.</summary>
    Read,

    /// <summary>Update chart data values in an existing chart shape while preserving all styling and formatting.</summary>
    Update
}
