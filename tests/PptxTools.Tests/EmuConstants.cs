namespace PptxTools.Tests;

/// <summary>
/// Named constants for English Metric Unit (EMU) values used in OpenXML PowerPoint test fixtures.
/// </summary>
/// <remarks>
/// 1 inch = 914,400 EMU. These constants replace hardcoded magic numbers
/// to improve readability and make layout intent explicit.
/// </remarks>
internal static class Emu
{
    /// <summary>0.25 inches (228,600 EMU). Common vertical gap between shapes.</summary>
    public const long QuarterInch = 228_600;

    /// <summary>0.3 inches (274,320 EMU). Standard title top-margin offset.</summary>
    public const long Inches0_3 = 274_320;

    /// <summary>0.375 inches (342,900 EMU). Minimum row height for tables.</summary>
    public const long ThreeEighthsInch = 342_900;

    /// <summary>0.5 inches (457,200 EMU). Half-inch offset used for title X positions.</summary>
    public const long HalfInch = 457_200;

    /// <summary>0.75 inches (685,800 EMU). Standard title shape height.</summary>
    public const long ThreeQuartersInch = 685_800;

    /// <summary>1 inch (914,400 EMU). Base conversion factor and common left margin.</summary>
    public const long OneInch = 914_400;

    /// <summary>1.25 inches (1,143,000 EMU). Body shape height for compact content.</summary>
    public const long Inches1_25 = 1_143_000;

    /// <summary>1.5 inches (1,371,600 EMU). Standard body content height and Y start.</summary>
    public const long Inches1_5 = 1_371_600;

    /// <summary>1.75 inches (1,600,200 EMU). Layout body Y position below title.</summary>
    public const long Inches1_75 = 1_600_200;

    /// <summary>2 inches (1,828,800 EMU). Used for shape widths and picture heights.</summary>
    public const long Inches2 = 1_828_800;

    /// <summary>2.5 inches (2,286,000 EMU). Source slide picture width.</summary>
    public const long Inches2_5 = 2_286_000;

    /// <summary>3 inches (2,743,200 EMU). Standard picture/shape dimension.</summary>
    public const long Inches3 = 2_743_200;

    /// <summary>3.5 inches (3,200,400 EMU). Layout body secondary Y position.</summary>
    public const long Inches3_5 = 3_200_400;

    /// <summary>4 inches (3,657,600 EMU). Standard picture width and shape offset.</summary>
    public const long Inches4 = 3_657_600;

    /// <summary>5 inches (4,572,000 EMU). Table X position for dashboard layouts.</summary>
    public const long Inches5 = 4_572_000;

    /// <summary>5.5 inches (5,029,200 EMU). Right-column shape X position.</summary>
    public const long Inches5_5 = 5_029_200;

    /// <summary>6 inches (5,486,400 EMU). Source slide picture X position.</summary>
    public const long Inches6 = 5_486_400;

    /// <summary>7.5 inches (6,858,000 EMU). Standard slide height (4:3 portrait) and notes width.</summary>
    public const long Inches7_5 = 6_858_000;

    /// <summary>8 inches (7,315,200 EMU). Default body content width.</summary>
    public const long Inches8 = 7_315_200;

    /// <summary>9 inches (8,229,600 EMU). Full-width title shape span.</summary>
    public const long Inches9 = 8_229_600;

    /// <summary>10 inches (9,144,000 EMU). Standard slide width (4:3 landscape).</summary>
    public const long Inches10 = 9_144_000;
}
