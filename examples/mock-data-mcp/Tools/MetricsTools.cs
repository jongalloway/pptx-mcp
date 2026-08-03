using System.Text.Json;
using ModelContextProtocol.Server;

namespace MockDataMcp.Tools;

[McpServerToolType]
public sealed class MetricsTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    /// Get weekly business KPIs: ARR, MRR, NRR, new logos, churn rate, and notable highlights.
    /// Returns mock data suitable for updating a board metrics slide.
    /// </summary>
    /// <param name="week">ISO week identifier (e.g. "2025-W24"). Defaults to the current week.</param>
    [McpServerTool(Title = "Get Weekly Metrics", ReadOnly = true, Idempotent = true)]
    public Task<string> get_weekly_metrics(string? week = null)
    {
        // Seed deterministic variation from week number so repeated calls return consistent data
        // while different weeks produce different values.
        var seed = ParseWeekSeed(week);

        var metrics = new
        {
            week = week ?? CurrentIsoWeek(),
            period = WeekLabel(seed),
            kpis = new
            {
                arr_millions = Math.Round(12.4 + seed * 0.3, 1),
                arr_change_pct = Math.Round(2.1 + (seed % 5) * 0.4, 1),
                mrr_thousands = (int)(1033 + seed * 25),
                nrr_pct = 109 + (seed % 6),
                new_logos = 4 + (seed % 5),
                churn_rate_pct = Math.Round(1.9 - (seed % 3) * 0.1, 1)
            },
            highlights = new[]
            {
                $"Closed {2 + seed % 4} enterprise deals in EMEA",
                $"NRR reached {109 + seed % 6}% — {(seed % 2 == 0 ? "second" : "third")} consecutive month above 109%",
                "Support ticket volume down 18% following documentation refresh",
                $"New integration: {new[] { "Slack", "Teams", "Notion", "Jira", "GitHub" }[seed % 5]} connector released"
            },
            last_updated = DateTimeOffset.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        };

        return Task.FromResult(JsonSerializer.Serialize(metrics, JsonOptions));
    }

    /// <summary>
    /// Get department-level updates for the weekly board presentation.
    /// Returns mock summaries for Engineering, Sales, Support, and Marketing.
    /// </summary>
    /// <param name="week">ISO week identifier (e.g. "2025-W24"). Defaults to the current week.</param>
    [McpServerTool(Title = "Get Team Updates", ReadOnly = true, Idempotent = true)]
    public Task<string> get_team_updates(string? week = null)
    {
        var seed = ParseWeekSeed(week);

        var updates = new
        {
            week = week ?? CurrentIsoWeek(),
            departments = new[]
            {
                new
                {
                    name = "Engineering",
                    status = "on-track",
                    updates = new[]
                    {
                        $"Shipped v{2 + seed % 2}.{seed % 10}.0 with {3 + seed % 4} bug fixes",
                        "Performance: p99 API latency down 12% vs last week",
                        $"Test coverage at {87 + seed % 8}% — up 2pp this sprint"
                    }
                },
                new
                {
                    name = "Sales",
                    status = "ahead",
                    updates = new[]
                    {
                        $"Pipeline: ${(14 + seed % 8) * 100}K in late-stage deals",
                        $"Win rate: {38 + seed % 10}% (vs {33 + seed % 7}% 30-day avg)",
                        "3 pilots converting to paid this week"
                    }
                },
                new
                {
                    name = "Support",
                    status = "on-track",
                    updates = new[]
                    {
                        $"CSAT: {4.2 + (seed % 4) * 0.1:F1}/5.0",
                        $"Median first-response: {2 + seed % 3}h (SLA: 4h)",
                        "Top issue: onboarding flow — eng fix in review"
                    }
                },
                new
                {
                    name = "Marketing",
                    status = "on-track",
                    updates = new[]
                    {
                        $"Blog post: '{new[] { "AI Agent Patterns", "MCP in Production", "OpenXML Deep Dive" }[seed % 3]}' — {1200 + seed * 200} views",
                        $"Webinar signup: {310 + seed * 20} registrations",
                        "Social: +12% follower growth vs last week"
                    }
                }
            }
        };

        return Task.FromResult(JsonSerializer.Serialize(updates, JsonOptions));
    }

    private static int ParseWeekSeed(string? week)
    {
        if (week is null) return (int)(DateTimeOffset.UtcNow.DayOfYear / 7);
        // Extract week number from ISO format "YYYY-Www"
        var parts = week.Split('-');
        if (parts.Length == 2 && parts[1].StartsWith('W') && int.TryParse(parts[1][1..], out var w))
            return w;
        return (int)(DateTimeOffset.UtcNow.DayOfYear / 7);
    }

    private static string CurrentIsoWeek()
    {
        var now = DateTimeOffset.UtcNow;
        var week = System.Globalization.ISOWeek.GetWeekOfYear(now.DateTime);
        return $"{now.Year}-W{week:D2}";
    }

    private static string WeekLabel(int seed)
    {
        // Generate a plausible week label
        var monday = new DateTimeOffset(2025, 1, 6, 0, 0, 0, TimeSpan.Zero).AddDays((seed - 2) * 7);
        var sunday = monday.AddDays(6);
        return $"Week of {monday:MMM d}–{sunday:MMM d, yyyy}";
    }
}
