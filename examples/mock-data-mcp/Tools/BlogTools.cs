using System.Text.Json;
using ModelContextProtocol.Server;

namespace MockDataMcp.Tools;

[McpServerToolType]
public sealed class BlogTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static readonly BlogPost[] Posts =
    [
        new(
            "2025-06-09",
            "Introducing MCP Composition Patterns for .NET Agents",
            "https://devblogs.microsoft.com/dotnet/mcp-composition-patterns",
            "Learn how to compose multiple MCP servers — pptx-tools, web browsing, and data APIs — " +
            "to build powerful multi-source agent workflows without writing custom glue code.",
            ["mcp", "dotnet", "ai-agents", "composition"]
        ),
        new(
            "2025-06-03",
            "What's New in .NET 10 Preview 4",
            "https://devblogs.microsoft.com/dotnet/dotnet-10-preview-4",
            "Preview 4 focuses on performance improvements in the JIT and GC, adds new LINQ overloads, " +
            "and ships the first stable bits of the new System.AI namespace.",
            ["dotnet", "dotnet10", "performance", "preview"]
        ),
        new(
            "2025-05-28",
            "Building Intelligent Agents with ModelContextProtocol SDK",
            "https://devblogs.microsoft.com/dotnet/building-agents-mcp-sdk",
            "A deep dive into the ModelContextProtocol SDK for .NET: tool registration, " +
            "stdio transport, and patterns for stateless, composable tools.",
            ["mcp", "dotnet", "ai-agents", "sdk"]
        ),
        new(
            "2025-05-21",
            "OpenXML SDK 3.4: Faster, Leaner, More Compatible",
            "https://devblogs.microsoft.com/dotnet/openxml-sdk-3-4",
            "OpenXML SDK 3.4 ships with significant memory reductions for large presentations, " +
            "improved validation accuracy, and a new fluent builder API for common shapes.",
            ["openxml", "dotnet", "powerpoint", "sdk"]
        ),
        new(
            "2025-05-14",
            "AI-Assisted Documentation: From Deck to Markdown in One Prompt",
            "https://devblogs.microsoft.com/dotnet/ai-docs-deck-to-markdown",
            "Walk through a real workflow where an AI agent reads a training presentation " +
            "and generates a structured markdown doc ready for your knowledge base.",
            ["mcp", "pptx-tools", "documentation", "ai-agents"]
        )
    ];

    /// <summary>
    /// Get recent blog post summaries. Optionally filter by tag (e.g. "mcp", "dotnet", "ai-agents").
    /// Returns titles, URLs, publish dates, and one-paragraph summaries — useful for
    /// updating a "What's New" or "Recent Updates" slide in a deck.
    /// </summary>
    /// <param name="tag">Optional tag filter. Returns only posts tagged with this value.</param>
    /// <param name="count">Maximum number of posts to return. Defaults to 5.</param>
    [McpServerTool(Title = "Get Latest Blog Posts", ReadOnly = true, Idempotent = true)]
    public Task<string> get_latest_blog_posts(string? tag = null, int count = 5)
    {
        var results = Posts.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(tag))
            results = results.Where(p => p.Tags.Contains(tag, StringComparer.OrdinalIgnoreCase));
        results = results.Take(Math.Max(1, count));

        var payload = new
        {
            fetched_at = DateTimeOffset.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
            filter_tag = tag,
            posts = results.Select(p => new
            {
                p.Published,
                p.Title,
                p.Url,
                p.Summary,
                p.Tags
            })
        };

        return Task.FromResult(JsonSerializer.Serialize(payload, JsonOptions));
    }

    private record BlogPost(string Published, string Title, string Url, string Summary, string[] Tags);
}
