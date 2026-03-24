using System.CommandLine;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PptxMcp.Commands;
using PptxMcp.Completions;
using PptxMcp.Prompts;
using PptxMcp.Resources;
using PptxMcp.Services;
using PptxMcp.Tools;

var mode = DetermineMode(args);
if (mode == "mcp")
{
    await RunMcpServerAsync(args);
    return 0;
}
else
{
    return await RunCliAsync(args);
}

static string DetermineMode(string[] args)
{
    if (args.Contains("--stdio"))
        return "mcp";

    if (args.Length == 0)
        return "cli";

    var first = args[0].ToLowerInvariant();
    if (first is "-h" or "--help" or "-v" or "--version")
        return "cli";

    string[] knownCommands = ["analyze", "optimize", "inspect", "export", "edit", "media", "slides"];
    if (knownCommands.Contains(first))
        return "cli";

    return "cli";
}

static async Task RunMcpServerAsync(string[] args)
{
    var builder = Host.CreateApplicationBuilder(args);

    builder.Logging.AddConsole(options =>
    {
        options.LogToStandardErrorThreshold = LogLevel.Trace;
    });

    builder.Services.AddSingleton<PresentationService>();

    builder.Services.AddMcpServer(options =>
    {
        options.ServerInfo = new ModelContextProtocol.Protocol.Implementation
        {
            Name = "pptx-mcp",
            Version = "1.0.0",
            Title = "PowerPoint MCP Server",
            Description = "MCP server for reading and modifying PowerPoint (.pptx) files using OpenXML SDK",
            WebsiteUrl = "https://github.com/jongalloway/pptx-mcp"
        };
    })
    .WithStdioServerTransport()
    .WithTools<PptxTools>()
    .WithResources<PptxResources>()
    .WithPrompts<PptxPrompts>()
    .WithCompleteHandler(PptxCompletionHandler.HandleAsync);

    await builder.Build().RunAsync();
}

static async Task<int> RunCliAsync(string[] args)
{
    var services = new ServiceCollection();
    services.AddSingleton<PresentationService>();
    var sp = services.BuildServiceProvider();
    var service = sp.GetRequiredService<PresentationService>();

    var rootCommand = new RootCommand("PowerPoint MCP - Analyze, optimize, and edit PowerPoint files");

    // Real commands
    rootCommand.Add(AnalyzeCommand.Create(service));
    rootCommand.Add(ExportCommand.Create(service));
    rootCommand.Add(InspectCommand.Create(service));
    rootCommand.Add(MediaCommand.Create(service));
    rootCommand.Add(SlidesCommand.Create(service));

    // Stubs for unimplemented commands
    (string name, string desc, int issue)[] stubs =
    [
        ("optimize", "Optimize presentation file size", 100),
        ("edit", "Edit presentation content", 103),
    ];

    foreach (var (name, desc, issue) in stubs)
    {
        var cmd = new Command(name, desc);
        cmd.SetAction(_ => Console.WriteLine($"Coming soon — see issue #{issue}"));
        rootCommand.Add(cmd);
    }

    return await rootCommand.Parse(args).InvokeAsync();
}
