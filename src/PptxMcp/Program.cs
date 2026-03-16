using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PptxMcp.Services;
using PptxMcp.Tools;

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
.WithTools<PptxTools>();

await builder.Build().RunAsync();
