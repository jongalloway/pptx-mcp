using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using MockDataMcp.Tools;

var builder = Host.CreateApplicationBuilder(args);

builder.Logging.AddConsole(options =>
{
    options.LogToStandardErrorThreshold = LogLevel.Trace;
});

builder.Services.AddMcpServer(options =>
{
    options.ServerInfo = new ModelContextProtocol.Protocol.Implementation
    {
        Name = "mock-data-mcp",
        Version = "1.0.0",
        Title = "Mock Data MCP Server",
        Description = "Sample MCP server providing mock business metrics and blog posts. Use with pptx-mcp to demonstrate multi-source composition.",
        WebsiteUrl = "https://github.com/jongalloway/pptx-mcp"
    };
})
.WithStdioServerTransport()
.WithTools<MetricsTools>()
.WithTools<BlogTools>();

await builder.Build().RunAsync();
