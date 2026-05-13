using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using EmailContentExtractor.Components;
using EmailContentExtractor.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddSingleton<HtmlLocalizationService>();
builder.Services.AddSingleton<FileStorageService>();
builder.Services.AddSingleton<IrisContentImportService>();

await builder.Build().RunAsync();
