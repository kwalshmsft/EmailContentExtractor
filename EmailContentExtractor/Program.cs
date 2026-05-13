using EmailContentExtractor.Components;
using EmailContentExtractor.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddSingleton<HtmlLocalizationService>();
builder.Services.AddSingleton<FileStorageService>();
builder.Services.AddSingleton<IrisContentImportService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

// API endpoints to serve stored files
app.MapGet("api/files/sources/{fileName}", (string fileName, FileStorageService storage) =>
{
    var content = storage.GetSourceContent(fileName);
    if (content == null) return Results.NotFound();
    return Results.Content(content, "text/html");
});

app.MapGet("api/files/generated/{fileName}", (string fileName, FileStorageService storage) =>
{
    var content = storage.GetGeneratedContent(fileName);
    if (content == null) return Results.NotFound();
    return Results.Content(content, "text/html");
});

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
