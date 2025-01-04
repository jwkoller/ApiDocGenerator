using APIDocGenerator.Services;
using APIDocGenerator.ViewModels;
using CommunityToolkit.Maui;
using CommunityToolkit.Maui.Storage;
using MetroLog.MicrosoftExtensions;
using Microsoft.Extensions.Logging;

namespace APIDocGenerator
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>()
                .UseMauiCommunityToolkit()
                .ConfigureFonts(fonts =>
                {
                    fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                    fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
                });
            builder.Services.AddSingleton<MainPage>();
            builder.Services.AddSingleton<MainViewModel>();
            builder.Services.AddSingleton<IFolderPicker>(FolderPicker.Default);
            builder.Services.AddSingleton<IFilePicker>(FilePicker.Default);

            builder.Services.AddScoped<FileReaderService>();
            builder.Services.AddScoped<TextParserService>();

            builder.Logging.AddStreamingFileLogger(options =>
            {
                options.FolderPath = $"{AppDomain.CurrentDomain.BaseDirectory}\\__Logs";
                options.MinLevel = LogLevel.Information;
                options.RetainDays = 15;
            });
            
#if DEBUG
            builder.Logging.AddTraceLogger(options =>
            {
                options.MinLevel = LogLevel.Debug;
            });
#endif

            return builder.Build();
        }
    }
}
