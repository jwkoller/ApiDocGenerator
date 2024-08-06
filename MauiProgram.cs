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
