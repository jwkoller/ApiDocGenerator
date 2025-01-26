using APIDocGenerator.ViewModels;
using CommunityToolkit.Maui.Alerts;
using Microsoft.Extensions.Logging;

namespace APIDocGenerator
{
    public partial class MainPage : ContentPage
    {
        private readonly ILogger<MainPage> _logger;
        private readonly MainViewModel _viewModel;
        private readonly string[] _commandLineArgs;
        public MainPage(ILogger<MainPage> logger, MainViewModel viewModel)
        {
            _viewModel = viewModel;
            _commandLineArgs = Environment.GetCommandLineArgs();
            _logger = logger;

            if (_commandLineArgs.Length == 4)
            {
                RunCommandLineArgs();
            }
            else
            {
                InitializeComponent();
                BindingContext = viewModel;
            }
        }

        private async void RunCommandLineArgs()
        {
            string source = _commandLineArgs[1];
            string target = _commandLineArgs[2];
            string name = _commandLineArgs[3];

            bool sourceIsJsonFile = source.LastIndexOf(".json") >= 0;

            if ((sourceIsJsonFile && !File.Exists(source)) || (!sourceIsJsonFile && !Directory.Exists(source)))
            {
                _logger.LogError("Source file or directory \"{source}\" is invalid.", source);
            }
            else if (!Directory.Exists(target)) 
            {
                _logger.LogError("Target directory \"{target}\" is invalid.", target);
            }
            else
            {
                _viewModel.SelectedSource = source;
                _viewModel.SelectedDestination = target;
                _viewModel.FileName = name;
                _viewModel.UseJsonFile = sourceIsJsonFile;

                try
                {
                    await _viewModel.GenerateDocument();
                    _logger.LogInformation("\"{name}.docx\" created successfully at {datetime}.", name, DateTime.Now.ToString("u"));
                } 
                catch(Exception ex)
                {
                    _logger.LogError("Command line run failed to generate document: {ex}", ex);
                }
            } 

            Application.Current?.Quit();
        }

        private void SourceFolderPathCompletedEvent(object sender, EventArgs e)
        {
            string path = ((Entry)sender).Text;
            _viewModel.SelectedSource = path;
        }

        private void SourceJsonFilePathCompletedEvent(object sender, EventArgs e)
        {
            string path = (((Entry)sender).Text);
            _viewModel.SelectedSource = path;
        }

        private void DestinationFolderPathCompletedEvent(object sender, EventArgs e)
        {
            string path = ((Entry)sender).Text;
            _viewModel.SelectedDestination = path;
        }

        private async void GenerateDocumentEvent(object sender, EventArgs e)
        {
            try
            {
                await _viewModel.GenerateDocument();
                await DisplayAlert("Success", $"{_viewModel.FileName}.docx created.", "Ok");
            } catch (Exception ex)
            {
                _logger.LogError("{ex}", ex);
                await DisplayAlert("Error", $"Document creation failed: {ex.Message}.", "Ok");
            }
        }

        private void FileNameCompletedEvent(object sender, EventArgs e)
        {
            string fileName = ((Entry)sender).Text;
            _viewModel.FileName = fileName;
        }

        private void OnSourceTypeRadioButtonChanged(object sender, CheckedChangedEventArgs e)
        {
            _viewModel.SwapSourceSelectionOption();
        }
    }

}
