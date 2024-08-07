using APIDocGenerator.ViewModels;
using CommunityToolkit.Maui.Alerts;
using Microsoft.Extensions.Logging;

namespace APIDocGenerator
{
    public partial class MainPage : ContentPage
    {
        private readonly ILogger<MainPage> _logger;
        private MainViewModel _viewModel;

        public MainPage(ILogger<MainPage> logger, MainViewModel viewModel)
        {
            InitializeComponent();
            BindingContext = viewModel;

            _viewModel = viewModel;
            _logger = logger;
        }

        private void SourceFolderPathCompletedEvent(object sender, EventArgs e)
        {
            string path = ((Entry)sender).Text;
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
                _logger.LogError(ex.Message);
                await DisplayAlert("Error", $"Document creation failed: {ex.Message}.", "Ok");
            }
        }

        private void FileNameCompletedEvent(object sender, EventArgs e)
        {
            string fileName = ((Entry)sender).Text;
            _viewModel.FileName = fileName;
        }
    }

}
