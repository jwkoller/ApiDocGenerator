using APIDocGenerator.Services;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;


namespace APIDocGenerator.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger<MainViewModel> _logger;
        private readonly IFolderPicker _folderPicker;

        [ObservableProperty]
        private string _selectedSource = string.Empty;
        [ObservableProperty]
        private string _selectedDestination = string.Empty;
        [ObservableProperty]
        private string _fileName = "apiDoc";

        public MainViewModel(ILogger<MainViewModel> logger, IFolderPicker folderPicker)
        {
            _logger = logger;
            _folderPicker = folderPicker;
        }

        [RelayCommand]
        public async Task SelectSourceFolder(CancellationToken token)
        {
            FolderPickerResult result = await _folderPicker.PickAsync(token);
            result.EnsureSuccess();

            SelectedSource = result.Folder.Path;
        }

        [RelayCommand]
        public async Task SelectDestinationFolder(CancellationToken token)
        {
            FolderPickerResult result = await _folderPicker.PickAsync(token);
            result.EnsureSuccess();

            SelectedDestination = result.Folder.Path;
        }

        public Task GenerateDocument()
        {
            IEnumerable<FileInfo> sourceFiles = FileReaderService.GetFiles(SelectedSource);
            DocumentGenerator docGenerator = new DocumentGenerator(SelectedDestination, FileName);

            foreach(FileInfo file in sourceFiles)
            {
                string heading = file.Name[..file.Name.IndexOf(".cs")];
                docGenerator.WriteNewParagraph(heading);
                
                IEnumerable<string> fileLines = FileReaderService.GetValidFileLines(file.FullName);
                string routeLine = fileLines.First(x => x.Contains("Route"));
                
                foreach(string line in fileLines)
                {
                    docGenerator.WriteNewLine(line);
                }             
            }

            docGenerator.Save();
            return Task.CompletedTask;
        }
    }
}
