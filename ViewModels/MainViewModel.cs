using APIDocGenerator.Services;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Windows.Data.Json;

namespace APIDocGenerator.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger<MainViewModel> _logger;
        private readonly IFolderPicker _folderPicker;
        private readonly IFilePicker _filePicker;

        [ObservableProperty]
        private string _selectedSource = string.Empty;
        [ObservableProperty]
        private string _selectedDestination = string.Empty;
        [ObservableProperty]
        private string _fileName = string.Empty;
        [ObservableProperty]
        private bool _jsonFileSelectionIsVisible = true;
        [ObservableProperty]
        private bool _folderSelectionIsVisible = false;
        [ObservableProperty]
        private bool _useJsonFileIsOn = true;
        [ObservableProperty]
        private bool _useControllersIsOn = false;

        public MainViewModel(ILogger<MainViewModel> logger, IFolderPicker folderPicker, IFilePicker filePicker)
        {
            _logger = logger;
            _folderPicker = folderPicker;
            _filePicker = filePicker;
        }

        /// <summary>
        /// Gets user selected source folder for controllers.
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        [RelayCommand]
        public async Task SelectSourceFolder(CancellationToken token)
        {
#pragma warning disable CA1416 // Validate platform compatibility
            FolderPickerResult result = await _folderPicker.PickAsync(token);
#pragma warning restore CA1416 // Validate platform compatibility
            try
            {
                result.EnsureSuccess();

                SelectedSource = result.Folder.Path;
            }
            catch (Exception)
            {
                // just eat it
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        [RelayCommand]
        public async Task SelectJsonSourceFile(CancellationToken token)
        {
            IDictionary<DevicePlatform, IEnumerable<string>> fileTypes = new Dictionary<DevicePlatform, IEnumerable<string>> { { DevicePlatform.WinUI, [".json"] } };
            FileResult? result = await _filePicker.PickAsync(new PickOptions { FileTypes = new FilePickerFileType(fileTypes) });

            if(result != null)
            {
                SelectedSource = result.FullPath;
            }
        }

        /// <summary>
        /// Gets user selected destination folder for the output file.
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        [RelayCommand]
        public async Task SelectDestinationFolder(CancellationToken token)
        {
#pragma warning disable CA1416 // Validate platform compatibility
            FolderPickerResult result = await _folderPicker.PickAsync(token);
#pragma warning restore CA1416 // Validate platform compatibility
            try
            {
                result.EnsureSuccess();

                SelectedDestination = result.Folder.Path;
            }
            catch (Exception)
            {
                // just eat it
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void SwapSourceSelectionOption()
        {
            JsonFileSelectionIsVisible = UseJsonFileIsOn;
            FolderSelectionIsVisible = !UseJsonFileIsOn;
            SelectedSource = string.Empty;
        }

        /// <summary>
        /// Generates and outputs the finalized document.
        /// </summary>
        /// <returns></returns>
        public async Task GenerateDocument()
        {
            if (string.IsNullOrWhiteSpace(FileName))
            {
                throw new Exception("File name must be set");
            } 
            else
            {
                if (FileName.Contains(".docx"))
                {
                    FileName = FileName.Replace(".docx", "");
                }

                if (FileName.Contains(".doc"))
                {
                    FileName = FileName.Replace(".doc", "");
                }
            }

            if (string.IsNullOrWhiteSpace(SelectedDestination))
            {
                throw new Exception("Destination folder must be set");
            }

            if (string.IsNullOrWhiteSpace(SelectedSource))
            {
                throw new Exception("Source must be set");
            }

            DocumentGenerator docGenerator = new DocumentGenerator(SelectedDestination, FileName);

            if (UseControllersIsOn) 
            {
                IEnumerable<FileInfo> sourceFiles = FileReaderService.GetFiles(SelectedSource);
                await docGenerator.GenerateFromControllerFiles(sourceFiles);
            }
            else
            {
                string jsonContent = await File.ReadAllTextAsync(SelectedSource);
                JObject jsonParse = JObject.Parse(jsonContent);
                await docGenerator.GenerateFromJson(jsonParse);
            }
        }
    }
}
