using APIDocGenerator.Services;
using CommunityToolkit.Maui.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;

namespace APIDocGenerator.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger<MainViewModel> _logger;
        private readonly IFolderPicker _folderPicker;
        private readonly TextParserService _parserService;

        [ObservableProperty]
        private string _selectedSource = string.Empty;
        [ObservableProperty]
        private string _selectedDestination = string.Empty;
        [ObservableProperty]
        private string _fileName = string.Empty;

        public MainViewModel(ILogger<MainViewModel> logger, IFolderPicker folderPicker, TextParserService parserService)
        {
            _logger = logger;
            _folderPicker = folderPicker;
            _parserService = parserService;
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
        /// Generates and outputs the finalized document.
        /// </summary>
        /// <returns></returns>
        public Task GenerateDocument()
        {
            if (string.IsNullOrWhiteSpace(FileName))
            {
                throw new Exception("File name must be set");
            } else
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
                throw new Exception("Controller source folder must be set");
            }

            IEnumerable<FileInfo> sourceFiles = FileReaderService.GetFiles(SelectedSource);
            DocumentGenerator docGenerator = new DocumentGenerator(SelectedDestination, FileName);

            docGenerator.AddTitleLine(FileName);

            foreach(FileInfo file in sourceFiles)
            {
                string controllerName = file.Name[..file.Name.IndexOf(".cs")];
                string controllerRouting = controllerName.Replace("Controller", "").ToLower();
                IEnumerable<string> fileLines = FileReaderService.GetValidFileLines(file.FullName);


                string versionString = _parserService.GetVersionInfo(fileLines);
                string? routeLine = fileLines.FirstOrDefault(x => x.Contains("Route("));

                // if the controller has no routing info, probably a base or abstract for inheritance
                if(routeLine != default)
                {
                    string parsedControllerRoute = routeLine.Split('"')[1]
                        .Replace("[controller]", controllerRouting)
                        .Replace("v{v:apiVersion}", $"{{{versionString}}}");

                    if (!parsedControllerRoute.Contains("api"))
                    {
                        parsedControllerRoute = $"api/{parsedControllerRoute}";
                    }

                    string paragraphHeader = $"{controllerName} {versionString}";
                    docGenerator.WriteNewParagraph(paragraphHeader);

                    List<string> endpointLines = _parserService.GetLinesAtFirstEndpoint(fileLines).ToList();

                    for (int i = 0; i < endpointLines.Count; i++)
                    {
                        string copy = endpointLines[i];
                        if (copy.StartsWith("[Http"))
                        {
                            var (type, endpoint) = _parserService.GetEndPointRouting(copy);
                            string outPut = $"{parsedControllerRoute}{endpoint}";
                            docGenerator.WriteRouteLine(type, outPut);
                        }

                        if (copy.StartsWith("///"))
                        {
                            var (lastIdx, output) = _parserService.GetParsedXMLString(endpointLines, i);
                            i = lastIdx; // skip past other lines in same comment section
                            docGenerator.WriteCommentLine(output);
                        }
                    }
                }       
            }

            docGenerator.Save();
            return Task.CompletedTask;
        }
    }
}
