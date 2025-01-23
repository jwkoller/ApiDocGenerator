using APIDocGenerator.Models.JsonParse;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Format = APIDocGenerator.Services.FormattingService;

namespace APIDocGenerator.Services
{
    public class DocumentGenerator
    {
        private const string TITLE_FONT_SIZE = "40";
        private const string HEADING_FONT_SIZE = "32";
        private const string TEXT_FONT_SIZE = "24";
        private const string JSON_FONT_SIZE = "20";

        private string _destinationFolder;
        private Components _jsonComponents;

        public string DocumentName { get; private set; }
        public WordprocessingDocument Document { get; private set; }
        public MainDocumentPart MainPart {  get; private set; }
        public Body Body { get; private set; }
       
        public DocumentGenerator(string destination, string fileName)
        {
            _destinationFolder = destination;
            DocumentName = fileName;
        }

        /// <summary>
        /// 
        /// </summary>
        private void CreateBlankDocument()
        {
            Document = WordprocessingDocument.Create($"{_destinationFolder}{Path.DirectorySeparatorChar}{DocumentName}.docx", WordprocessingDocumentType.Document);
            MainPart = Document.AddMainDocumentPart();
            MainPart.Document = new Document();
            Body = MainPart.Document.AppendChild(new Body());
        }

        /// <summary>
        /// Creates a new paragraph with a bolded heading.
        /// </summary>
        /// <param name="heading"></param>
        public void WriteNewParagraph(string heading)
        {
            Paragraph paragraph = Body.AppendChild(new Paragraph());
            if(!paragraph.Elements<ParagraphProperties>().Any())
            {
                paragraph.PrependChild(new ParagraphProperties());
            }

            Run run = new Run();
            RunProperties props = new RunProperties();
            props.Bold = new Bold();
            props.FontSize = new FontSize() { Val = HEADING_FONT_SIZE};
            
            run.Append(props);
            run.AppendChild(new Break());
            run.AppendChild(new Text() { Text = heading, Space = SpaceProcessingModeValues.Preserve });
            run.AppendChild(new Break());
            paragraph.AppendChild(run);
        }

        /// <summary>
        /// Writes a new comment section to the last paragraph element.
        /// </summary>
        /// <param name="text"></param>
        public void WriteCommentLine(string text)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();         
            Run run = last.AppendChild(new Run());
            RunProperties props = new RunProperties();
            props.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };

            run.AppendChild(props);
            run.AppendChild(new Break());
            run.AppendChild(new Text() { Text = text, Space = SpaceProcessingModeValues.Preserve });
            run.AppendChild(new Break());
        }

        /// <summary>
        /// Constructs a new comment section under the last paragraph with from the list of
        /// key value pairs. The key is bolded for each line.
        /// </summary>
        /// <param name="commentLines"></param>
        public void WriteCommentLine(List<KeyValuePair<string, string>> commentLines)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();
            Run newLine = last.AppendChild(new Run());
            newLine.AppendChild(new Break());

            foreach(KeyValuePair<string, string> line in commentLines)
            {
                Run run = last.AppendChild(new Run());
                RunProperties props = new RunProperties();
                props.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };
                props.Bold = new Bold();
                run.AppendChild(props);
                run.AppendChild(new Text() { Text = line.Key, Space = SpaceProcessingModeValues.Preserve });

                Run next = last.AppendChild(new Run());
                RunProperties nextProps = new RunProperties();
                props.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };
                next.AppendChild(nextProps);
                next.AppendChild(new Text() { Text = line.Value, Space = SpaceProcessingModeValues.Preserve });
                next.AppendChild(new Break());
            }
        }

        /// <summary>
        /// Writes a new formatted route to the last paragraph element.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="text"></param>
        public void WriteRouteLine(string type, string text)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();
            Run run = last.AppendChild(new Run());
            RunProperties props = new RunProperties();
            props.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };
            props.Bold = new Bold();

            switch (type)
            {
                case "HttpGet":
                    props.Color = new Color() { Val = "15a612" };
                    break;
                case "HttpPost":
                    props.Color = new Color() { Val = "467be3" };
                    break;
                case "HttpPut":
                    props.Color = new Color() { Val = "e0da1d" };
                    break;
                case "HttpDelete":
                    props.Color = new Color() { Val = "e03614" };
                    break;
            }

            run.Append(props);          
            run.AppendChild(new Text() { Text = $"{type}: ", Space = SpaceProcessingModeValues.Preserve });

            Run next = last.AppendChild(new Run());
            RunProperties nextProps = new RunProperties();
            nextProps.Bold = new Bold();
            nextProps.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };

            next.Append(nextProps);
            next.AppendChild(new Text() { Text = $"/{text}", Space = SpaceProcessingModeValues.Preserve });
            next.AppendChild(new Break());
        }

        /// <summary>
        /// Adds a 20pt font-size, bolded, centered line of text.
        /// </summary>
        /// <param name="headerText"></param>
        public void AddTitleLine(string headerText)
        {
            Paragraph paragraph = Body.AppendChild(new Paragraph());
            if (!paragraph.Elements<ParagraphProperties>().Any())
            {
                paragraph.PrependChild(new ParagraphProperties());
            }

            Justification centered = new Justification() { Val = JustificationValues.Center };
            paragraph.ParagraphProperties?.Append(centered);

            Run run = new Run();
            RunProperties props = new RunProperties();
            props.Bold = new Bold();
            props.FontSize = new FontSize() { Val = TITLE_FONT_SIZE };

            run.Append(props);
            run.AppendChild(new Break());
            run.AppendChild(new Text() { Text = headerText });
            run.AppendChild(new Break());
            paragraph.AppendChild(run);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceFiles"></param>
        /// <returns></returns>
        public Task GenerateFromControllerFiles(IEnumerable<FileInfo> sourceFiles)
        {
            CreateBlankDocument();
            AddTitleLine(DocumentName);

            foreach (FileInfo file in sourceFiles)
            {
                string controllerName = file.Name[..file.Name.IndexOf(".cs")];
                string controllerRouting = controllerName.Replace("Controller", "").ToLower();
                IEnumerable<string> fileLines = FileReaderService.GetValidFileLines(file.FullName);


                string versionString = TextParserService.GetVersionInfo(fileLines);
                string? routeLine = fileLines.FirstOrDefault(x => x.Contains("Route("));

                // if the controller has no routing info, probably a base or abstract for inheritance
                if (routeLine != default)
                {
                    string parsedControllerRoute = routeLine.Split('"')[1]
                        .Replace("[controller]", controllerRouting)
                        .Replace("v{v:apiVersion}", $"{{{versionString}}}");

                    if (!parsedControllerRoute.Contains("api"))
                    {
                        parsedControllerRoute = $"api/{parsedControllerRoute}";
                    }

                    string paragraphHeader = $"{controllerName} {versionString}";
                    WriteNewParagraph(paragraphHeader);

                    List<string> endpointLines = TextParserService.GetLinesAtFirstEndpoint(fileLines).ToList();

                    for (int i = 0; i < endpointLines.Count; i++)
                    {
                        string copy = endpointLines[i];
                        if (copy.StartsWith("[Http"))
                        {
                            var (type, endpoint) = TextParserService.GetEndPointRouting(copy);
                            string outPut = $"{parsedControllerRoute}{endpoint}";
                            WriteRouteLine(type, outPut);
                        }

                        if (copy.StartsWith("///"))
                        {
                            var (lastIdx, output) = TextParserService.GetParsedXMLString(endpointLines, i);
                            i = lastIdx; // skip past other lines in same comment section
                            WriteCommentLine(output);
                        }
                    }
                }
            }

            Save();
            return Task.CompletedTask;
        }

        /// <summary>
        /// Generates a new document from a Swagger generated JSON string.
        /// </summary>
        /// <returns></returns>
        public Task GenerateFromJson(string json)
        {
            var settings = new JsonSerializerSettings { MetadataPropertyHandling = MetadataPropertyHandling.Ignore };
            RootApiJson? apiRoot = JsonConvert.DeserializeObject<RootApiJson>(json, settings);

            if (apiRoot == null)
            {
                throw new Exception("Error encountered parsing the JSON file.");
            }
            _jsonComponents = apiRoot.Components;

            CreateBlankDocument();
            string version = !string.IsNullOrEmpty(apiRoot.Info?.Version) ? $" v{apiRoot.Info.Version}" : string.Empty;
            AddTitleLine($"{DocumentName}{version}");
            Dictionary<string, List<Paragraph>> controllerSections = [];

            foreach (KeyValuePair<string, Route> path in apiRoot.Paths) 
            { 
                string uriPath = path.Key;
                string controllerName = string.Empty;

                Paragraph routeHeader = CreateNewRouteSection(uriPath);

                Route routeDetails = path.Value;

                if(routeDetails.Get != null)
                {
                    Run get = CreateNewRequestTypeSection("GET", routeDetails.Get);
                    routeHeader.AppendChild(get);

                    controllerName = routeDetails.Get.Tags.First();           
                }

                if (routeDetails.Put != null) 
                {
                    Run put = CreateNewRequestTypeSection("PUT", routeDetails.Put);
                    routeHeader.AppendChild(put);

                    controllerName = routeDetails.Put.Tags.First();
                }

                if (routeDetails.Post != null)
                {
                    Run post = CreateNewRequestTypeSection("POST", routeDetails.Post);
                    routeHeader.AppendChild(post);

                    controllerName = routeDetails.Post.Tags.First();
                }

                if (routeDetails.Delete != null)
                {
                    Run delete = CreateNewRequestTypeSection("DELETE", routeDetails.Delete);
                    routeHeader.AppendChild(delete);

                    controllerName = routeDetails.Delete.Tags.First();
                }

                if (!controllerSections.TryGetValue(controllerName, out List<Paragraph>? value)) 
                {
                    value = new List<Paragraph>();
                    controllerSections.Add(controllerName, value);
                }

                value.Add(routeHeader);
            }

            CompileDocument(controllerSections);

            Save();
            return Task.CompletedTask;
        }

        /// <summary>
        /// Creates a new endpoint url section to contains the various HTTP request types accepted.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static Paragraph CreateNewRouteSection(string path)
        {
            Paragraph paragraph = new Paragraph();
            Run run =paragraph.AppendChild(new Run());
            run.AppendChild(Format.CreateBoldTextLine(path, HEADING_FONT_SIZE));
            run.AppendChild(new CarriageReturn());

            return paragraph;
        }

        /// <summary>
        /// Creates a new HTTP Request type (GET, POST, etc) sub-section that contains details about the request including parameters, 
        /// responses, and request body content.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="details"></param>
        /// <returns></returns>
        private Run CreateNewRequestTypeSection(string type, RequestType details)
        {
            Run container = new Run();
            Run run = container.AppendChild(new Run());
            RunProperties props = new RunProperties();
            props.FontSize = new FontSize() { Val = TEXT_FONT_SIZE };
            props.Bold = new Bold();

            switch (type)
            {
                case "GET":
                    props.Color = new Color() { Val = "15a612" };
                    break;
                case "POST":
                    props.Color = new Color() { Val = "467be3" };
                    break;
                case "PUT":
                    props.Color = new Color() { Val = "e0da1d" };
                    break;
                case "DELETE":
                    props.Color = new Color() { Val = "e03614" };
                    break;
            }

            run.Append(props);
            run.AppendChild(new Text() { Text = $"{type}: ", Space = SpaceProcessingModeValues.Preserve });
            if (!string.IsNullOrEmpty(details.Summary))
            {
                Run next = container.AppendChild(new Run());
                next.AppendChild(Format.CreateTextLine(details.Summary, TEXT_FONT_SIZE));
                next.AppendChild(new CarriageReturn());
            }
            else
            {
                run.AppendChild(new CarriageReturn());
            }

            if (details.Parameters != null)
            {
                container.AppendChild(CreateNewParameterSection(details.Parameters));
            }

            if (details.RequestBody != null) 
            {
                container.AppendChild(CreateNewRequestBodySection(details.RequestBody));
            }



            return container;
        }

        /// <summary>
        /// Create a new Parameter section for a HTTP request type.
        /// </summary>
        /// <returns></returns>
        private static Run CreateNewParameterSection(IEnumerable<Parameter> parameters)
        {
            Run container = new Run();
            Run paramSection = container.AppendChild(new Run());
            paramSection.AppendChild(Format.CreateBoldTextLine("Parameters", TEXT_FONT_SIZE));
            paramSection.AppendChild(new CarriageReturn());

            foreach (Parameter param in parameters)
            {
                container.AppendChild(CreateNewParameter(param));
            }
            return container;
        }

        /// <summary>
        /// Creates a single new formatted parameter sub-section. 
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        private static Paragraph CreateNewParameter(Parameter param)
        {          
            Paragraph paragraph = new Paragraph();
            Run container = paragraph.AppendChild(new Run());
            container.AppendChild(new TabChar());

            string typeText = string.IsNullOrEmpty(param.Schema.Format) ? param.Schema.Type : param.Schema.Format;
            container.AppendChild(Format.CreateLabelValuePair($"{param.Name} : ", typeText, JSON_FONT_SIZE));
            container.AppendChild(new Break());

            if (!string.IsNullOrEmpty(param.Description))
            {
                Run summary = container.AppendChild(new Run());
                summary.AppendChild(new TabChar());
                summary.AppendChild(new TabChar());
                summary.AppendChild(Format.CreateTextLine(param.Description, JSON_FONT_SIZE));
                summary.AppendChild(new Break());
            }
            
            Run locationRun = container.AppendChild(new Run());
            locationRun.AppendChild(new TabChar());
            locationRun.AppendChild(new TabChar());
            locationRun.AppendChild(Format.CreateLabelValuePair("In: ", param.In, JSON_FONT_SIZE));

            if (param.Required)
            {
                locationRun.AppendChild(new Break());
                Run required = container.AppendChild(new Run());
                required.AppendChild(new TabChar());
                required.AppendChild(new TabChar());
                required.AppendChild(Format.CreateBoldTextLine("Required", JSON_FONT_SIZE));
            }

            return paragraph;
        }

        /// <summary>
        /// Creates a new HTTP POST request body section, including schema obj formatting.
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        private Run CreateNewRequestBodySection(RequestBody body)
        {
            Run container = new Run();
            Run reqBody = container.AppendChild(new Run());
            reqBody.AppendChild(Format.CreateBoldTextLine("Request Body", TEXT_FONT_SIZE));
            reqBody.AppendChild(new CarriageReturn());

            if (!string.IsNullOrEmpty(body.Description))
            {
                Run description = container.AppendChild(new Run());
                description.AppendChild(new TabChar());
                description.AppendChild(Format.CreateTextLine(body.Description, JSON_FONT_SIZE));
                description.AppendChild(new CarriageReturn());
            }
            
            if (body.Content.TryGetValue("application/json", out Content? appJsonContent))
            {
                Run appJson = container.AppendChild(new Run());
                appJson.AppendChild(new TabChar());
                appJson.AppendChild(Format.CreateBoldTextLine("application/json", JSON_FONT_SIZE));
                appJson.AppendChild(new CarriageReturn());
                appJson.AppendChild(CreateFormattedSchema(appJsonContent.Schema, 2));
            }

            if(body.Content.TryGetValue("multipart/form-data", out Content? formData))
            {
                Run formRun = container.AppendChild(new Run());
                formRun.AppendChild(new TabChar());
                formRun.AppendChild(Format.CreateBoldTextLine("multipart/form-data", JSON_FONT_SIZE));
                formRun.AppendChild(new CarriageReturn());
                formRun.AppendChild(CreateFormattedSchema(formData.Schema, 2));
            }

            return container;
        }

        /// <summary>
        /// Creates a heading for a controller that will contain the various available endpoints.
        /// </summary>
        /// <param name="controllerName"></param>
        /// <returns></returns>
        private static Paragraph CreateNewControllerHeading(string controllerName)
        {
            Paragraph paragraph = new Paragraph();
            if (!paragraph.Elements<ParagraphProperties>().Any())
            {
                paragraph.PrependChild(new ParagraphProperties());
            }

            Justification centered = new Justification() { Val = JustificationValues.Center };
            paragraph.ParagraphProperties?.Append(centered);

            Run run = new Run();
            RunProperties props = new RunProperties();
            props.Bold = new Bold();
            props.FontSize = new FontSize() { Val = TITLE_FONT_SIZE };

            string headingText = $"{controllerName} Endpoints";
            run.Append(props);
            run.AppendChild(new Text() { Text = headingText });
            run.AppendChild(new Break());
            paragraph.AppendChild(run);

            return paragraph;
        }

        /// <summary>
        /// Creates a formatted section for a given schema, recursively drilling down to sub-properties. Number of tab indentations
        /// is tracked so I don't have to mess with a bulleted list cause it's a giant pita in OpenXMl.
        /// </summary>
        /// <param name="schemaToFormat"></param>
        /// <param name="numTabs"></param>
        /// <returns></returns>
        private Run CreateFormattedSchema(Schema schemaToFormat, int numTabs = 0)
        {
            Schema schema = schemaToFormat;
            if (!string.IsNullOrEmpty(schema.Ref))
            {
                schema = GetSchemaComponent(schema.Ref);
            }

            Run container = new Run();

            // pretty much always going to be true, since even arrays are wrapped in brackets
            if (schema.Type == "object")
            {
                foreach (KeyValuePair<string, Schema> objProps in schema.Properties)
                {
                    string propName = objProps.Key;
                    Schema propSchema = objProps.Value;

                    Run propertyRun = container.AppendChild(new Run());
                    for (int i = 0; i < numTabs; i++)
                    {
                        propertyRun.AppendChild(new TabChar());
                    }

                    if (!string.IsNullOrEmpty(propSchema.Ref))
                    {
                        propSchema = GetSchemaComponent(propSchema.Ref);
                        Run propParagraph = propertyRun.AppendChild(new Run());
                        propParagraph.AppendChild(CreateFormattedSchema(propSchema, numTabs+1));

                    } else
                    {
                        string typeText = string.IsNullOrEmpty(propSchema.Format) ? propSchema.Type : propSchema.Format;
                        propertyRun.AppendChild(Format.CreateLabelValuePair($"{propName}: ", typeText, JSON_FONT_SIZE));
                        propertyRun.AppendChild(new CarriageReturn());
                        if (!string.IsNullOrEmpty(propSchema.Description))
                        {
                            for (int i = 0; i < numTabs; i++)
                            {
                                propertyRun.AppendChild(new TabChar());
                            }
                            propertyRun.AppendChild(Format.CreateTextLine(propSchema.Description, JSON_FONT_SIZE));
                            propertyRun.AppendChild(new CarriageReturn());
                        }

                        if (propSchema.Items != null)
                        {
                            for (int i = 0; i < numTabs; i++)
                            {
                                propertyRun.AppendChild(new TabChar());
                            }
                            propertyRun.AppendChild(new TabChar());
                            string valueText = !string.IsNullOrEmpty(propSchema.Items.Ref) ? "object" : string.IsNullOrEmpty(propSchema.Format) ? propSchema.Type : propSchema.Format;
                            propertyRun.AppendChild(Format.CreateLabelValuePair("Items: ", valueText, JSON_FONT_SIZE));
                            propertyRun.AppendChild(new CarriageReturn());
                            
                            Run itemsRun = propertyRun.AppendChild(new Run());
                            itemsRun.AppendChild(CreateFormattedSchema(propSchema.Items, numTabs+2));
                        }
                    }
                }
            }

            return container;
        }

        /// <summary>
        /// Finds the requested json component from the reference string. This assumes a singular list of components.
        /// </summary>
        /// <param name="refString"></param>
        /// <returns></returns>
        private Schema GetSchemaComponent(string refString)
        {
            string name = refString.Split('/').Last();

            return _jsonComponents.Schemas[name];
        }

        /// <summary>
        /// Appends all the created paragraphs to the current document body in order they were added.
        /// </summary>
        /// <param name="paragraphs"></param>
        private void CompileDocument(Dictionary<string, List<Paragraph>> paragraphs)
        {
            foreach (KeyValuePair<string, List<Paragraph>> items in paragraphs) 
            {
                Paragraph heading = CreateNewControllerHeading(items.Key);
                Body.AppendChild(heading);

                foreach (Paragraph paragraph in items.Value) 
                {
                    Body.AppendChild(paragraph);
                }               
            }           
        }

        /// <summary>
        /// Disposes of the active document.
        /// </summary>
        public void Save()
        {
            Document.Dispose();
        }
    }
}
