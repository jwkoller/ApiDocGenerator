using APIDocGenerator.Models.JsonParse;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Schema;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Tab = DocumentFormat.OpenXml.Office.CustomUI.Tab;

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
        /// 
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
                    Paragraph get = CreateNewRequestTypeSection("GET", routeDetails.Get);
                    routeHeader.AppendChild(get);

                    controllerName = routeDetails.Get.Tags.First();           
                }

                if (routeDetails.Put != null) 
                {
                    Paragraph put = CreateNewRequestTypeSection("PUT", routeDetails.Put);
                    routeHeader.AppendChild(put);

                    controllerName = routeDetails.Put.Tags.First();
                }

                if (routeDetails.Post != null)
                {
                    Paragraph post = CreateNewRequestTypeSection("POST", routeDetails.Post);
                    routeHeader.AppendChild(post);

                    controllerName = routeDetails.Post.Tags.First();
                }

                if (routeDetails.Delete != null)
                {
                    Paragraph delete = CreateNewRequestTypeSection("DELETE", routeDetails.Delete);
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
            Run run = new Run();
            RunProperties props = new RunProperties();
            props.Bold = new Bold();
            props.FontSize = new FontSize() { Val = HEADING_FONT_SIZE };

            run.Append(props);
            run.AppendChild(new Text() { Text = path, Space = SpaceProcessingModeValues.Preserve });
            run.AppendChild(new Break());
            paragraph.AppendChild(run);

            return paragraph;
        }

        /// <summary>
        /// Creates a new HTTP Request type (GET, POST, etc) sub-section that contains details about the request including parameters, 
        /// responses, and request body content.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="details"></param>
        /// <returns></returns>
        private Paragraph CreateNewRequestTypeSection(string type, RequestType details)
        {
            Paragraph paragraph = new Paragraph();
            Run run = paragraph.AppendChild(new Run());
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
                Run next = paragraph.AppendChild(new Run());
                RunProperties nextProps = new RunProperties();
                nextProps.FontSize = new FontSize { Val = TEXT_FONT_SIZE };
                next.Append(nextProps);
                next.AppendChild(new Text() { Text = $"{details.Summary}", Space = SpaceProcessingModeValues.Preserve });
                next.AppendChild(new Break());
            }
            else
            {
                run.AppendChild(new Break());
            }

            if (details.Parameters != null)
            {
                paragraph.AppendChild(CreateNewParameterSection(details.Parameters));
            }

            if (details.RequestBody != null) 
            {
                paragraph.AppendChild(CreateNewRequestBodySection(details.RequestBody));
            }



            return paragraph;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static Paragraph CreateNewParameterSection(IEnumerable<Parameter> parameters)
        {
            Paragraph paragraph = new Paragraph();
            Run paramSection = paragraph.AppendChild(new Run());
            RunProperties paramProperties = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = TEXT_FONT_SIZE }
            };
            paramSection.Append(paramProperties);
            paramSection.AppendChild(new Text { Text = "PARAMS", Space = SpaceProcessingModeValues.Preserve });
            paramSection.AppendChild(new Break());

            foreach (Parameter param in parameters)
            {
                paragraph.AppendChild(CreateNewParameter(param));
            }
            return paragraph;
        }

        /// <summary>
        /// Creates a single new formatted parameter sub-section. 
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        private static Paragraph CreateNewParameter(Parameter param)
        {          
            Paragraph paragraph = new Paragraph();

            Run name = paragraph.AppendChild(new Run());
            RunProperties nameProps = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = JSON_FONT_SIZE }
            };
            name.Append(nameProps);
            name.AppendChild(new TabChar());
            name.AppendChild(new Text { Text = $"{param.Name} : ", Space = SpaceProcessingModeValues.Preserve });

            Run valueType = paragraph.AppendChild(new Run());
            RunProperties typeProps = new RunProperties
            {
                FontSize = new FontSize { Val = JSON_FONT_SIZE }
            };
            valueType.Append(typeProps);
            string typeText = string.IsNullOrEmpty(param.Schema.Format) ? param.Schema.Type : param.Schema.Format;
            valueType.AppendChild(new Text { Text = typeText, Space = SpaceProcessingModeValues.Preserve });
            valueType.AppendChild(new Break());

            if (!string.IsNullOrEmpty(param.Description))
            {
                Run summary = paragraph.AppendChild(new Run());
                RunProperties summaryProps = new RunProperties
                {
                    FontSize = new FontSize { Val = JSON_FONT_SIZE }
                };
                summary.Append(summaryProps);
                summary.AppendChild(new TabChar());
                summary.AppendChild(new TabChar());
                summary.AppendChild(new Text { Text = param.Description, Space = SpaceProcessingModeValues.Preserve });
                summary.AppendChild(new Break());
            }
            
            Run locationKey = paragraph.AppendChild(new Run());
            RunProperties keyProps = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = JSON_FONT_SIZE }
            };
            locationKey.Append(keyProps);
            locationKey.AppendChild(new TabChar());
            locationKey.AppendChild(new TabChar());
            locationKey.AppendChild(new Text { Text = "In: ", Space = SpaceProcessingModeValues.Preserve });

            Run locationValue = paragraph.AppendChild(new Run());
            RunProperties locValueProps = new RunProperties
            {
                FontSize = new FontSize { Val = JSON_FONT_SIZE }
            };
            locationValue.Append(locValueProps);
            locationValue.AppendChild(new Text { Text = param.In, Space = SpaceProcessingModeValues.Preserve });

            if (param.Required)
            {
                locationValue.AppendChild(new Break());
                Run required = paragraph.AppendChild(new Run());
                RunProperties reqProps = new RunProperties
                {
                    Bold = new Bold(),
                    FontSize = new FontSize { Val = JSON_FONT_SIZE }
                };
                required.Append(reqProps);
                required.AppendChild(new TabChar());
                required.AppendChild(new TabChar());
                required.AppendChild(new Text { Text = "Required", Space = SpaceProcessingModeValues.Preserve });
            }

            return paragraph;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        private Paragraph CreateNewRequestBodySection(RequestBody body)
        {
            Paragraph paragraph = new Paragraph();
            Run reqBody = paragraph.AppendChild(new Run());
            RunProperties reqBodyProps = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = TEXT_FONT_SIZE }
            };
            reqBody.Append(reqBodyProps);
            reqBody.AppendChild(new Text { Text = "REQUEST BODY", Space = SpaceProcessingModeValues.Preserve });
            reqBody.AppendChild(new Break());

            if (!string.IsNullOrEmpty(body.Description))
            {
                Run content = paragraph.AppendChild(new Run());
                RunProperties contentProps = new RunProperties
                {
                    FontSize = new FontSize { Val = JSON_FONT_SIZE }
                };
                content.Append(contentProps);
                content.AppendChild(new TabChar());
                content.AppendChild(new Text { Text = body.Description, Space = SpaceProcessingModeValues.Preserve });
                content.AppendChild(new Break());
            }
            
            if (body.Content.TryGetValue("application/json", out Content? value))
            {
                Run appJson = paragraph.AppendChild(new Run());
                RunProperties appJsonProps = new RunProperties
                {
                    FontSize = new FontSize { Val = JSON_FONT_SIZE },
                    Bold = new Bold()
                };
                appJson.Append(appJsonProps);
                appJson.AppendChild(new TabChar());
                appJson.AppendChild(new Text { Text = "application/json", Space = SpaceProcessingModeValues.Preserve });
                appJson.AppendChild(new Break());

                Schema jsonSchema = value.Schema;
                if (!string.IsNullOrEmpty(jsonSchema.Ref))
                {
                    jsonSchema = GetSchemaComponent(jsonSchema.Ref);
                }

                Run content = paragraph.AppendChild(new Run());
                RunProperties contentProps = new RunProperties
                {
                    FontSize = new FontSize { Val = JSON_FONT_SIZE }
                };
                content.AppendChild(new TabChar());
                content.AppendChild(new TabChar());
                content.AppendChild(GetFormattedSchema(jsonSchema));


                //foreach(KeyValuePair<string,Schema> objProps in jsonSchema.Properties)
                //{
                //    Run name = paragraph.AppendChild(new Run());
                //    RunProperties nameProps = new RunProperties
                //    {
                //        Bold = new Bold(),
                //        FontSize = new FontSize { Val = JSON_FONT_SIZE }
                //    };
                //    name.Append(nameProps);
                //    name.AppendChild(new TabChar());
                //    name.AppendChild(new TabChar());
                //    name.AppendChild(new Text { Text = $"{objProps.Key}: ", Space = SpaceProcessingModeValues.Preserve });

                //    Run val = paragraph.AppendChild(new Run());
                //    RunProperties valProps = new RunProperties
                //    {
                //        FontSize = new FontSize { Val = JSON_FONT_SIZE }
                //    };
                //    val.Append(valProps);
                //    string typeText = string.IsNullOrEmpty(objProps.Value.Format) ? objProps.Value.Type : objProps.Value.Format;
                //    val.AppendChild(new Text { Text = typeText, Space = SpaceProcessingModeValues.Preserve });


                //}
            }


            return paragraph;
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

        private Paragraph GetFormattedSchema(Schema schemaToFormat)
        {
            Schema schema = schemaToFormat;
            if (!string.IsNullOrEmpty(schema.Ref))
            {
                schema = GetSchemaComponent(schema.Ref);
            }

            Paragraph paragraph = new Paragraph();

            if(schema.Type == "object")
            {
                foreach (KeyValuePair<string, Schema> objProps in schema.Properties)
                {
                    string propName = objProps.Key;
                    Schema propSchema = objProps.Value;

                    if (!string.IsNullOrEmpty(propSchema.Ref))
                    {
                        propSchema = GetSchemaComponent(propSchema.Ref);
                        paragraph.AppendChild(GetFormattedSchema(propSchema));

                    } else
                    {
                        Run name = paragraph.AppendChild(new Run());
                        RunProperties nameProps = new RunProperties
                        {
                            Bold = new Bold(),
                            FontSize = new FontSize { Val = JSON_FONT_SIZE }
                        };
                        name.Append(nameProps);
                        name.AppendChild(new Text { Text = $"{propName}: ", Space = SpaceProcessingModeValues.Preserve });

                        Run val = paragraph.AppendChild(new Run());
                        RunProperties valProps = new RunProperties
                        {
                            FontSize = new FontSize { Val = JSON_FONT_SIZE }
                        };
                        val.Append(valProps);
                        string typeText = string.IsNullOrEmpty(propSchema.Format) ? propSchema.Type : propSchema.Format;
                        val.AppendChild(new Text { Text = typeText, Space = SpaceProcessingModeValues.Preserve });
                        val.AppendChild(new Break());

                        Run desc = paragraph.AppendChild(new Run());
                        RunProperties descProps = new RunProperties
                        {
                            FontSize = new FontSize { Val = JSON_FONT_SIZE }
                        };
                        desc.Append(descProps);
                        desc.AppendChild(new TabChar());
                        desc.AppendChild(new Text { Text = propSchema.Description, Space = SpaceProcessingModeValues.Preserve });
                        desc.AppendChild(new Break());

                        if (propSchema.Items != null)
                        {
                            Run items = paragraph.AppendChild(new Run());
                            RunProperties itemsProp = new RunProperties
                            {
                                FontSize = new FontSize { Val = JSON_FONT_SIZE }
                            };
                            items.Append(itemsProp);
                            items.AppendChild(new TabChar());
                            items.AppendChild(GetFormattedSchema(propSchema.Items));
                        }
                    }
                }
            }

            //if(schema.Type == "array")
            //{
            //    Run items = paragraph.AppendChild(new Run());
            //    RunProperties itemsProps = new RunProperties
            //    {
            //        FontSize = new FontSize { Val = JSON_FONT_SIZE },
            //        Bold = new Bold()
            //    };
            //    items.Append(itemsProps);
            //    items.AppendChild(new TabChar());
            //    items.AppendChild(new TabChar());
            //    items.AppendChild(new Text { Text = "items: ", Space = SpaceProcessingModeValues.Preserve });

            //    Run itemsType = paragraph.AppendChild(new Run());
            //    RunProperties itemsTypeProps = new RunProperties
            //    {
            //        FontSize = new FontSize { Val = JSON_FONT_SIZE },
            //    };
            //    itemsType.Append(itemsTypeProps);
            //    itemsType.AppendChild(new Text { Text = schema.Items.Type, Space = SpaceProcessingModeValues.Preserve });
            //    itemsType.AppendChild(new Break());


            //}

            return paragraph;
        }

        /// <summary>
        /// 
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
