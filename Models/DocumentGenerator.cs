using APIDocGenerator.Models.JsonParse;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
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
        private NumberingDefinitionsPart _numberingDefinitionsPart;
        private Dictionary<string, OpenXmlElement> _componentBulletLists;

        public string DocumentName { get; private set; }
        public WordprocessingDocument Document { get; private set; }
        public MainDocumentPart MainPart {  get; private set; }
        public Body Body { get; private set; }
       
        public DocumentGenerator(string destination, string fileName)
        {
            _destinationFolder = destination;
            _componentBulletLists = new Dictionary<string, OpenXmlElement>();
            DocumentName = fileName;
        }

        /// <summary>
        /// Creates the blank document
        /// </summary>
        private void CreateBlankDocument()
        {
            Document = WordprocessingDocument.Create($"{_destinationFolder}{Path.DirectorySeparatorChar}{DocumentName}.docx", WordprocessingDocumentType.Document);
            MainPart = Document.AddMainDocumentPart();
            MainPart.Document = new Document();
            Body = MainPart.Document.AppendChild(new Body());
        }

        /// <summary>
        /// Appends a new empty bulleted list to the main document and returns it's numbering id for use.
        /// </summary>
        /// <returns></returns>
        private int CreateNewBulletedList()
        {
            int abstractId = _numberingDefinitionsPart.Numbering.Elements<AbstractNum>().Count() + 1;
            Level abstractLevel = new Level(new NumberingFormat { Val = NumberFormatValues.None }, new LevelText { Val = "" }) { LevelIndex = 0 };
            AbstractNum abstractNum = new AbstractNum(abstractLevel) { AbstractNumberId = abstractId };

            if (abstractId == 1)
            {
                _numberingDefinitionsPart.Numbering.Append(abstractNum);
            }
            else
            {
                AbstractNum last = _numberingDefinitionsPart.Numbering.Elements<AbstractNum>().Last();
                _numberingDefinitionsPart.Numbering.InsertAfter(abstractNum, last);
            }

            int numberId = _numberingDefinitionsPart.Numbering.Elements<NumberingInstance>().Count() + 1;
            NumberingInstance numInstance = new NumberingInstance { NumberID = numberId };
            AbstractNumId abstractNumId = new AbstractNumId { Val = abstractId };
            numInstance.Append(abstractNumId);

            if (numberId == 1)
            {
                _numberingDefinitionsPart.Numbering.Append(numInstance);
            }
            else
            {
                NumberingInstance last = _numberingDefinitionsPart.Numbering.Elements<NumberingInstance>().Last();
                _numberingDefinitionsPart.Numbering.InsertAfter(numInstance, last);
            }

            return numberId;
        }

        /// <summary>
        /// Creates bulleted lists out of any component schemas and adds them to the internal dictionary for use.
        /// </summary>
        private void CreateComponentBulletParagraphs()
        {
            if(MainPart.NumberingDefinitionsPart == null)
            {
                _numberingDefinitionsPart = MainPart.AddNewPart<NumberingDefinitionsPart>("NumberDefintionsPart01");
                Numbering element = new Numbering();
                element.Save(_numberingDefinitionsPart);
            }

            foreach(KeyValuePair<string, Schema> componentSchema in _jsonComponents.Schemas)
            {
                string componentName = componentSchema.Key;
                Schema schema = componentSchema.Value;
                if (!string.IsNullOrEmpty(schema.Ref))
                {
                    schema = GetSchemaComponent(schema.Ref);
                }

                int listNumberId = CreateNewBulletedList();

                int startingIndent = 0;

                Run formattedSchema = CreateSchemaFormattedBulletList(startingIndent, schema, listNumberId);
                Paragraph componentContainer = new Paragraph();
                componentContainer.AppendChild(formattedSchema);

                _componentBulletLists.Add(componentName, componentContainer);
            }
        }

        /// <summary>
        /// Formats a single schema and it's properties into a bulleted list
        /// </summary>
        /// <param name="indent"></param>
        /// <param name="schemaToFormat"></param>
        /// <param name="bulletNumberId"></param>
        /// <returns></returns>
        private Run CreateSchemaFormattedBulletList(int indent, Schema schemaToFormat, int bulletNumberId)
        {
            Run container = new Run();

            Schema schema = schemaToFormat;
            if (!string.IsNullOrEmpty(schema.Ref))
            {
                schema = GetSchemaComponent(schema.Ref);
            }

            if(schema.Type == "object")
            {
                foreach (KeyValuePair<string, Schema> property in schema.Properties)
                {
                    Schema propSchema = property.Value;

                    if (!string.IsNullOrEmpty(propSchema.Ref))
                    {
                        propSchema = GetSchemaComponent(propSchema.Ref);
                        if(propSchema == schema)
                        {
                            Run propertyRun = Format.CreateLabelValuePair($"{property.Key}: ", "Same object as parent", JSON_FONT_SIZE);
                            Paragraph propParagraph = Format.CreateBulletedListItem(bulletNumberId, indent, propertyRun);
                            container.AppendChild(propParagraph);
                        }
                        else
                        {
                            Run itemsParagraph = CreateSchemaFormattedBulletList(indent + 1, propSchema, bulletNumberId);
                            container.AppendChild(itemsParagraph);
                        }
                    } 
                    else
                    {
                        Run propertyRun = Format.CreateLabelValuePair($"{property.Key}: ", propSchema.DisplayTypeText, JSON_FONT_SIZE);
                        Paragraph propParagraph = Format.CreateBulletedListItem(bulletNumberId, indent, propertyRun);
                        container.AppendChild(propParagraph);

                        if (propSchema.Items != null)
                        {
                            Run itemsParagraph = CreateSchemaFormattedBulletList(indent + 1, propSchema.Items, bulletNumberId);
                            container.AppendChild(itemsParagraph);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(propSchema.Description))
                            {
                                Run description = Format.CreateTextLine(propSchema.Description, JSON_FONT_SIZE);
                                Paragraph descParagraph = Format.CreateBulletedListItem(bulletNumberId, indent + 1, description);
                                container.AppendChild(descParagraph);
                            }
                        }
                    }
                }
            }
            else if (schema.Type == "array")
            {
                Run propertyRun = Format.CreateBoldTextLine($"Array", JSON_FONT_SIZE);
                Paragraph propParagraph = Format.CreateBulletedListItem(bulletNumberId, indent, propertyRun);
                container.AppendChild(propParagraph);

                if (!string.IsNullOrEmpty(schema.Description))
                {
                    Run description = Format.CreateTextLine(schema.Description, JSON_FONT_SIZE);
                    Paragraph descParagraph = Format.CreateBulletedListItem(bulletNumberId, indent + 1, description);
                    container.AppendChild(descParagraph);
                }

                Run itemsParagraph = CreateSchemaFormattedBulletList(indent + 1, schema.Items, bulletNumberId);
                container.AppendChild(itemsParagraph);
            }
            else
            {
                string label = string.IsNullOrEmpty(schema.Name) ? string.Empty : schema.Name;
                Run run = Format.CreateLabelValuePair(label, schema.DisplayTypeText, JSON_FONT_SIZE);
                Paragraph para = Format.CreateBulletedListItem(bulletNumberId, indent, run);
                container.AppendChild(para);

                if (!string.IsNullOrEmpty(schema.Description))
                {
                    Run description = Format.CreateTextLine(schema.Description, JSON_FONT_SIZE);
                    Paragraph descParagraph = Format.CreateBulletedListItem(bulletNumberId, indent + 1, description);
                    container.AppendChild(descParagraph);
                }
            }

            return container;
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
            // remove carriage returns and dead spacing
            string cleanedJson = Regex.Replace(json, @"((\\r\\n\s{2,})|(\\r\\n))", " ");
            // ignore meta properties like $ref so we can capture and use them
            var settings = new JsonSerializerSettings { MetadataPropertyHandling = MetadataPropertyHandling.Ignore };
            RootApiJson? apiRoot = JsonConvert.DeserializeObject<RootApiJson>(cleanedJson, settings);

            if (apiRoot == null)
            {
                throw new Exception("Error encountered parsing the JSON file.");
            }
            _jsonComponents = apiRoot.Components;

            CreateBlankDocument();
            CreateComponentBulletParagraphs();

            string version = !string.IsNullOrEmpty(apiRoot.Info?.Version) ? $" v{apiRoot.Info.Version}" : string.Empty;
            AddTitleLine($"{DocumentName}{version}");
            Dictionary<string, List<OpenXmlElement>> controllerSections = [];

            foreach (KeyValuePair<string, Route> path in apiRoot.Paths) 
            { 
                string uriPath = path.Key;
                string controllerName = string.Empty;

                List<OpenXmlElement> elements = new List<OpenXmlElement>();

                Paragraph routeHeader = CreateNewRouteSection(uriPath);
                elements.Add(routeHeader);

                Route routeDetails = path.Value;

                if(routeDetails.Get != null)
                {
                    elements.AddRange(CreateNewRequestTypeSection("GET", routeDetails.Get));

                    controllerName = routeDetails.Get.Tags.First();           
                }

                if (routeDetails.Put != null) 
                {
                    elements.AddRange(CreateNewRequestTypeSection("PUT", routeDetails.Put));

                    controllerName = routeDetails.Put.Tags.First();
                }

                if (routeDetails.Post != null)
                {
                    elements.AddRange(CreateNewRequestTypeSection("POST", routeDetails.Post));

                    controllerName = routeDetails.Post.Tags.First();
                }

                if (routeDetails.Delete != null)
                {
                    elements.AddRange(CreateNewRequestTypeSection("DELETE", routeDetails.Delete));

                    controllerName = routeDetails.Delete.Tags.First();
                }

                if (!controllerSections.TryGetValue(controllerName, out List<OpenXmlElement>? value)) 
                {
                    value = new List<OpenXmlElement>();
                    controllerSections.Add(controllerName, value);
                }

                value.AddRange(elements);
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
            ParagraphProperties properties = new ParagraphProperties(new SpacingBetweenLines { After = "0"});
            Paragraph paragraph = new Paragraph(properties);
            Run run = paragraph.AppendChild(new Run());
            run.AppendChild(new CarriageReturn());
            run.AppendChild(Format.CreateBoldTextLine(path, HEADING_FONT_SIZE));

            return paragraph;
        }

        /// <summary>
        /// Creates a new HTTP Request type (GET, POST, etc) sub-section that contains details about the request including parameters, 
        /// responses, and request body content.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="details"></param>
        /// <returns></returns>
        private List<OpenXmlElement> CreateNewRequestTypeSection(string type, RequestType details)
        {
            List<OpenXmlElement> sections = new List<OpenXmlElement>();

            ParagraphProperties properties = new ParagraphProperties(new SpacingBetweenLines { After = "0" });
            Paragraph container = new Paragraph(properties);
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
            }

            sections.Add(container);

            if (details.Parameters != null)
            {
                sections.AddRange(CreateNewParameterSection(details.Parameters));
            }

            if (details.RequestBody != null) 
            {
                sections.AddRange(CreateNewRequestBodySection(details.RequestBody));
            }

            if (details.Responses != null)
            {
                sections.AddRange(CreateNewResponseSection(details.Responses));
            }

            return sections;
        }

        /// <summary>
        /// Create a new Parameter section for a HTTP request type.
        /// </summary>
        /// <returns></returns>
        private List<OpenXmlElement> CreateNewParameterSection(IEnumerable<Parameter> parameters)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();

            ParagraphProperties properties = new ParagraphProperties(new SpacingBetweenLines { After = "0" });
            Paragraph container = new Paragraph(properties);
            //Run paramSection = container.AppendChild(new Run());
            container.AppendChild(Format.CreateBoldTextLine("Parameters", TEXT_FONT_SIZE));
            elements.Add(container);

            foreach (Parameter param in parameters)
            {
                elements.Add(CreateNewParameter(param));
            }
            return elements;
        }

        /// <summary>
        /// Creates a single new formatted parameter sub-section. 
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        private Run CreateNewParameter(Parameter param)
        {                    
            Run container = new Run();

            int bulletId = CreateNewBulletedList();
            Run typeRun = Format.CreateLabelValuePair($"{param.Name} : ", param.Schema.DisplayTypeText, JSON_FONT_SIZE);
            Paragraph typeParagraph = Format.CreateBulletedListItem(bulletId, 1, typeRun);
            container.AppendChild(typeParagraph);

            if (!string.IsNullOrEmpty(param.Description))
            {
                Run summary = Format.CreateTextLine(param.Description, JSON_FONT_SIZE);
                Paragraph summaryParagraph = Format.CreateBulletedListItem(bulletId, 2, summary);
                container.AppendChild(summaryParagraph);
            }

            Run locationRun = Format.CreateLabelValuePair("In: ", param.In, JSON_FONT_SIZE);
            Paragraph locationParagraph = Format.CreateBulletedListItem(bulletId, 2, locationRun);
            container.AppendChild(locationParagraph);

            if (param.Required)
            {
                Run required = Format.CreateBoldTextLine("Required", JSON_FONT_SIZE);
                Paragraph reqParagraph = Format.CreateBulletedListItem(bulletId, 2, required);
                container.AppendChild(reqParagraph);
            }

            return container;
        }

        /// <summary>
        /// Creates a new HTTP POST request body section, including schema obj formatting.
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        private List<OpenXmlElement> CreateNewRequestBodySection(RequestBody body)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();

            ParagraphProperties properties = new ParagraphProperties(new SpacingBetweenLines { After = "0" });
            Paragraph reqBody = new Paragraph(properties);
            reqBody.AppendChild(Format.CreateBoldTextLine("Request Body", TEXT_FONT_SIZE));
            elements.Add(reqBody);

            if (!string.IsNullOrEmpty(body.Description))
            {
                Run description = new Run();
                description.AppendChild(new TabChar());
                description.AppendChild(Format.CreateTextLine(body.Description, JSON_FONT_SIZE));
                description.AppendChild(new CarriageReturn());
                elements.Add(description);
            }
            
            if (body.Content.TryGetValue("application/json", out Content? appJsonContent))
            {
                Run appJson = new Run();
                appJson.AppendChild(new TabChar());
                appJson.AppendChild(Format.CreateBoldTextLine("application/json", JSON_FONT_SIZE));
                appJson.AppendChild(new CarriageReturn());
                appJson.AppendChild(CreateFormattedSchema(appJsonContent.Schema, 2));
                elements.Add(appJson);
            }

            if(body.Content.TryGetValue("multipart/form-data", out Content? formData))
            {
                Run formRun = new Run();
                formRun.AppendChild(new TabChar());
                formRun.AppendChild(Format.CreateBoldTextLine("multipart/form-data", JSON_FONT_SIZE));
                formRun.AppendChild(new CarriageReturn());
                formRun.AppendChild(CreateFormattedSchema(formData.Schema, 2));
                elements.Add(formRun);
            }

            return elements;
        }

        private List<OpenXmlElement> CreateNewResponseSection(Dictionary<string, Response> responses)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();

            ParagraphProperties properties = new ParagraphProperties(new SpacingBetweenLines { After = "0" });
            Paragraph responseParagraph = new Paragraph(properties);
            responseParagraph.AppendChild(Format.CreateBoldTextLine("Responses", TEXT_FONT_SIZE));
            elements.Add(responseParagraph);

            foreach(KeyValuePair<string, Response> response in responses)
            {
                string code = response.Key;
                Response responseValue = response.Value;
                Run run = new Run();
                run.AppendChild(new TabChar());
                run.AppendChild(Format.CreateLabelValuePair($"{code}: ", responseValue.Description, JSON_FONT_SIZE));
                run.AppendChild(new CarriageReturn());

                if (responseValue.Content != null && responseValue.Content.TryGetValue("application/json", out Content? appJsonContent)) 
                {
                    run.AppendChild(new TabChar());
                    run.AppendChild(Format.CreateBoldTextLine("application/json", JSON_FONT_SIZE));
                    run.AppendChild(new CarriageReturn());

                    run.AppendChild(CreateFormattedSchema(appJsonContent.Schema, 2));
                }

                elements.Add(run);
            }

            return elements;
        }

        /// <summary>
        /// Creates a heading for a controller that will contain the various available endpoints.
        /// </summary>
        /// <param name="controllerName"></param>
        /// <returns></returns>
        private static Paragraph CreateNewControllerHeading(string controllerName)
        {
            Justification centered = new Justification() { Val = JustificationValues.Center };
            ParagraphProperties props = new ParagraphProperties(centered);
            Paragraph paragraph = new Paragraph(props);

            paragraph.AppendChild(new CarriageReturn());
            string headingText = $"{controllerName} Endpoints";
            paragraph.AppendChild(Format.CreateBoldTextLine(headingText, TITLE_FONT_SIZE));
            
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
                        if(propSchema == schema)
                        {
                            propertyRun.AppendChild(Format.CreateLabelValuePair($"{propName}: ", "Same object type as parent", JSON_FONT_SIZE));
                            propertyRun.AppendChild(new CarriageReturn());
                        } else
                        {
                            Run propParagraph = propertyRun.AppendChild(new Run());
                            propParagraph.AppendChild(CreateFormattedSchema(propSchema, numTabs + 1));
                        }

                    } else
                    {
                        propertyRun.AppendChild(Format.CreateLabelValuePair($"{propName}: ", propSchema.DisplayTypeText, JSON_FONT_SIZE));
                        propertyRun.AppendChild(new CarriageReturn());

                        if (propSchema.Items != null)
                        {
                            propertyRun.AppendChild(CreateFormattedSchema(propSchema, numTabs + 1));

                        } else
                        {
                            // if there are items, it's an array and we'll let the description fall under it
                            if (!string.IsNullOrEmpty(propSchema.Description))
                            {
                                for (int i = 0; i < numTabs; i++)
                                {
                                    propertyRun.AppendChild(new TabChar());
                                }
                                propertyRun.AppendChild(Format.CreateTextLine(propSchema.Description, JSON_FONT_SIZE));
                                propertyRun.AppendChild(new CarriageReturn());
                            }
                        }
                    }
                }
            }
            else if (schema.Type == "array")
            {
                Run arrayRun = container.AppendChild(new Run());

                for (int i = 0; i < numTabs; i++)
                {
                    arrayRun.AppendChild(new TabChar());
                }

                string valueText = !string.IsNullOrEmpty(schema.Items.Ref) ? "object" : schema.Items.DisplayTypeText;
                arrayRun.AppendChild(Format.CreateLabelValuePair("Array Content: ", valueText, JSON_FONT_SIZE));
                arrayRun.AppendChild(new CarriageReturn());

                if (!string.IsNullOrEmpty(schema.Description))
                {
                    Paragraph desc = arrayRun.AppendChild(new Paragraph());

                    for (int i = 0; i < numTabs; i++)
                    {
                        desc.AppendChild(new TabChar());
                    }
                    desc.AppendChild(Format.CreateTextLine(schema.Description, JSON_FONT_SIZE));
                    arrayRun.AppendChild(new CarriageReturn());
                }

                if (valueText == "object")
                {
                    Run itemsRun = arrayRun.AppendChild(new Run());
                    itemsRun.AppendChild(CreateFormattedSchema(schema.Items, numTabs + 1));
                }
            }
            else
            {
                Run itemRun = container.AppendChild(new Run());

                for (int i = 0; i < numTabs; i++)
                {
                    itemRun.AppendChild(new TabChar());
                }

                itemRun.AppendChild(Format.CreateTextLine(schema.DisplayTypeText, JSON_FONT_SIZE));

                if (!string.IsNullOrEmpty(schema.Description))
                {
                    itemRun.AppendChild(new CarriageReturn());
                    Paragraph desc = itemRun.AppendChild(new Paragraph());

                    for (int i = 0; i < numTabs; i++)
                    {
                        desc.AppendChild(new TabChar());
                    }
                    desc.AppendChild(Format.CreateTextLine(schema.Description, JSON_FONT_SIZE));
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
        private void CompileDocument(Dictionary<string, List<OpenXmlElement>> paragraphs)
        {
            foreach(KeyValuePair<string, OpenXmlElement> components in _componentBulletLists)
            {
                Body.AppendChild(components.Value);
            }

            foreach (KeyValuePair<string, List<OpenXmlElement>> items in paragraphs) 
            {
                OpenXmlElement heading = CreateNewControllerHeading(items.Key);
                Body.AppendChild(heading);

                foreach (OpenXmlElement paragraph in items.Value) 
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
