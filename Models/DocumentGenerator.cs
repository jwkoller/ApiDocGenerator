using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using DocumentFormat.OpenXml;
using CommunityToolkit.Mvvm.DependencyInjection;

namespace APIDocGenerator.Services
{
    public class DocumentGenerator
    {
        public string DocumentName { get; private set; }
        public WordprocessingDocument Document { get; private set; }
        public MainDocumentPart MainPart {  get; private set; }
        public Body Body { get; private set; }

        public DocumentGenerator(string destination, string fileName)
        {
            DocumentName = fileName;
            Document = WordprocessingDocument.Create($"{destination}/{fileName}.docx", WordprocessingDocumentType.Document);
            MainPart = Document.AddMainDocumentPart();
            MainPart.Document = new Document();
            Body = MainPart.Document.AppendChild(new Body());
            AddDocumentStyles();
        }

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
            props.FontSize = new FontSize() { Val = "36"};

            run.Append(props);
            run.AppendChild(new Text(Environment.NewLine));
            run.AppendChild(new Text(heading));
            run.AppendChild(new Text(Environment.NewLine));
            paragraph.AppendChild(run);
        }

        public void WriteNewLine(string newLine)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();         
            Run run = last.AppendChild(new Run());
            run.AppendChild(new Text(Environment.NewLine));
            run.AppendChild(new Text(newLine));
        }

        public void AddDocumentStyles()
        {
            if(Document.MainDocumentPart == null)
            {
                throw new ArgumentNullException();
            }

            StyleDefinitionsPart? stylePart = Document.MainDocumentPart?.StyleDefinitionsPart;

            if (stylePart == null) {
                stylePart = Document.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
            }

            Styles? stylesCollection = stylePart?.Styles;

            if(stylesCollection == null)
            {
                stylesCollection = new Styles();
                stylesCollection.Save(stylePart);
            }

            Style style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading 1",
                CustomStyle = true,
                Default = true
            };

            style.Append(new StyleName() { Val = "Heading 1" });

            StyleRunProperties props = new StyleRunProperties();
            props.Append(new Color() { ThemeColor = ThemeColorValues.Accent2 });
            props.Append(new Bold());
            props.Append(new FontSize() { Val = "36" });

            style.Append(props);
            stylesCollection.Append(style);
        }

        public void Save()
        {
            Document.Dispose();
        }
    }
}
