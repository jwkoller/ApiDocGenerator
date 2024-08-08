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
            props.FontSize = new FontSize() { Val = "40"};

            run.Append(props);
            run.AppendChild(new Text(Environment.NewLine));
            run.AppendChild(new Text(heading));
            run.AppendChild(new Text(Environment.NewLine));
            paragraph.AppendChild(run);
        }

        public void WriteCommentLine(string text)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();         
            Run run = last.AppendChild(new Run());

            run.AppendChild(new Text(Environment.NewLine));
            run.AppendChild(new Text(text));
            run.AppendChild(new Text(Environment.NewLine));
        }

        public void WriteRouteLine(string type, string text)
        {
            Paragraph last = Body.Elements<Paragraph>().Last();
            Run run = last.AppendChild(new Run());
            RunProperties props = new RunProperties();
            props.FontSize = new FontSize() { Val = "28" };
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
            run.AppendChild(new Text($"{type}: "));

            Run next = last.AppendChild(new Run());
            next.AppendChild(new Text(text));
            next.AppendChild(new Text(Environment.NewLine));
        }


        public void Save()
        {
            Document.Dispose();
        }
    }
}
