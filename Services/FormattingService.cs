using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace APIDocGenerator.Services
{
    public static class FormattingService
    {
        public static Run CreateLabelValuePair(string label, string value, string fontSize)
        {
            Run container = new Run();

            RunProperties labelProps = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = fontSize }
            };
            container.AppendChild(CreateTextLine(label, labelProps));


            RunProperties valueProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize }
            };
            container.AppendChild(CreateTextLine(value, valueProps));

            return container;
        }

        public static Run CreateTextLine(string text, string fontSize)
        {
            RunProperties textProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize }
            };

            return CreateTextLine(text, textProps);
        }

        public static Run CreateBoldTextLine(string text, string fontSize)
        {
            RunProperties textProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize },
                Bold = new Bold(),
            };

            return CreateTextLine(text, textProps);
        }

        public static Run CreateTextLine(string text, RunProperties props)
        {
            Run textRun = new Run();
            textRun.Append(props);
            textRun.AppendChild(new Text { Text = text, Space = SpaceProcessingModeValues.Preserve });

            return textRun;
        }

        public static Paragraph CreateBulletedListItem(int numberingId, int indent, Run text)
        {
            int indentUnitSize = 240;
            string indentValue = $"{(indent + 1) * indentUnitSize}";

            var propNumberProps = new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numberingId });
            var propSpaceBetween = new SpacingBetweenLines { After = "0" };
            
            var propIndentation = new Indentation { Left = indentValue, Hanging = "300" };

            Paragraph paragraph = new Paragraph(new ParagraphProperties(propNumberProps, propSpaceBetween, propIndentation));
            paragraph.AppendChild(text);

            return paragraph;
        }
    }
}
