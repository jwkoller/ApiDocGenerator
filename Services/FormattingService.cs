using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace APIDocGenerator.Services
{
    public static class FormattingService
    {
        public static Run CreateLabelValuePair(string label, string value, string fontSize)
        {
            Run container = new Run();
            Run labelRun = container.AppendChild(new Run());
            RunProperties labelProps = new RunProperties
            {
                Bold = new Bold(),
                FontSize = new FontSize { Val = fontSize }
            };
            labelRun.AppendChild(labelProps);
            labelRun.AppendChild(new Text { Text = label, Space = SpaceProcessingModeValues.Preserve });

            Run valueRun = container.AppendChild(new Run());
            RunProperties valueProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize }
            };
            valueRun.AppendChild(valueProps);
            valueRun.AppendChild(new Text { Text = value, Space = SpaceProcessingModeValues.Preserve });

            return container;
        }

        public static Run CreateTextLine(string text, string fontSize)
        {
            Run textRun = new Run();
            RunProperties textProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize }
            };
            textRun.Append(textProps);
            textRun.AppendChild(new Text { Text = text, Space = SpaceProcessingModeValues.Preserve });

            return textRun;
        }

        public static Run CreateBoldTextLine(string text, string fontSize)
        {
            Run textRun = new Run();
            RunProperties textProps = new RunProperties
            {
                FontSize = new FontSize { Val = fontSize },
                Bold = new Bold()
            };
            textRun.Append(textProps);
            textRun.AppendChild(new Text { Text = text, Space = SpaceProcessingModeValues.Preserve });

            return textRun;
        }

        public static Paragraph CreateBulletedListItem(int numberingId, int indent, Run text)
        {
            int indentUnitSize = 240;
            string indentValue = $"{indent * indentUnitSize}";

            var propNumberProps = new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numberingId });
            var propSpaceBetween = new SpacingBetweenLines { After = "0" };
            var propIndentation = new Indentation { Left = indentValue };

            Paragraph paragraph = new Paragraph(new ParagraphProperties(propNumberProps, propSpaceBetween, propIndentation));
            paragraph.AppendChild(text);

            return paragraph;
        }
    }
}
