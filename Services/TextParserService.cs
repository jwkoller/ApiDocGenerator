using DocumentFormat.OpenXml.Office2013.Excel;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace APIDocGenerator.Services
{
    public class TextParserService
    {
        /// <summary>
        /// Returns a concatenated string of versions controller is available for, based on [ApiVersion] attribute.
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public string GetVersionInfo(IEnumerable<string> lines) {
            IEnumerable<string> versions = lines.Where(x => x.StartsWith("[ApiVersion"));
            IEnumerable<string> parsedVersions = versions.Select(x => $"v{x.Substring(x.IndexOf('(') + 1, 1)}");

            return string.Join(", ", parsedVersions);
        }

        /// <summary>
        /// Returns the type of endpont (HttpGet, HttpPost, etc) and the routing information as defined by the [Http...()] attribute.
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public (string, string) GetEndPointRouting(string line)
        {
            string parsed = new StringBuilder(line).Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").ToString();
            string[] split = parsed.Split('"');
            string type = split[0];
            string endpoint = split.Length > 1 ? $"/{split[1]}" : "";

            return (type, endpoint);
        }

        /// <summary>
        /// Returns the lines from the file starting at the comment section for the first endpoint, 
        /// as defined by the first instance of the [Http...()] type attribute.
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public IEnumerable<string> GetLinesAtFirstEndpoint(IEnumerable<string> lines)
        {
            string firstHttp = lines.First(x => x.StartsWith("[Http"));
            int index = lines.ToList().IndexOf(firstHttp);

            for (int i = index - 1; i > -1; i--)
            {
                if (!lines.ElementAt(i).StartsWith("///"))
                {
                    index = i + 1; 
                    break;
                }
            }

            return lines.Skip(index);
        }

        /// <summary>
        /// Returns a formatted string of the xml comments in plain language with elements separated by line breaks, and the index of 
        /// the last line of comments for this section so they can be skipped over.
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="currIdx"></param>
        /// <returns></returns>
        public (int, string) GetParsedXMLString(List<string> lines, int currIdx)
        {
            int lastIdx = 0;
            for(int i = currIdx + 1; i < lines.Count; i++)
            {
                if (!lines[i].StartsWith("///"))
                {
                    lastIdx = i - 1;
                    break;
                }
            }

            IEnumerable<string> comments = lines.GetRange(currIdx, (lastIdx - currIdx) + 1).Select(x => x.Replace("///", "").Trim());
            string xmlString = string.Join("", comments);
            // .Parse won't do fragments like the comments, needs a root node
            XElement xml = XElement.Parse($"<root>{xmlString}</root>");

            string? summaryElem = (string?)xml.Element("summary");

            IEnumerable<XElement> paramNodes = xml.Elements("param");
            string? paramElem = string.Join(Environment.NewLine, paramNodes.Select(x => $"Param: ({x.Attribute("name")?.Value}) {x.Value}"));

            string? returnsElem = (string?)xml.Element("returns");

            IEnumerable<XElement> exceptNodes = xml.Elements("exception");
            string? exceptElem = string.Join(Environment.NewLine, exceptNodes.Select(x => $"Exception: (Type: {x.Attribute("cref")?.Value}) {x.Value}"));
            
            StringBuilder output = new StringBuilder();
            output.Append($"Summary: {summaryElem}");
            if(!string.IsNullOrEmpty(paramElem)) 
            {
                output.AppendLine();
                output.Append(paramElem);
            }
            
            if(!string.IsNullOrEmpty(returnsElem))
            {
                output.AppendLine();
                output.Append($"Returns: {returnsElem}");
            }

            if (!string.IsNullOrEmpty(exceptElem))
            {
                output.AppendLine();
                output.Append(exceptElem);
            }

            return (lastIdx, output.ToString());
        }
    }
}
