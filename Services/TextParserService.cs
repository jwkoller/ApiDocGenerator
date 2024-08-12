using APIDocGenerator.Models;
using System.Text;
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

            // find the method signature line for this block of comments
            int methodSigIdx = 0;
            for(int i = lastIdx + 1; i < lines.Count; i++)
            {
                if (lines[i].StartsWith("public "))
                {
                    methodSigIdx = i;
                    break;
                }
            }

            // get the comment block and convert back to single XML parseable string
            IEnumerable<string> comments = lines.GetRange(currIdx, (lastIdx - currIdx) + 1).Select(x => x.Replace("///", "").Trim());
            string xmlString = string.Join("", comments);
            // .Parse won't do fragments like the comment structure, needs a root node
            XElement xml = XElement.Parse($"<root>{xmlString}</root>");

            // get the method signature string
            string methodSig = lines[methodSigIdx];
            // substring out the parameters - this is predicated upon only one set of parens in the method sig
            string methodSigParams = methodSig[(methodSig.IndexOf('(') + 1)..];
            methodSigParams = methodSigParams[..methodSigParams.LastIndexOf(')')];
            // split any parameters into parts
            IEnumerable<MethodSignatureParams> paramObjects = methodSigParams.Split(",").Select(x => new MethodSignatureParams(x.Trim()));

            // go through each parameter comment and get where it's coming from by the method sig info
            IEnumerable<XElement> paramNodes = xml.Elements("param");
            List<string> listOfParamsStrings = new List<string>();
            foreach(XElement node in paramNodes)
            {
                string? name = node.Attribute("name")?.Value;
                string nodeValue = node.Value;
                string parameterOrigin = string.Empty;
                MethodSignatureParams? paramObj = paramObjects.FirstOrDefault(x => x.Name == name);
                if(paramObj != default)
                {
                    parameterOrigin = $"{paramObj.FromLocation} ";
                }
                listOfParamsStrings.Add($"Param: ({parameterOrigin}{name}) {nodeValue}");
            }

            string paramElemString = string.Join(Environment.NewLine, listOfParamsStrings);
            string? summaryElemString = (string?)xml.Element("summary");
            string? returnsElemString = (string?)xml.Element("returns");

            IEnumerable<XElement> exceptNodes = xml.Elements("exception");
            string? exceptElem = string.Join(Environment.NewLine, exceptNodes.Select(x => $"Exception: (Type: {x.Attribute("cref")?.Value}) {x.Value}"));
            
            StringBuilder output = new StringBuilder();
            output.Append($"Summary: {summaryElemString}");
            if(!string.IsNullOrEmpty(paramElemString)) 
            {
                output.AppendLine();
                output.Append(paramElemString);
            }
            
            if(!string.IsNullOrEmpty(returnsElemString))
            {
                output.AppendLine();
                output.Append($"Returns: {returnsElemString}");
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
