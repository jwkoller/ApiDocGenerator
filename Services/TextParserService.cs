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
            List<string> listOfLines = lines.ToList();
            string firstHttp = listOfLines.First(x => x.StartsWith("[Http"));
            int index = listOfLines.IndexOf(firstHttp);

            for (int i = index - 1; i > -1; i--)
            {
                if (!listOfLines[i].StartsWith("///") && !listOfLines[i].StartsWith('['))
                {
                    index = i + 1; 
                    break;
                }
            }

            return lines.Skip(index);
        }

        /// <summary>
        /// Returns a list of key value pairs with of the parsed XML comments, with the key of each KVP being the type,
        /// and the value being the content.
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="currIdx"></param>
        /// <returns></returns>
        public (int, List<KeyValuePair<string, string>>) GetParsedXMLString(List<string> lines, int currIdx)
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
            string methodSignature = string.Empty;
            for(int i = lastIdx + 1; i < lines.Count; i++)
            {
                if (lines[i].StartsWith("public "))
                {
                    methodSignature = lines[i];
                    int nextLine = 1;
                    // longer method sigs can be split over multiple lines
                    while (methodSignature.LastIndexOf(')') == -1 && methodSignature.LastIndexOf('{') == -1)
                    {
                        methodSignature += $" {lines[i + nextLine]}";
                        nextLine++;
                    }
                    break;
                }
            }

            // get the comment block and convert back to single XML parseable string
            IEnumerable<string> comments = lines.GetRange(currIdx, (lastIdx - currIdx) + 1).Select(x => x.Replace("///", "").Trim());
            string xmlString = string.Join("", comments);
            // .Parse won't do fragments like the comment structure, needs a root node
            XElement xml = XElement.Parse($"<root>{xmlString}</root>");

            // substring out the parameters - this is predicated upon only one set of parens in the method sig
            string methodSigParams = methodSignature[(methodSignature.IndexOf('(') + 1)..];
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
                listOfParamsStrings.Add($"({parameterOrigin}{name}) {nodeValue}");
            }

            string? summaryElemString = (string?)xml.Element("summary");
            string? returnsElemString = (string?)xml.Element("returns");

            IEnumerable<XElement> exceptNodes = xml.Elements("exception");
            IEnumerable<string> exceptionElemStrings = exceptNodes.Select(x => $"(Type: {x.Attribute("cref")?.Value}) {x.Value}");
           
            List<KeyValuePair<string, string>> commentLines = new List<KeyValuePair<string, string>>();
            if(!string.IsNullOrEmpty(summaryElemString))
            {
                commentLines.Add(new KeyValuePair<string, string>("Summary: ", summaryElemString));
            }

            foreach (string line in listOfParamsStrings)
            {
                commentLines.Add(new KeyValuePair<string, string>("Param: ", line));
            }

            if (!string.IsNullOrEmpty(returnsElemString))
            {
                commentLines.Add(new KeyValuePair<string, string>("Returns: ", returnsElemString));
            }

            foreach(string line in exceptionElemStrings)
            {
                commentLines.Add(new KeyValuePair<string, string>("Exception: ", line));
            }

            return (lastIdx, commentLines);
        }
    }
}
