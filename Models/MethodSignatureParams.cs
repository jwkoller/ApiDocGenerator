using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APIDocGenerator.Models
{
    internal class MethodSignatureParams
    {
        public string FromLocation { get; set; }
        public string Type {  get; set; }
        public string Name { get; set; }

        public MethodSignatureParams(string inputLine)
        {
            string[] linesSplit = inputLine.Split(" ");
            switch (linesSplit.Length)
            {
                case 2:
                    FromLocation = string.Empty;
                    Type = linesSplit[0];
                    Name = linesSplit[1];
                    break;
                case 3:
                    FromLocation = linesSplit[0];
                    Type = linesSplit[1];
                    Name = linesSplit[2];
                    break;
                default:
                    FromLocation = string.Empty;
                    Type = string.Empty;
                    Name = string.Empty;
                    break;
            }
        }
    }
}
