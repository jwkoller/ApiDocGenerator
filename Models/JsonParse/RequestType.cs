namespace APIDocGenerator.Models.JsonParse
{
    public class RequestType
    {
        public IEnumerable<string> Tags { get; set; }
        public string Type { get; set; }
        public string Summary { get; set; }
        public IEnumerable<Parameter> Parameters { get; set; }
        public RequestBody RequestBody { get; set; }
        public Responses Responses { get; set; }
    }
}
