namespace APIDocGenerator.Models.JsonParse
{
    public class RequestBody
    {
        public string Description { get; set; }
        public Dictionary<string, Content> Content { get; set; }
    }
}
