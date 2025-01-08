namespace APIDocGenerator.Models.JsonParse
{
    public class Responses
    {
        public int ResponseCode { get; set; }
        public string Description { get; set; }
        public IEnumerable<Content> Content { get; set; }
    }
}
