namespace APIDocGenerator.Models.JsonParse
{
    public class Components
    {
        public Dictionary<string, Schema> Schemas { get; set; }
        public Dictionary<string, Schema> SecuritySchemes { get; set; }
    }
}
