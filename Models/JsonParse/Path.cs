namespace APIDocGenerator.Models.JsonParse
{
    public class Path
    {
        public string Route { get; set; }
        public RequestType? Get {  get; set; }
        public RequestType? Post { get; set; }
        public RequestType? Put { get; set; }
        public RequestType? Delete { get; set; }
    }
}
