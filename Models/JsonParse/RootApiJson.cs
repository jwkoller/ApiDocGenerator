﻿namespace APIDocGenerator.Models.JsonParse
{
    public class RootApiJson
    {
        public string OpenApi {  get; set; }
        public ApiInfo Info { get; set; }
        public Dictionary<string, Route> Paths { get; set; }
        public Components Components { get; set; }
    }
}
