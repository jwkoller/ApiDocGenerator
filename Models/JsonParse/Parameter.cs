﻿namespace APIDocGenerator.Models.JsonParse
{
    public class Parameter
    {
        public string Name { get; set; }
        public string In {  get; set; }
        public string Description { get; set; }
        public bool Required { get; set; }
        public Schema Schema { get; set; }
    }
}
