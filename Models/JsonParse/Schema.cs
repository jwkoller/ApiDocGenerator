﻿using Newtonsoft.Json;

namespace APIDocGenerator.Models.JsonParse
{
    public class Schema
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public Schema Items { get; set; }
        public string Description { get; set; }
        public string Format { get; set; }
        public bool Nullable { get; set; }
        public bool ReadOnly { get; set; }
        public Dictionary<string, Schema> Properties { get; set; }
        [JsonProperty(PropertyName = "$ref")]
        public string Ref { get; set; }
        [JsonIgnore]
        public string DisplayTypeText { get => string.IsNullOrEmpty(Format) ? Type : Format; }
    }
}
