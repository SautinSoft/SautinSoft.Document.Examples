using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.IO;

namespace Sample
{
    class Sample
    {

        static void Main()
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            /// <summary>
            /// Filling out a document from a database.
            /// </summary>
            /// <remarks>
            /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/read_content_from_db.php
            /// </remarks>

            string db_content = @"
                {
                    'SomeField': 'SautinSoft',
                    'Field': 'RATES',
                    'Table1': [
                        {
                            'Name': 'John S. Director',
                            'Description': '$475.00',
                            'Bla': 'In case of deviations from the schedule, analyzes the reasons for the deviations and takes measures to correct the situation.'
                        },
                        {
                            'Name': 'Donna Di Group',
                            'Description': '$150.67',
                            'Bla': 'Plans to perform the work using own resources or the need to involve a subcontractor in order to determine the cost of designing the facility and ensure the release of documentation within the timeframe specified in the technical specifications.'
                        },
                        {
                            'Name': 'Frank H.Rogel',
                            'Description': '$623.09',
                            'Bla': 'Forms a commercial proposal for his sections and submits it to the marketing and tender department for his section of work.'
                        }
                    ],
                    'Table2': [
                        {
                            'Name': 'Name 4',
                            'Description': 'Description 4'
                        },
                        {
                            'Name': 'Name 5',
                            'Description': 'Description 5'
                        },
                        {
                            'Name': 'Name 6',
                            'Description': 'Description 6'
                        }
                    ]
                }
            ";

            var jsonSettings = new JsonSerializerSettings { Converters = { new JsonObjectConverter() } };
            var jsonSource = JsonConvert.DeserializeObject<Dictionary<string, object>>(db_content, jsonSettings);

            var document = DocumentCore.Load(@"..\..\..\Template.docx");
            document.MailMerge.Execute(jsonSource);
            document.Save(@"..\..\..\Result.pdf");
        }

        public sealed class JsonObjectConverter : JsonConverter
        {
            public override bool CanConvert(Type objectType)
            {
                return objectType == typeof(object);
            }

            public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
            {
                switch (reader.TokenType)
                {
                    case JsonToken.StartObject:
                        return serializer.Deserialize<Dictionary<string, object>>(reader);

                    case JsonToken.StartArray:
                        var array = JArray.Load(reader);
                        var list = new List<Dictionary<string, object>>();

                        foreach (var item in array)
                        {
                            if (item.Type == JTokenType.Object)
                                list.Add(item.ToObject<Dictionary<string, object>>());
                        }

                        return list;

                    case JsonToken.String:
                        return reader.Value?.ToString();

                    case JsonToken.Integer:
                    case JsonToken.Float:
                        return Convert.ToDouble(reader.Value);

                    case JsonToken.Boolean:
                        return Convert.ToBoolean(reader.Value);

                    default:
                        var token = JToken.ReadFrom(reader);
                        return token.ToObject<object>();
                }
            }

            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                serializer.Serialize(writer, value);
            }
        }
    }
}