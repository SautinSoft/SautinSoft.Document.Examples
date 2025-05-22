using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using System.Text.Json;
using System.Text.Json.Serialization;

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

            string db_content = """
            {
        "SomeField": "SautinSoft",
        "Field": "RATES",
        "Table1": [
        {
            "Name": "John S. Director",
            "Description": "$475.00",
            "Bla": "In case of deviations from the schedule, analyzes the reasons for the deviations and takes measures to correct the situation."
        },
        {
            "Name": "Donna Di Group",
            "Description": "$150.67",
            "Bla": "Plans to perform the work using own resources or the need to involve a subcontractor in order to determine the cost of designing the facility and ensure the release of documentation within the timeframe specified in the technical specifications."
        },
        {
            "Name": "Frank H.Rogel",
            "Description": "$623.09",
            "Bla": "Forms a commercial proposal for his sections and submits it to the marketing and tender department for his section of work."
        }
        ],
        "Table2": [
        {
            "Name": "Name 4",
            "Description": "Description 4"
        },
        {
            "Name": "Name 5",
            "Description": "Description 5"
        },
        {
            "Name": "Name 6",
            "Description": "Description 6"
        }
        ]
            }
    """;

            var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true, Converters = { new JsonObjectConverter() } };
            var jsonSource = JsonSerializer.Deserialize<Dictionary<string, object>>(db_content, jsonOptions);

            var document = DocumentCore.Load(@"..\..\..\Template.docx");
            document.MailMerge.Execute(jsonSource);
            document.Save(@"..\..\..\Result.pdf");
        }

        public sealed class JsonObjectConverter : JsonConverter<object>
        {
         public override object Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                // The objects ({}) are mapped as Dictionary<string, object>
                // and arrays ([]) are mapped as List<Dictionary<string, object>>.
                return reader.TokenType switch
                {
                    JsonTokenType.StartObject => JsonSerializer.Deserialize<Dictionary<string, object>>(ref reader, options),
                    JsonTokenType.StartArray => JsonSerializer.Deserialize<List<Dictionary<string, object>>>(ref reader, options),
                    JsonTokenType.String => reader.GetString(),
                    JsonTokenType.Number => reader.GetDouble(),
                    JsonTokenType.True or JsonTokenType.False => reader.GetBoolean(),
                    _ => JsonSerializer.Deserialize<object>(ref reader, options)
                };
            }
            public override void Write(Utf8JsonWriter writer, object value, JsonSerializerOptions options)
                => JsonSerializer.Serialize(writer, value, value.GetType(), options);
        }
    }
}