


// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
static Dictionary<string, object> FlattenJson(JsonNode? jsonObject)
{
    var result = new Dictionary<string, object>();
    
    if (jsonObject != null)
    {
        Flatten(jsonObject, "", result);
    }

    return result;
}

static void Flatten(JsonNode node, string prefix, Dictionary<string, object> result)
{
    if (node is JsonObject obj)
    {
        foreach (var kvp in obj)
        {
            string newPrefix = string.IsNullOrEmpty(prefix) ? kvp.Key : $"{prefix}.{kvp.Key}";
            Flatten(kvp.Value, newPrefix, result);
        }
    }
    else if (node is JsonArray array)
    {
        for (int i = 0; i < array.Count; i++)
        {
            Flatten(array[i], $"{prefix}[{i}]", result);
        }
    }
    else if (node != null)
    {
        result[prefix] = node.ToString();
    }
    else
    {
        result[prefix] = null;
    }
}
#region main

var inputFile = args[0];
Console.WriteLine($"Input: {inputFile}");

var jsonText = File.ReadAllText(inputFile);

var node = JsonNode.Parse(jsonText);
foreach (var item in node.AsArray())
{
    var flattenedd = FlattenJson(item["project"]);
    foreach (var item2 in flattenedd)
    {
        Console.WriteLine(item2.Key);
    }


}
#endregion