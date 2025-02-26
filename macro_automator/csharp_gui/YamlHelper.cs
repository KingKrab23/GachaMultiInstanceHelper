using System;
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace MacroAutomatorGUI
{
    public static class YamlHelper
    {
        /// <summary>
        /// Serialize an object to a YAML file
        /// </summary>
        public static void SaveToYaml<T>(T obj, string filePath)
        {
            try
            {
                var serializer = new SerializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();

                string yaml = serializer.Serialize(obj);
                File.WriteAllText(filePath, yaml);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving YAML file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Deserialize a YAML file to an object
        /// </summary>
        public static T LoadFromYaml<T>(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"YAML file not found: {filePath}");
                }

                string yaml = File.ReadAllText(filePath);
                
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();

                return deserializer.Deserialize<T>(yaml);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading YAML file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Convert macro sequences to YAML format for Python script
        /// </summary>
        public static void SaveMacroSequences(List<MacroSequence> sequences, string filePath)
        {
            try
            {
                // Convert to the format expected by the Python script
                var macrosDict = new Dictionary<string, object>();
                
                foreach (var sequence in sequences)
                {
                    var sequenceDict = new Dictionary<string, object>
                    {
                        ["actions"] = ConvertActionsToDict(sequence.Actions),
                        ["iteration_delay"] = 0.5
                    };
                    
                    macrosDict[sequence.Name] = sequenceDict;
                }
                
                var config = new Dictionary<string, object>
                {
                    ["settings"] = new Dictionary<string, object>
                    {
                        ["default_delay"] = 0.1
                    },
                    ["macros"] = macrosDict
                };
                
                SaveToYaml(config, filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving macro sequences: {ex.Message}", ex);
            }
        }
        
        private static List<Dictionary<string, object>> ConvertActionsToDict(List<MacroAction> actions)
        {
            var result = new List<Dictionary<string, object>>();
            
            foreach (var action in actions)
            {
                var actionDict = new Dictionary<string, object>
                {
                    ["type"] = action.Type
                };
                
                // Add all parameters
                foreach (var param in action.Parameters)
                {
                    actionDict[param.Key] = param.Value;
                }
                
                result.Add(actionDict);
            }
            
            return result;
        }
    }
}
