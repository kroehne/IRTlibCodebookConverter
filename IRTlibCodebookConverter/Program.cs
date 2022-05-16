using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.Json;

namespace CodebookConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            { 
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("IRTlib: CodebookConverter ({0})\n", typeof(Program).Assembly.GetName().Version.ToString());
                Console.ResetColor();

                string _codebookExcelFile = "";
                string _codebookJSONFile = "";
                
                if (args.Length >= 2)
                {
                    _codebookExcelFile = args[0];
                    _codebookJSONFile = args[1];
                } 
                else
                {
                    Console.WriteLine("- Please provide two arguments: Path to the Excel-File (Input) and path to the JSON-File (Output)");
                    return;
                }

                if (!File.Exists(_codebookExcelFile))
                {
                    Console.WriteLine("- Codebook Excel-File (Input): {0} not found.", _codebookExcelFile);
                    return;
                }
                else
                {
                    Console.WriteLine("- Codebook Excel-File (Input): {0}", _codebookExcelFile);
                }

                Console.WriteLine("- Codebook JSON-File (Output): {0}", _codebookJSONFile);

                Convert(_codebookExcelFile, _codebookJSONFile);
            }
            catch (Exception _ex)
            {
                Console.WriteLine("Error:");
                Console.WriteLine(_ex.ToString());
            }
            
        }

        static void Convert(string ExcelFile, string JSONFile)
        {
            Dictionary<string, Variable> _variablesDict = new Dictionary<string, Variable>();

            using (var stream = new FileStream(ExcelFile, FileMode.Open))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                ISheet sheet = xssWorkbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);

                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    if (sheet.GetRow(row) != null && sheet.GetRow(row).GetCell(10) != null)
                    {
                        
                        string _resultText =  sheet.GetRow(row).GetCell(9).ToString(); 
                        string _itemname = sheet.GetRow(row).GetCell(1).StringCellValue;
                        string _taskname = sheet.GetRow(row).GetCell(2).StringCellValue;
                        string _classname = sheet.GetRow(row).GetCell(3).StringCellValue;
                        string _variableKey = _itemname + "." + _classname;
                        bool _userForRouting = sheet.GetRow(row).GetCell(10).ToString().Trim()=="0" ? false: true ;

                        if (!_variablesDict.ContainsKey(_variableKey))
                        {
                            if (_resultText == "1")
                            {
                                _variablesDict.Add(_variableKey, new Variable()
                                {
                                    Item = _itemname,
                                    Class = _classname,
                                    Task = _taskname,
                                    Name = sheet.GetRow(row).GetCell(5).StringCellValue,
                                    Label = RemoveUmlaute(sheet.GetRow(row).GetCell(6).StringCellValue),
                                    Type = "String",
                                    Values = new Dictionary<string, HitMiss>(),
                                    UseForRouting = _userForRouting

                                });
                            } 
                            else
                            {
                                _variablesDict.Add(_variableKey, new Variable()
                                {
                                    Item = _itemname,
                                    Class = _classname,
                                    Task = _taskname,
                                    Name = sheet.GetRow(row).GetCell(5).StringCellValue,
                                    Label = RemoveUmlaute(sheet.GetRow(row).GetCell(6).StringCellValue),
                                    Type = "Integer",
                                    Values = new Dictionary<string, HitMiss>(),
                                    UseForRouting = _userForRouting

                                });  
                            }
                         
                        }
                         
                        string _valueKey = sheet.GetRow(row).GetCell(4).ToString() + "_" + sheet.GetRow(row).GetCell(7).ToString();
                        if (!_variablesDict[_variableKey].Values.ContainsKey(_valueKey))
                            _variablesDict[_variableKey].Values.Add(_valueKey, new HitMiss()
                            {
                                Hit = sheet.GetRow(row).GetCell(4).ToString(),
                                Label = RemoveUmlaute(sheet.GetRow(row).GetCell(8).StringCellValue),
                                Value = sheet.GetRow(row).GetCell(7).ToString(),   
                                ResultText = _resultText
                            });;


                    }
                }

            }
             
            string _json = "{\n	\"DefaultItemMissingCode\": \"-94\",\n	\"MissingCodes\": [\n        {\n            \"Code\": \"-94\",\n            \"Label\": \"Timeout\",\n            \"Context\": \"Timeout\"\n        },\n        {\n            \"Code\": \"-91\",\n            \"Label\": \"Abbruch\",\n            \"Context\": \"Abort\"\n        },\n        {\n            \"Code\": \"-54\",\n            \"Label\": \"Designbedingt fehlend\",\n            \"Context\": \"SkippedByDesign\"\n        }\n    ],\n	\"CategoricalVariables\": [";

            List<string> _categoricalVariableKeyList = new List<string>(_variablesDict.Keys);

            for (var j = 0; j < _categoricalVariableKeyList.Count; j++)
            {
                var v = _variablesDict[_categoricalVariableKeyList[j]];
                _json += "{ \n";
                _json += "\"Values\": [\n";

                List<string> _valueKeyList = new List<string>(v.Values.Keys);
                for (int i = 0; i < _valueKeyList.Count; i++)
                {
                    _json += "\t{ \n ";
                    _json += "\t\t\"Hit\": \"" + v.Values[_valueKeyList[i]].Hit + "\",\n ";
                    _json += "\t\t\"Score\": \"" + v.Values[_valueKeyList[i]].Value + "\",\n ";
                    _json += "\t\t\"Label\": \"" + v.Values[_valueKeyList[i]].Label + "\"\n ";
                    _json += "\t}";
                    if (i != v.Values.Count - 1)
                        _json += ",\n ";
                    else
                        _json += "\n ";


                }

                _json += "],";

                _json += " \"Item\": \"" + v.Item + "\",\n";
                _json += " \"Task\": \"" + v.Task + "\",\n";
                _json += " \"Class\": \"" + v.Class + "\",\n";
                _json += " \"Name\": \"" + v.Name + "\",\n";
                _json += " \"Label\": \"" + v.Label + "\",\n";
                _json += " \"UseForRouting\": " + v.UseForRouting.ToString().ToLower() + ",\n"; 
                _json += " \"Type\": \"" + v.Type + "\"\n";

                if (j != _categoricalVariableKeyList.Count - 1)
                    _json += "},";
                else
                    _json += "}\n";

            }

            _json += "]\n";
            _json += "\n}";


            // Write to File

            File.WriteAllText(JSONFile, _json);

        }

        public static string RemoveUmlaute(string value)
        {
            return value.Replace("ä", "ae").Replace("ö", "oe").Replace("ü", "ue").Replace("Ä", "Ae").Replace("Ö", "Oe").Replace("Ü", "Ue").Replace("ß", "ss");
        }
    }

    public class Variable
    {
        public string Item { get; set; }
        public string Class { get; set; }
        public string Task { get; set; }
        public string Name { get; set; }
        public string Label { get; set; }
        public string Type { get; set; }
        public bool UseForRouting { get; set; }

        public Dictionary<string, HitMiss> Values { get; set; }
    }

    public class HitMiss
    {
        public string Hit { get; set; }
        public string Value { get; set; }
        public string Label { get; set; }
        public string ResultText { get; set; }
    }

     

}
