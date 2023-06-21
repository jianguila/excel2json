using System;
using CommandLine;
using CommandLine.Text;

namespace excel2json
{
    partial class Program
    {
        /// <summary>
        /// 命令行参数定义
        /// </summary>
        internal sealed class Options
        {
            public Options()
            {
                this.HeaderRows = 3;
                this.Encoding = "utf8-nobom";
                this.Lowcase = false;
                this.ExportArray = false;
                this.ForceSheetName = false;
            }

            [Option("excel", Required = true, HelpText = "input excel file path.")]
            public string ExcelPath {
                get;
                set;
            }

            [Option("json", Required = false, HelpText = "export json file path.")]
            public string JsonPath {
                get;
                set;
            }

            [Option("csharp", Required = false, HelpText = "export C# data struct code file path.")]
            public string CSharpPath {
                get;
                set;
            }

            [Option("header", Required = false, DefaultValue = 1, HelpText = "number lines in sheet as header.")]
            public int HeaderRows {
                get;
                set;
            }

            [Option("encoding", Required = false, DefaultValue = "utf8-nobom", HelpText = "export file encoding.")]
            public string Encoding {
                get;
                set;
            }

            [Option("lowcase", Required = false, DefaultValue = false, HelpText = "convert filed name to lowcase.")]
            public bool Lowcase {
                get;
                set;
            }

            [Option("array", Required = false, DefaultValue = false, HelpText = "export as array, otherwise as dict object.")]
            public bool ExportArray {
                get;
                set;
            }

            [Option("date", Required = false, DefaultValue = "yyyy/MM/dd", HelpText = "Date Format String, example: dd / MM / yyy hh: mm:ss.")]
            public string DateFormat {
                get;
                set;
            }

            [Option("sheet", Required = false, DefaultValue = false, HelpText = "export with sheet name, even there's only one sheet.")]
            public bool ForceSheetName {
                get;
                set;
            }

            [Option("exclude_prefix", Required = false, DefaultValue = "", HelpText = "exclude sheet or column start with specified prefix.")]
            public string ExcludePrefix {
                get;
                set;
            }

            [Option("cell_json", Required = false, DefaultValue = true, HelpText = "convert json string in cell")]
            public bool CellJson {
                get;
                set;
            }

            [Option("all_string", Required = false, DefaultValue = false, HelpText = "all string")]
            public bool AllString
            {
                get;
                set;
            }
        }
    }
}
