using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Spsl.ExcelUpdater
{
    class ProgramOptions
    {
        [Option('f', "FileName", DefaultValue = "PowerQueryNpgSql.xlsx", HelpText = "Name of file which will be updated")]//, Required = true)]
        public string FileName { get; set; }

    }
}
