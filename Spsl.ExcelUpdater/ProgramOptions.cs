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
        //@"D:\YandexDisk\Работа\Spsl.ExcelUpdater\Spsl.ExcelUpdater\PowerQueryNpgSql.xlsx"
        [Option('f', "FileName", DefaultValue = @"https://portal.spsl.sbras.ru/to/Shared documents/Тест/FooTable.xlsx", HelpText = "Name of file which will be updated")]//, Required = true)]
        public string FileName { get; set; }

    }
}
