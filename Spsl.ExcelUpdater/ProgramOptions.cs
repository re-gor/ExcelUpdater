﻿using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Spsl.ExcelUpdater
{
    class ProgramOptions
    {
        [Option('s', "SiteUrl", DefaultValue = @"https://portal.spsl.sbras.ru/to/", HelpText = "Site where excel files located")]//, Required = true)]
        public string SiteUrl { get; set; }

        [Option('l', "LibraryName", DefaultValue = @"Документы", HelpText = "Library where excel files located")]//, Required = true)]
        public string LibraryName { get; set; }

        [Option('f', "FolderName", DefaultValue = null, HelpText = "Folder in library where excel files located")]//, Required = true)]
        public string SubFolder { get; set; }
    }
}
