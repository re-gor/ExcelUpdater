﻿using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelUpdater
{
    class ProgramOptions
    {
        [Option('s', "SiteUrl",  HelpText = "Site where excel files located", Required = true)]
        public string SiteUrl { get; set; }

        [Option('l', "LibraryName",  HelpText = "Library where excel files located", Required = true)]
        public string LibraryName { get; set; }

        [Option('f', "FolderName", HelpText = "Folder in library where excel files located")]//, Required = true)]
        public string SubFolder { get; set; }

        [Option('v', "ExcelVisible", DefaultValue = false, HelpText = "Show excel instance during updating of files (usefull for configuring connections of PowerQuery)")]
        public bool ExcelVisible { get; set; }
    }
}
