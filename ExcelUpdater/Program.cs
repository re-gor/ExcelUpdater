using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUpdater
{
    class Program
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));

        static void Main(string[] args)
        {
            var result = CommandLine.Parser.Default.ParseArguments<ProgramOptions>(args);
            if (result.Errors.Any())
            {
                if (result.Errors.First().Tag == CommandLine.ErrorType.HelpRequestedError)
                    return;

                foreach (var error in result.Errors)
                {
                    _log.ErrorFormat("Can not start program: {0}", error.Tag);
                }
                return;
            }

            try
            {
                ExcelHandler.ExternalDataUpdater.UpdateSharepointFiles(result.Value.SiteUrl, result.Value.LibraryName, result.Value.SubFolder);
            }
            catch (Exception ex)
            {
                _log.Error(ex);
                return;
            }
        }
    }
}
