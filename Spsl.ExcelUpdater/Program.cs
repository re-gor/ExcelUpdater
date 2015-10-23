using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spsl.ExcelUpdater
{
    class Program
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));

        static void Main(string[] args)
        {
            var result = CommandLine.Parser.Default.ParseArguments<ProgramOptions>(args);
            if (result.Errors.Any())
            {
                foreach(var error in result.Errors)
                {
                    _log.ErrorFormat("Can not start program: {0}", error.Tag);
                }
                return;
            }

            Console.WriteLine("Working file name: {0}", result.Value.FileName);

            try
            {
                ExcelHandler.ExternalDataUpdater.UpdateSharepointFile(result.Value.FileName);
            }
            catch (Exception ex)
            {
                _log.Error(ex);
                return;
            }
        }
    }
}
