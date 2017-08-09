using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;

namespace mergeConvertedFolders
{

    class Program
    {
        static void Main(string[] args)
        {
            Merger theMerger = new Merger();  
            theMerger.Run();  //runs the mergers and returns list of failed mergers

            ErrorHandler theErrorHandler = new ErrorHandler();
            theErrorHandler.ReportStagnantFolders();

            // program will run only once if looperKey != "true"
            string looperPath = ConfigurationManager.AppSettings["looperPath"];
            string looperKey;
            /*NOTE:
             *Var looperPath holds the path to a text document. Var looperKey holds the contents of the text document.
             *The while loop continues indefinitely as long as the file contains the specified string "true". Anything else breaks the loop.
             * String is case insensitive and whitespaces have no effect.
             */
            while (string.Equals(looperKey = File.ReadAllText(looperPath).Trim(), "true", StringComparison.OrdinalIgnoreCase))
            {
                theMerger = new Merger();
                theMerger.Run();
                theErrorHandler = new ErrorHandler();
                theErrorHandler.ReportStagnantFolders();
                Thread.Sleep(1000);
                theErrorHandler.RemoveBrokenFolders();
            }
        }
    }

}
