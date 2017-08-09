using docConverter;
using Neevia;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;


namespace mergeConvertedFolders
{
    /// <summary>
    /// The main Merger class.
    /// Contains all methods for performing mergers on multiple pdf documents using DC Pro's API.
    /// </summary>
    class Merger
    {
        private string connString;  //specified in app.config
        private string toConvertDir;  //specified in app.config, where folders are dropped to be converted to pdf
        private string toMergeDir;  //specified in app.config, where converted folders are to be merged into one pdf file
        private string finalDestinationDir;  //specified in app.config, the final \OUT folder where all pdfs go
        private string tempMergeDir; //specified in app.config, folder to temporarily hold merged files until it is verified that the merge process is finished
        private string itemID;
        private DirectoryInfo[] foldersToConvert;  //this information is only for checking the contents of toConvertDir
        private DirectoryInfo[] foldersToMerge;
        private FileInfo[] mergedFiles;
        private static NpgsqlConnection conn;


        public Merger()
        {
            Setup();
        }

        /// <summary>
        /// Assigns variables using app.config settings.
        /// </summary>
        private void Setup()
        { 
            try
            {
                connString = ConfigurationManager.AppSettings["connString"];
                toConvertDir = ConfigurationManager.AppSettings["toConvertDir"];
                toMergeDir = ConfigurationManager.AppSettings["toMergeDir"];
                finalDestinationDir = ConfigurationManager.AppSettings["finalDestinationDir"];
                tempMergeDir = ConfigurationManager.AppSettings["tempMergeDir"];
                foldersToConvert = new DirectoryInfo(toConvertDir).GetDirectories("*.*");
                foldersToMerge = new DirectoryInfo(toMergeDir).GetDirectories("*.*");
                conn = new NpgsqlConnection(connString);
                LoadMergedFolder();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error accessing app.config: " + e.Message);
            }
        }

        /// <summary>
        /// Makes the API call to DC Pro and performs the merger.
        /// </summary>
        /// <param name="folder">The folder with subcontents to merge.</param>
        /// <param name="mergedFileName">The destination name of the merged pdf.</param>
        /// <returns>0 on error, 1 on success.</returns>
        private int MergeFolder(DirectoryInfo folder, string mergedFileName)
        {
            Neevia.docConverter DC = new Neevia.docConverter();
            string mergeDestination = tempMergeDir + mergedFileName;
            string filesToMerge = GetDelimitedFilesNames(folder, "+");
            int Res = DC.mergeMultiplePDF(filesToMerge, mergeDestination);

            return Res;
        }

        /// <summary>
        /// Checks if all files in a given folder have successfully been converted by DC Pro so we can begin merging.
        /// </summary>
        /// <param name="folder">The folder to check.</param>
        /// <returns>True if all files are accounted for.</returns>
        private bool SafeToMerge(DirectoryInfo folder)
        {
            bool safeToMerge;
            string fileCount = folder.GetFiles().Length.ToString();

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                string itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
                com.CommandText = String.Format("SELECT num_files FROM redacted.items WHERE id={0};", itemID);
                string numFilesOnDb = com.ExecuteScalar().ToString();
                com.CommandText = String.Format("SELECT converter_error FROM redacted.items WHERE id={0};", itemID);
                string converterError = com.ExecuteScalar().ToString();
                conn.Close();
                if (numFilesOnDb == fileCount && converterError == "False")
                {
                    safeToMerge = true;
                }
                else
                {
                    safeToMerge = false;
                }

                return safeToMerge;
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not complete the database command: " + e.Message);
                return false;
            }
        }

        /// <summary>
        /// Generates a string of delimited file names for DC Pro to use in the merge process.
        /// </summary>
        /// <param name="folder">The folder to merge.</param>
        /// <param name="delimiter">Set to "+" by default.</param>
        /// <returns>The delimited filename.</returns>
        private string GetDelimitedFilesNames(DirectoryInfo folder, string delimiter)
        {
            string result = "";
            foreach (FileInfo f in folder.GetFiles())
            {
                if (!string.IsNullOrEmpty(result))
                {
                    result += delimiter;
                }
                result += f.FullName;
            }
            return result;
        }


        private void LoadMergedFolder()
        {
            mergedFiles = new DirectoryInfo(tempMergeDir).GetFiles("*.*");
        }

        /// <summary>
        /// Formats the folder-to-merge name as the final name of the merged pdf.
        /// </summary>
        /// <param name="folder">The folder to merge.</param>
        /// <returns>Formatted name for the final merged pdf file.</returns>
        private string SetMergedFileName(DirectoryInfo folder)
        {
            //returns mergedFileName as all characters after the last occurrence of "m," assuming that the folder name format is itemXXXXXXX
            string mergedFileName = folder.Name.Split('m').Last() + @".pdf";
            return mergedFileName;
        }

        /// <summary>
        /// Deletes the leftover folder after the merged file has been created.
        /// </summary>
        /// <param name="folder">The folder to merge.</param>
        /// <param name="mergedFileName">The name of the merged file.</param>
        private void DeleteLeftoverFolder(DirectoryInfo folder, string mergedFileName)
        {
            while (!File.Exists(tempMergeDir + mergedFileName))
            {
                Thread.Sleep(1000);
            }
            Directory.Delete(folder.FullName, true);
        }


        private static bool SpecialBackslashes()
        {
            IDbCommand com = conn.CreateCommand();
            com.CommandText = "show standard_conforming_strings;";
            string standard_strings = com.ExecuteScalar().ToString();

            return (standard_strings == "off");
        }


        private static string SqlEscape(string query)
        {
            query = query.Replace("'", "''");
            if (SpecialBackslashes())
                query = query.Replace(@"\", @"\\");
            return query;
        }

        /// <summary>
        /// Updates the database according to the results of the merger.
        /// </summary>
        /// <param name="folder">The folder to merge.</param>
        /// <param name="status">True on success, false on failure.</param>
        private void UpdateDb(DirectoryInfo folder, bool status)
        {
            itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                com.CommandText = String.Format("UPDATE redacted.items SET converter_error=false, converter_errormsg='' WHERE id = {0};", itemID);
                com.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
            }

        }

        /// <summary>
        /// Runs the program in logical order. 
        /// Updates the database according to errors.
        /// Finally, moves folders to the final output directory if they have succeeded.
        /// </summary>
        public void Run()
        {
            List<DirectoryInfo> failedMergers = new List<DirectoryInfo>();
            foreach (DirectoryInfo folder in foldersToMerge)
            {
                if (SafeToMerge(folder))
                {
                    int Res = MergeFolder(folder, SetMergedFileName(folder));

                    if (Res == 0)
                    {
                        UpdateDb(folder, true);
                        DeleteLeftoverFolder(folder, SetMergedFileName(folder));
                    }
                    else
                    {
                        failedMergers.Add(folder);
                        WriteOut.HandleMessage("Merger failed on " + folder.FullName);
                        continue;
                    }
                }
            }

            LoadMergedFolder();

            foreach (FileInfo file in mergedFiles)
            {
                try
                {
                    if(File.Exists(finalDestinationDir + file.Name))
                    {
                        File.Delete(finalDestinationDir + file.Name);
                    }
                    File.Move(file.FullName, finalDestinationDir + file.Name);
                }
                catch (Exception e)
                {
                    WriteOut.HandleMessage("Error moving merged file " + file.Name + " into OUT: " + e.Message);
                    continue;
                }
            }
            
            if (failedMergers.Count != 0)
            {
                ErrorHandler theErrorHandler = new ErrorHandler();
                theErrorHandler.Report(failedMergers);
            }
        }
    }
}
