using docConverter;
using Neevia;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;

namespace mergeConvertedFolders
{
    /// <summary>
    /// Handles folders that failed to merge.
    /// </summary>
    class ErrorHandler
    {
        private static string connString;  //app.config
        private string emailHost;
        private string errorDir;  //specified in app.config
        private string fromAddress;  //for SendEmail()
        private string toAddress;  //for SendEmail()
        private string toMergeDir;  //specified in app.config
        private string itemID;
        private DirectoryInfo[] errorFolders;
        private DirectoryInfo[] foldersToMerge;
        private List<DirectoryInfo> emailReportList;
        private static NpgsqlConnection conn;


        public ErrorHandler()
        {
            Setup();
        }

        /// <summary>
        /// Assigns variables based on app.config settings.
        /// </summary>
        private void Setup()
        {
            try
            {
                connString = ConfigurationManager.AppSettings["connString"];
                emailHost = ConfigurationManager.AppSettings["emailHost"];
                emailReportList = new List<DirectoryInfo>();
                errorDir = ConfigurationManager.AppSettings["errorDir"];
                errorFolders = new DirectoryInfo(errorDir).GetDirectories("*.*");
                fromAddress = ConfigurationManager.AppSettings["fromAddress"];
                toAddress = ConfigurationManager.AppSettings["toAddress"];
                toMergeDir = ConfigurationManager.AppSettings["toMergeDir"];
                foldersToMerge = new DirectoryInfo(toMergeDir).GetDirectories("*.*");
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Error accessing app.config: " + e.Message);
            }
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
        /// Checks if a folder has been stagnating in the "to-merge" directory.
        /// </summary>
        /// <param name="folder">The folder to check.</param>
        /// <returns>True if the folder is stagnant.</returns>
        private bool StagnantFolder(DirectoryInfo folder)  //returns true if folder is over x mins old
        {
            DateTime creationTime = folder.CreationTime;
            DateTime currentTime = DateTime.Now.AddMinutes(-30);  //subtract x mins from current time
            if (currentTime > creationTime)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if the folder has a flagged error.
        /// </summary>
        /// <param name="folder">The folder to check.</param>
        /// <returns>True if converter_erro=true</returns>
        private bool IsError(DirectoryInfo folder)
        {
            bool isError;
            conn = new NpgsqlConnection(connString);
            conn.Open();
            NpgsqlCommand com = conn.CreateCommand();
            itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
            com.CommandText = String.Format("SELECT converter_error FROM redacted.items WHERE id={0};", itemID);
            string response = com.ExecuteScalar().ToString();
            conn.Close();
            if (response == "True")
            {
                isError = true;
            }
            else
            {
                isError = false;
            }
            return isError;
        }


        /// <summary>
        /// Moves the folder to the error directory.
        /// </summary>
        /// <param name="folder">The folder to move.</param>
        private void MoveFolderToError(DirectoryInfo folder)
        {
            try
            {
                folder.MoveTo(errorDir + folder.Name);
            }
            catch
            {
                try
                {
                    Directory.Delete(errorDir + folder.Name, true);
                    folder.MoveTo(errorDir + folder.Name);
                }
                catch (Exception e)
                {
                    WriteOut.HandleMessage("Error moving folder to ERROR: " + e.Message);
                }
            }
        }


        /// <summary>
        /// Updates the database and reports the error.
        /// </summary>
        /// <param name="folder">Folder to report.</param>
        private void UpdateDb(DirectoryInfo folder)
        {
            itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
            string msg = "Failed to merge:" + " (" + errorDir + folder.Name + ")";

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                com.CommandText = String.Format("UPDATE redacted.items SET converted=false, converter_error=true, converter_errormsg='{1}' WHERE id = {0};", itemID, SqlEscape(msg));
                com.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
            }
        }


        /// <summary>
        /// Updates the database and reports the error.
        /// </summary>
        /// <param name="folder">Folder to report.</param>
        /// <param name="msg">Error message.</param>
        private void UpdateDb(DirectoryInfo folder, string msg)
        {
            itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX

            try
            {
                conn = new NpgsqlConnection(connString);
                conn.Open();
                NpgsqlCommand com = conn.CreateCommand();
                com.CommandText = String.Format("UPDATE redacted.items SET converted=false, converter_error=true, converter_errormsg='{1}' WHERE id = {0};", itemID, SqlEscape(msg));
                com.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
            }
        }

        /// <summary>
        /// Generates the error report for the email body.
        /// </summary>
        /// <param name="failedMergers">Failures to report.</param>
        /// <returns>The email report.</returns>
        private string GenerateErrorMessage(List<DirectoryInfo> failedMergers)
        {
            string emailBody = "Errors occurred while attempting to merge the following folders: \r";
            foreach (DirectoryInfo folder in failedMergers)
            {
                emailBody += "\r\n";
                emailBody += failedMergers.IndexOf(folder).ToString() + ") ";
                emailBody += folder.FullName;

                try
                {
                    itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX
                    conn = new NpgsqlConnection(connString);
                    conn.Open();
                    NpgsqlCommand com = conn.CreateCommand();
                    com.CommandText = String.Format("SELECT converter_errormsg FROM redacted.items WHERE id={0};", itemID);
                    string pgResponse = com.ExecuteScalar().ToString();
                    conn.Close();
                    emailBody += " | " + pgResponse;
                }
                catch (Exception e)
                {
                    WriteOut.HandleMessage("Error connecting to database: " + e.Message);
                }
            }
            return emailBody;
        }

        /// <summary>
        /// Sends the email using smtp.
        /// </summary>
        /// <param name="emailBody">The message to send.</param>
        private void SendEmail(string emailBody)
        {
            string emailSubject = "PDF Merger Error: " + DateTime.Now.ToString();

            try
            {
                SmtpClient myClient = new SmtpClient(emailHost);
                MailMessage myMessage = new MailMessage(fromAddress, toAddress, emailSubject, emailBody);
                myClient.Send(myMessage);
            }
            catch (Exception e)
            {
                WriteOut.HandleMessage("Error while attempting to send mail: " + e.Message);
            }
        }


        /// <summary>
        /// Removes folders in the toMergeDir if parts of the folder items exist in error.
        /// </summary>
        public void RemoveBrokenFolders()
        {
            foreach (DirectoryInfo folder in foldersToMerge)
            {
                foreach (DirectoryInfo f in errorFolders)
                {
                    if (folder.Name == f.Name)
                    {
                        UpdateDb(folder);
                        MoveFolderToError(folder);
                    }
                }
            }
        }



        /// <summary>
        /// Reports failures taken directly from the list of failed mergers after the Merge class runs
        /// </summary>
        /// <param name="failedMergers">Failed mergers to report.</param>
        public void ReportFailedMergers(List<DirectoryInfo> failedMergers)
        {
            foreach (DirectoryInfo folder in failedMergers)
            {
                itemID = folder.Name.Split('m').Last();  //assumes folder name is format itemXXXXXXXX

                try
                {
                    conn = new NpgsqlConnection(connString);
                    conn.Open();
                    NpgsqlCommand com = conn.CreateCommand();
                    string msg = "Failed to merge:" + " (" + errorDir + folder.Name + ")";
                    com.CommandText = String.Format("UPDATE redacted.items SET converted=false, converter_error=true, converter_errormsg='{1}' WHERE id={0};", itemID, SqlEscape(msg));
                    com.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception e)
                {
                    WriteOut.HandleMessage("Could not connect to the database: " + e.Message);
                }

                emailReportList.Add(folder);
                MoveFolderToError(folder);
            }
        }

        /// <summary>
        /// Checks one folder at a time if it is stagnant and moves those stagnant folders to the error directory.
        /// Generates a list of folders to email error report.
        /// </summary>
        public void ReportStagnantFolders()
        {
            foreach (DirectoryInfo folder in foldersToMerge)
            {
                if (StagnantFolder(folder))
                {
                    string msg = "Failed to merge. Stagnant folder:" + " (" + errorDir + folder.Name + ")"; ;
                    emailReportList.Add(folder);
                    UpdateDb(folder, msg);
                    MoveFolderToError(folder);
                }
            }
        }


        /// <summary>
        /// Called by the merger class when a merge fails
        /// </summary>
        /// <param name="failedMergers">The failed mergers to report</param>
        public void Report(List<DirectoryInfo> failedMergers)
        {
            RemoveBrokenFolders();
            ReportStagnantFolders();
            ReportFailedMergers(failedMergers);

            if (emailReportList.Count != 0)
            {
                SendEmail(GenerateErrorMessage(emailReportList));
            }
        }
    }
}
