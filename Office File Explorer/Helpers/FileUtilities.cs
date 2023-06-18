using System;
using System.IO;
using System.Security.Cryptography;

namespace Office_File_Explorer.Helpers
{
    class FileUtilities
    {
        static readonly string[] sizeSuffixes = { "bytes", "KB", "MB", "GB" };

        /// <summary>
        /// this function takes a file size in bytes and converts it to the equivalent file size label
        /// </summary>
        /// <param name="value">the size in bytes of the attached file being added</param>
        /// <returns></returns>
        public static string SizeSuffix(long value)
        {
            if (value < 0)
            {
                return "-" + SizeSuffix(-value);
            }
            if (value == 0)
            {
                return "0.0 bytes";
            }

            int mag = (int)Math.Log(value, 1024);
            decimal adjustedSize = (decimal)value / (1L << (mag * 10));

            return string.Format("{0:n1} {1}", adjustedSize, sizeSuffixes[mag]);
        }

        public static bool IsZipArchiveFile(string filePath)
        {
            byte[] buffer = new byte[2];
            try
            {
                // open the file and populate the first 2 bytes into the buffer
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.Read(buffer, 0, buffer.Length);
                }

                // if the buffer starts with PK the file is a zip archive
                if (buffer[0].ToString() == "80" && buffer[1].ToString() == "75")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                WriteToLog(Strings.fLogFilePath, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// used for random number where predictability is not important
        /// </summary>
        /// <returns></returns>
        public static int GetRandomNumber()
        {
            Random r = new Random();
            int rNum = r.Next(1, 1000);
            return rNum;
        }

        /// <summary>
        /// used for scenarios where preditability is important
        /// </summary>
        /// <returns></returns>
        public static int GetRandomCryptoNumber()
        {
            var min = 1;
            var max = 1000;
            return RandomNumberGenerator.GetInt32(min, max);
        }

        public static string ConvertUriToFilePath(string path)
        {
            if (path.StartsWith("http"))
            {
                return path;
            }

            var filePath = new Uri(path).LocalPath;
            return filePath;
        }

        public static string ConvertFilePathToUri(string path)
        {
            var uri = new Uri(path);
            var convertedUri = uri.AbsoluteUri;
            return convertedUri;
        }

        /// <summary>
        /// given a broken uri this function will return a generic non-broken uri
        /// </summary>
        /// <param name="brokenUri">the uri that is failing in the sdk</param>
        /// <returns></returns>
        public static Uri FixUri(string brokenUri)
        {
            brokenUri = "http://broken-link/";
            return new Uri(brokenUri);
        }

        public static void WriteToLog(string fPath, string sOutput)
        {
            if (!File.Exists(fPath))
            {
                File.Create(fPath).Close();
            }

            using (StreamWriter sw = File.AppendText(fPath))
            {
                sw.WriteLine(DateTime.Now + Strings.wColon + sOutput);
            }
        }

        public static string GetAppFromFileExtension(string fPath)
        {
            if (fPath.EndsWith(Strings.docxFileExt) || fPath.EndsWith(Strings.dotxFileExt) || fPath.EndsWith(Strings.docmFileExt) || fPath.EndsWith(Strings.dotmFileExt))
            {
                return Strings.oAppWord;
            }
            else if (fPath.EndsWith(Strings.xlsxFileExt) || fPath.EndsWith(Strings.xltxFileExt) || fPath.EndsWith(Strings.xlsmFileExt) || fPath.EndsWith(Strings.xltmFileExt))
            {
                return Strings.oAppExcel;
            }
            else if (fPath.EndsWith(Strings.pptxFileExt) || fPath.EndsWith(Strings.potxFileExt) || fPath.EndsWith(Strings.pptmFileExt) || fPath.EndsWith(Strings.potmFileExt))
            {
                return Strings.oAppPowerPoint;
            }
            else
            {
                return Strings.oAppUnknown;
            }
        }
    }
}
