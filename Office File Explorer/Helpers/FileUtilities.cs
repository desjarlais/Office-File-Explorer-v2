using System;
using System.IO;
using System.Security.Cryptography;
using static Office_File_Explorer.FrmMain;

namespace Office_File_Explorer.Helpers
{
    class FileUtilities
    {
        /// <summary>
        /// use the file extension to get the file type
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static OpenXmlInnerFileTypes GetFileType(string path)
        {
            switch (Path.GetExtension(path))
            {
                case ".docx":
                case ".dotx":
                case ".dotm":
                case ".docm":
                    return OpenXmlInnerFileTypes.Word;
                case ".xlsx":
                case ".xlsm":
                case ".xltm":
                case ".xltx":
                case ".xlsb":
                    return OpenXmlInnerFileTypes.Excel;
                case ".pptx":
                case ".pptm":
                case ".ppsx":
                case ".ppsm":
                case ".potx":
                case ".potm":
                    return OpenXmlInnerFileTypes.PowerPoint;
                case ".msg":
                    return OpenXmlInnerFileTypes.Outlook;
                case ".doc":
                case ".dot":
                case ".xls":
                case ".xlt":
                case ".ppt":
                case ".pot":
                    return OpenXmlInnerFileTypes.CompoundFile;
                case ".jpeg":
                case ".jpg":
                case ".bmp":
                case ".png":
                case ".gif":
                case ".emf":
                case ".wmf":
                    return OpenXmlInnerFileTypes.Image;
                case ".xml":
                case ".vml":
                case ".rels":
                    return OpenXmlInnerFileTypes.XML;
                case ".mp4":
                case ".avi":
                case ".wmv":
                case ".mov":
                    return OpenXmlInnerFileTypes.Video;
                case ".mp3":
                case ".wav":
                case ".wma":
                    return OpenXmlInnerFileTypes.Audio;
                case ".txt":
                    return OpenXmlInnerFileTypes.Text;
                case ".bin":
                case ".sigs":
                case ".odttf":
                    return OpenXmlInnerFileTypes.Binary;
                default:
                    return OpenXmlInnerFileTypes.Other;
            }
        }

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

        /// <summary>
        /// this function is used to load binary files into the part viewer
        /// pp.GetStream().CopyTo(MemoryStream) will throw errors when trying to open multiple times
        /// this function will get around that problem
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static byte[] ReadToEnd(Stream stream)
        {
            long originalPosition = 0;

            if (stream.CanSeek)
            {
                originalPosition = stream.Position;
                stream.Position = 0;
            }

            try
            {
                byte[] readBuffer = new byte[4096];

                int totalBytesRead = 0;
                int bytesRead;

                while ((bytesRead = stream.Read(readBuffer, totalBytesRead, readBuffer.Length - totalBytesRead)) > 0)
                {
                    totalBytesRead += bytesRead;

                    if (totalBytesRead == readBuffer.Length)
                    {
                        int nextByte = stream.ReadByte();
                        if (nextByte != -1)
                        {
                            byte[] temp = new byte[readBuffer.Length * 2];
                            Buffer.BlockCopy(readBuffer, 0, temp, 0, readBuffer.Length);
                            Buffer.SetByte(temp, totalBytesRead, (byte)nextByte);
                            readBuffer = temp;
                            totalBytesRead++;
                        }
                    }
                }

                byte[] buffer = readBuffer;
                if (readBuffer.Length != totalBytesRead)
                {
                    buffer = new byte[totalBytesRead];
                    Buffer.BlockCopy(readBuffer, 0, buffer, 0, totalBytesRead);
                }
                return buffer;
            }
            finally
            {
                if (stream.CanSeek)
                {
                    stream.Position = originalPosition;
                }
            }
        }

        /// <summary>
        /// check if file is encrypted or compound file format
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsFileEncrypted(string filePath)
        {
            byte[] buffer = new byte[8];
            try
            {
                // open the file and populate the first 8 bytes into the buffer
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.ReadExactly(buffer);
                }

                // encrypted files from Office start with D0 CF 11 E0 A1 B1 1A E1 (compound file format header)
                // hex = 208 207 17 224 161 177 26 225
                if (buffer[0].ToString() == "208" && buffer[1].ToString() == "207" && buffer[2].ToString() == "17" && buffer[3].ToString() == "224" && 
                    buffer[4].ToString() == "161" && buffer[5].ToString() == "177" && buffer[6].ToString() == "26" && buffer[7].ToString() == "225")
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

        public static bool IsZipArchiveFile(string filePath)
        {
            byte[] buffer = new byte[2];
            try
            {
                // open the file and populate the first 2 bytes into the buffer
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.ReadExactly(buffer);
                }

                // if the buffer starts with PK (hex = 80 75) the file is a zip archive
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
