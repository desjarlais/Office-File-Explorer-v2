namespace Office_File_Explorer.Helpers
{
    public class LocalZipFileHeader
    {
        public string Version { get; set; }
        public string GeneralPurposeBitFlag { get; set; }
        public string CompressionMethod { get; set; }
        public string LastModifiedTime { get; set; }
        public string LastModifiedDate { get; set; }
        public string CRC32 { get; set; }
        public string CompressedSize { get; set; }
        public string UncompressedSize { get; set; }
        public string FileNameLength { get; set; }
        public string ExtraFieldLength { get; set; }
        public string FileName { get; set; }
    }
}
