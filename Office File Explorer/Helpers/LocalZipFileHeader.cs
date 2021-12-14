namespace Office_File_Explorer.Helpers
{
    public class LocalZipFileHeader
    {
        public int Version { get; set; }
        public int GeneralPurposeBitFlag { get; set; }
        public int CompressionMethod { get; set; }
        public int LastModifiedTime { get; set; }
        public int LastModifiedDate { get; set; }
        public int CRC32 { get; set; }
        public int CompressedSize { get; set; }
        public int UncompressedSize { get; set; }
        public int FileNameLength { get; set; }
        public int ExtraFieldLength { get; set; }
        public string FileName { get; set; }
    }
}
