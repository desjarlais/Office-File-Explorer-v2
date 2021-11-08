namespace Office_File_Explorer.Helpers
{
    public class NotesSlides
    {
        public T2dHeader t2dHeader;
        public T2dDate t2dDate;
        public T2dSlideNumber t2dSlideNumber;
        public T2dSlideImage t2dSlideImage;
        public T2dPicture t2dPicture;
        public T2dFooter t2dFooter;
        public T2dNotes t2dNotes;
        public PresNotesSz pNotesSz;
    }

    public struct PresNotesSz
    {
        public long Cx;
        public long Cy;
    }

    public struct T2dHeader
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dDate
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dSlideNumber
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dPicture
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dFooter
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dNotes
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }

    public struct T2dSlideImage
    {
        public long OffsetX;
        public long OffsetY;
        public long ExtentsCx;
        public long ExtentsCy;
    }
}
