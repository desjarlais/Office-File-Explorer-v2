using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using A = DocumentFormat.OpenXml.Drawing;
using PShape = DocumentFormat.OpenXml.Presentation.Shape;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using ShapeStyle = DocumentFormat.OpenXml.Presentation.ShapeStyle;
using ModernComment = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using Comment = DocumentFormat.OpenXml.Presentation.Comment;

namespace Office_File_Explorer.Helpers
{
    class PowerPoint
    {
        public static bool fSuccess;

        public static TextStyles GenerateDefaultTextStyles()
        {
            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing1.Append(spacingPercent1);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent() { Val = 0 };

            spaceBefore1.Append(spacingPercent2);
            A.NoBullet noBullet1 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 4400, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mj-cs" };

            defaultRunProperties1.Append(solidFill1);
            defaultRunProperties1.Append(latinFont1);
            defaultRunProperties1.Append(eastAsianFont1);
            defaultRunProperties1.Append(complexScriptFont1);

            level1ParagraphProperties1.Append(lineSpacing1);
            level1ParagraphProperties1.Append(spaceBefore1);
            level1ParagraphProperties1.Append(noBullet1);
            level1ParagraphProperties1.Append(defaultRunProperties1);

            titleStyle1.Append(level1ParagraphProperties1);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties() { LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing2.Append(spacingPercent3);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints() { Val = 1000 };

            spaceBefore2.Append(spacingPoints1);
            A.BulletFont bulletFont1 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 2800, Kerning = 1200 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill2);
            defaultRunProperties2.Append(latinFont2);
            defaultRunProperties2.Append(eastAsianFont2);
            defaultRunProperties2.Append(complexScriptFont2);

            level1ParagraphProperties2.Append(lineSpacing2);
            level1ParagraphProperties2.Append(spaceBefore2);
            level1ParagraphProperties2.Append(bulletFont1);
            level1ParagraphProperties2.Append(characterBullet1);
            level1ParagraphProperties2.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties() { LeftMargin = 685800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing3.Append(spacingPercent4);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints() { Val = 500 };

            spaceBefore3.Append(spacingPoints2);
            A.BulletFont bulletFont2 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 2400, Kerning = 1200 };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor3);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill3);
            defaultRunProperties3.Append(latinFont3);
            defaultRunProperties3.Append(eastAsianFont3);
            defaultRunProperties3.Append(complexScriptFont3);

            level2ParagraphProperties1.Append(lineSpacing3);
            level2ParagraphProperties1.Append(spaceBefore3);
            level2ParagraphProperties1.Append(bulletFont2);
            level2ParagraphProperties1.Append(characterBullet2);
            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties() { LeftMargin = 1143000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing4.Append(spacingPercent5);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints() { Val = 500 };

            spaceBefore4.Append(spacingPoints3);
            A.BulletFont bulletFont3 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill4.Append(schemeColor4);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill4);
            defaultRunProperties4.Append(latinFont4);
            defaultRunProperties4.Append(eastAsianFont4);
            defaultRunProperties4.Append(complexScriptFont4);

            level3ParagraphProperties1.Append(lineSpacing4);
            level3ParagraphProperties1.Append(spaceBefore4);
            level3ParagraphProperties1.Append(bulletFont3);
            level3ParagraphProperties1.Append(characterBullet3);
            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties() { LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing5.Append(spacingPercent6);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints() { Val = 500 };

            spaceBefore5.Append(spacingPoints4);
            A.BulletFont bulletFont4 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill5.Append(schemeColor5);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill5);
            defaultRunProperties5.Append(latinFont5);
            defaultRunProperties5.Append(eastAsianFont5);
            defaultRunProperties5.Append(complexScriptFont5);

            level4ParagraphProperties1.Append(lineSpacing5);
            level4ParagraphProperties1.Append(spaceBefore5);
            level4ParagraphProperties1.Append(bulletFont4);
            level4ParagraphProperties1.Append(characterBullet4);
            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties() { LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing6.Append(spacingPercent7);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints() { Val = 500 };

            spaceBefore6.Append(spacingPoints5);
            A.BulletFont bulletFont5 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor6);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill6);
            defaultRunProperties6.Append(latinFont6);
            defaultRunProperties6.Append(eastAsianFont6);
            defaultRunProperties6.Append(complexScriptFont6);

            level5ParagraphProperties1.Append(lineSpacing6);
            level5ParagraphProperties1.Append(spaceBefore6);
            level5ParagraphProperties1.Append(bulletFont5);
            level5ParagraphProperties1.Append(characterBullet5);
            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties() { LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing7.Append(spacingPercent8);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints() { Val = 500 };

            spaceBefore7.Append(spacingPoints6);
            A.BulletFont bulletFont6 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill7);
            defaultRunProperties7.Append(latinFont7);
            defaultRunProperties7.Append(eastAsianFont7);
            defaultRunProperties7.Append(complexScriptFont7);

            level6ParagraphProperties1.Append(lineSpacing7);
            level6ParagraphProperties1.Append(spaceBefore7);
            level6ParagraphProperties1.Append(bulletFont6);
            level6ParagraphProperties1.Append(characterBullet6);
            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties() { LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing8.Append(spacingPercent9);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints() { Val = 500 };

            spaceBefore8.Append(spacingPoints7);
            A.BulletFont bulletFont7 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill8);
            defaultRunProperties8.Append(latinFont8);
            defaultRunProperties8.Append(eastAsianFont8);
            defaultRunProperties8.Append(complexScriptFont8);

            level7ParagraphProperties1.Append(lineSpacing8);
            level7ParagraphProperties1.Append(spaceBefore8);
            level7ParagraphProperties1.Append(bulletFont7);
            level7ParagraphProperties1.Append(characterBullet7);
            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties() { LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing9.Append(spacingPercent10);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints() { Val = 500 };

            spaceBefore9.Append(spacingPoints8);
            A.BulletFont bulletFont8 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor9);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill9);
            defaultRunProperties9.Append(latinFont9);
            defaultRunProperties9.Append(eastAsianFont9);
            defaultRunProperties9.Append(complexScriptFont9);

            level8ParagraphProperties1.Append(lineSpacing9);
            level8ParagraphProperties1.Append(spaceBefore9);
            level8ParagraphProperties1.Append(bulletFont8);
            level8ParagraphProperties1.Append(characterBullet8);
            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties() { LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing10.Append(spacingPercent11);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints() { Val = 500 };

            spaceBefore10.Append(spacingPoints9);
            A.BulletFont bulletFont9 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill10.Append(schemeColor10);
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill10);
            defaultRunProperties10.Append(latinFont10);
            defaultRunProperties10.Append(eastAsianFont10);
            defaultRunProperties10.Append(complexScriptFont10);

            level9ParagraphProperties1.Append(lineSpacing10);
            level9ParagraphProperties1.Append(spaceBefore10);
            level9ParagraphProperties1.Append(bulletFont9);
            level9ParagraphProperties1.Append(characterBullet9);
            level9ParagraphProperties1.Append(defaultRunProperties10);

            bodyStyle1.Append(level1ParagraphProperties2);
            bodyStyle1.Append(level2ParagraphProperties1);
            bodyStyle1.Append(level3ParagraphProperties1);
            bodyStyle1.Append(level4ParagraphProperties1);
            bodyStyle1.Append(level5ParagraphProperties1);
            bodyStyle1.Append(level6ParagraphProperties1);
            bodyStyle1.Append(level7ParagraphProperties1);
            bodyStyle1.Append(level8ParagraphProperties1);
            bodyStyle1.Append(level9ParagraphProperties1);

            OtherStyle otherStyle1 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties1.Append(defaultRunProperties11);

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill11.Append(schemeColor11);
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties12.Append(solidFill11);
            defaultRunProperties12.Append(latinFont11);
            defaultRunProperties12.Append(eastAsianFont11);
            defaultRunProperties12.Append(complexScriptFont11);

            level1ParagraphProperties3.Append(defaultRunProperties12);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill12.Append(schemeColor12);
            A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties13.Append(solidFill12);
            defaultRunProperties13.Append(latinFont12);
            defaultRunProperties13.Append(eastAsianFont12);
            defaultRunProperties13.Append(complexScriptFont12);

            level2ParagraphProperties2.Append(defaultRunProperties13);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor13);
            A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties14.Append(solidFill13);
            defaultRunProperties14.Append(latinFont13);
            defaultRunProperties14.Append(eastAsianFont13);
            defaultRunProperties14.Append(complexScriptFont13);

            level3ParagraphProperties2.Append(defaultRunProperties14);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor14);
            A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties15.Append(solidFill14);
            defaultRunProperties15.Append(latinFont14);
            defaultRunProperties15.Append(eastAsianFont14);
            defaultRunProperties15.Append(complexScriptFont14);

            level4ParagraphProperties2.Append(defaultRunProperties15);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor15);
            A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties16.Append(solidFill15);
            defaultRunProperties16.Append(latinFont15);
            defaultRunProperties16.Append(eastAsianFont15);
            defaultRunProperties16.Append(complexScriptFont15);

            level5ParagraphProperties2.Append(defaultRunProperties16);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor16);
            A.LatinFont latinFont16 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont16 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties17.Append(solidFill16);
            defaultRunProperties17.Append(latinFont16);
            defaultRunProperties17.Append(eastAsianFont16);
            defaultRunProperties17.Append(complexScriptFont16);

            level6ParagraphProperties2.Append(defaultRunProperties17);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor17);
            A.LatinFont latinFont17 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont17 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont17 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties18.Append(solidFill17);
            defaultRunProperties18.Append(latinFont17);
            defaultRunProperties18.Append(eastAsianFont17);
            defaultRunProperties18.Append(complexScriptFont17);

            level7ParagraphProperties2.Append(defaultRunProperties18);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor18);
            A.LatinFont latinFont18 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont18 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties19.Append(solidFill18);
            defaultRunProperties19.Append(latinFont18);
            defaultRunProperties19.Append(eastAsianFont18);
            defaultRunProperties19.Append(complexScriptFont18);

            level8ParagraphProperties2.Append(defaultRunProperties19);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor19);
            A.LatinFont latinFont19 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties20.Append(solidFill19);
            defaultRunProperties20.Append(latinFont19);
            defaultRunProperties20.Append(eastAsianFont19);
            defaultRunProperties20.Append(complexScriptFont19);

            level9ParagraphProperties2.Append(defaultRunProperties20);

            otherStyle1.Append(defaultParagraphProperties1);
            otherStyle1.Append(level1ParagraphProperties3);
            otherStyle1.Append(level2ParagraphProperties2);
            otherStyle1.Append(level3ParagraphProperties2);
            otherStyle1.Append(level4ParagraphProperties2);
            otherStyle1.Append(level5ParagraphProperties2);
            otherStyle1.Append(level6ParagraphProperties2);
            otherStyle1.Append(level7ParagraphProperties2);
            otherStyle1.Append(level8ParagraphProperties2);
            otherStyle1.Append(level9ParagraphProperties2);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);
            return textStyles1;
        }

        public static List<string> GetFonts(string fPath)
        {
            List<string> fonts = new List<string>();
            int fCount = 0;

            using (PresentationDocument pptDoc = PresentationDocument.Open(fPath, false))
            {
                // list the embedded fonts
                if (pptDoc.PresentationPart.Presentation.EmbeddedFontList is null)
                {
                    return fonts;
                }

                foreach (EmbeddedFont ef in pptDoc.PresentationPart.Presentation.EmbeddedFontList)
                {
                    fCount++;
                    if (ef.Features.IsReadOnly)
                    {
                        fonts.Add(fCount + Strings.wPeriod + ef.Font.Typeface + " || Character Set = " + AppUtilities.GetFontCharacterSet(ef.Font.CharacterSet) + " (Read-Only)");
                    }
                    else
                    {
                        fonts.Add(fCount + Strings.wPeriod + ef.Font.Typeface + " || Character Set = " + AppUtilities.GetFontCharacterSet(ef.Font.CharacterSet));
                    }
                }
            }

            return fonts;
        }

        public static bool RemoveComments(string path)
        {
            fSuccess = false;

            using (PresentationDocument pptDoc = PresentationDocument.Open(path, true))
            {
                PresentationPart pPart = pptDoc.PresentationPart;

                foreach (SlidePart sPart in pPart.SlideParts)
                {
                    SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                    if (sCPart is null)
                    {
                        return fSuccess;
                    }

                    foreach (Comment cmt in sCPart.CommentList)
                    {
                        cmt.Remove();
                        fSuccess = true;
                    }
                }

                if (fSuccess)
                {
                    pptDoc.PresentationPart.Presentation.Save();
                }
            }

            return fSuccess;
        }

        /// <summary>
        /// Move a slide to a different position in the slide order in the presentation.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="from">slide index # of the source slide</param>
        /// <param name="to">slide index # of the target slide</param>
        public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
        {
            // Get the presentation part from the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // The slide count is not zero, so the presentation must contain slides.
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the source slide.
            SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;
            SlideId targetSlide = slideIdList.ChildElements[to] as SlideId;

            // Remove the source slide from its current position.
            sourceSlide.Remove();

            // Insert the source slide at its new position after the target slide.
            // if the slide being moved is before the target position, use InsertAfter
            // otherwise, we want to use InsertBefore
            if (from < to)
            {
                slideIdList.InsertAfter(sourceSlide, targetSlide);
            }
            else
            {
                slideIdList.InsertBefore(sourceSlide, targetSlide);
            }

            // Save the modified presentation.
            presentation.Save();
        }

        /// <summary>
        /// Change the fill color of a shape, docName must have a filled shape as the first shape on the first slide.
        /// </summary>
        /// <param name="docName">path to the file</param>
        public static void SetPPTShapeColor(string docName)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                if (slide != null)
                {
                    // Get the shape tree that contains the shape to change.
                    ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

                    // Get the first shape in the shape tree.
                    PShape shape = tree.GetFirstChild<PShape>();

                    if (shape != null)
                    {
                        // Get the style of the shape.
                        ShapeStyle style = shape.ShapeStyle;

                        // Get the fill reference.
                        FillReference fillRef = style.FillReference;

                        // Set the fill color to SchemeColor Accent 6;
                        fillRef.SchemeColor = new SchemeColor
                        {
                            Val = SchemeColorValues.Accent6
                        };

                        // Save the modified slide.
                        slide.Slide.Save();
                    }
                }
            }
        }

        /// <summary>
        /// Function to retrieve the number of slides
        /// </summary>
        /// <param name="fileName">path to the file</param>
        /// <param name="includeHidden">default is true, pass false if you don't want hidden slides counted</param>
        /// <returns></returns>
        public static int RetrieveNumberOfSlides(string fPath, bool includeHidden = true)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fPath, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart is not null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.SlideParts.Count();
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden 
                        // will contain the value "0". The Show property may not 
                        // exist, and most likely will not, for non-hidden slides.
                        var slides = presentationPart.SlideParts.Where((s) => (s.Slide is not null) && ((s.Slide.Show is null) ||
                            (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }

        public static List<string> GetHyperlinks(string fPath)
        {
            List<string> tList = new List<string>();
            
            int linkCount = 0;
            foreach (string s in GetAllExternalHyperlinksInPresentation(fPath))
            {
                linkCount++;
                tList.Add(linkCount + Strings.wPeriod + s);
            }

            return tList;
        }

        /// <summary>
        /// check for both legacy and modern comments
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> GetComments(string fPath)
        {
            List<string> tList = new List<string>();

            PresentationDocument presentationDocument = PresentationDocument.Open(fPath, false);
            PresentationPart pPart = presentationDocument.PresentationPart;
            int commentCount = 0;

            foreach (SlidePart sPart in pPart.SlideParts)
            {
                // legacy comments
                SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                if (sCPart is not null)
                {
                    foreach (Comment cmt in sCPart.CommentList)
                    {
                        commentCount++;
                        tList.Add(commentCount + Strings.wPeriod + cmt.InnerText);
                    }
                }

                // modern comments
                if (sPart.commentParts is not null)
                {
                    IEnumerable<PowerPointCommentPart> modernComments = sPart.commentParts;
                    foreach (PowerPointCommentPart modernComment in modernComments)
                    {
                        foreach (ModernComment.Comment c in modernComment.CommentList)
                        {
                            string commentAuthor = string.Empty;
                            foreach (ModernComment.Author a in pPart.authorsPart.AuthorList)
                            {
                                if (a.Id == c.AuthorId)
                                {
                                    commentAuthor = a.Name;
                                }
                            }

                            commentCount++;
                            tList.Add(commentCount + Strings.wPeriod + "Author: " + commentAuthor + " Comment: " + c.InnerText);
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetSlideTitles(string fPath)
        {
            List<string> tList = new List<string>();
            using (PresentationDocument presentationDocument = PresentationDocument.Open(fPath, false))
            {
                int slideCount = 0;

                foreach (string s in GetSlideTitles(presentationDocument))
                {
                    slideCount++;
                    tList.Add(slideCount + Strings.wPeriod + s);
                }
            }

            return tList;
        }

        public static List<string> GetSlideText(string fPath)
        {
            List<string> tList = new List<string>();

            int sCount = RetrieveNumberOfSlides(fPath);
            if (sCount > 0)
            {
                int count = 0;

                do
                {
                    GetSlideIdAndText(out string sldText, fPath, count);
                    tList.Add("Slide " + (count + 1) + Strings.wPeriod + sldText);
                    count++;
                } while (count < sCount);
            }

            return tList;
        }

        public static List<string> GetSlideTransitions(string fPath)
        {
            List<string> tList = new List<string>();
            using (PresentationDocument ppt = PresentationDocument.Open(fPath, false))
            {
                int transitionCount = 0;
                foreach (string s in GetSlideTransitions(ppt))
                {
                    transitionCount++;
                    tList.Add(transitionCount + Strings.wPeriod + s);
                }
            }

            return tList;
        }

        // Get a list of the transitions of all the slides in the presentation.
        public static IList<string> GetSlideTransitions(PresentationDocument presentationDocument)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    List<string> transitionsList = new List<string>();

                    // Get the transition of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        string transition = string.Empty;

                        if (slidePart.Slide.Transition != null)
                        {
                            foreach (var t in slidePart.Slide.Transition)
                            {
                                transition = t.LocalName;
                            }
                        }
                        else
                        {
                            transition = "none";
                        }

                        // An empty title can also be added.
                        transitionsList.Add(transition);
                    }

                    return transitionsList;
                }
            }

            return null;
        }

        /// <summary>
        /// Get the slideId and text for that slide
        /// </summary>
        /// <param name="sldText">string returned to caller</param>
        /// <param name="docName">path to powerpoint file</param>
        /// <param name="index">slide number</param>
        public static void GetSlideIdAndText(out string sldText, string fPath, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(fPath, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }
                sldText = paragraphText.ToString();
            }
        }

        public static void DeleteUnusedSlideLayoutParts(PresentationDocument ppt, List<string> usedSlideLayoutIds)
        {
            foreach (SlideMasterPart smp in ppt.PresentationPart.SlideMasterParts)
            {
                //parts.Clear();
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    bool sIdNotFound = true;

                    foreach (string sId in usedSlideLayoutIds)
                    {
                        if (sId == slp.Uri.ToString())
                        {
                            sIdNotFound = false;
                        }
                    }

                    if (sIdNotFound)
                    {
                        smp.DeletePart(slp);
                        ppt.Save();
                    }
                }
            }
        }

        public static List<string> GetSlideLayoutId(PresentationDocument ppt)
        {
            List<string> slideLayoutIds = new List<string>();

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = ppt.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        SlideLayoutPart slideLayoutPart = slidePart.SlideLayoutPart;
                        slideLayoutIds.Add(slideLayoutPart.Uri.ToString());
                    }
                }
            }

            slideLayoutIds = slideLayoutIds.Distinct().ToList();

            return slideLayoutIds;
        }

        public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    List<string> titlesList = new List<string>();

                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                        // Get the slide title.
                        string title = GetSlideTitle(slidePart);

                        // An empty title can also be added.
                        titlesList.Add(title);
                    }

                    return titlesList;
                }
            }

            return null;
        }

        // Get the title string of the slide.
        public static string GetSlideTitle(SlidePart slidePart)
        {
            if (slidePart is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Declare a paragraph separator.
            string paragraphSeparator = null;

            if (slidePart.Slide != null)
            {
                // Find all the title shapes.
                var shapes = from shape in slidePart.Slide.Descendants<PShape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();

                foreach (var shape in shapes)
                {
                    // Get the text in each paragraph in this shape.
                    foreach (var paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
                    {
                        // Add a line break.
                        paragraphText.Append(paragraphSeparator);

                        foreach (var text in paragraph.Descendants<Drawing.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                return paragraphText.ToString();
            }

            return string.Empty;
        }

        /// <summary>
        /// Determines whether the shape is a title shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static bool IsTitleShape(PShape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch (placeholderShape.Type.ToString())
                {
                    case "title":
                    case "ctrTitle":
                    case "subtitle":
                        return true;
                    default:
                        return false;
                }
            }
            return false;
        }

        // Returns all the external hyperlinks in the slides of a presentation.
        public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fPath)
        {
            // Declare a list of strings.
            List<string> ret = new List<string>();

            PresentationDocument document = PresentationDocument.Open(fPath, false);

            // Iterate through all the slide parts in the presentation part.
            foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
            {
                IEnumerable<HyperlinkType> links = slidePart.Slide.Descendants<HyperlinkType>();

                // Iterate through all the links in the slide part.
                foreach (HyperlinkType link in links)
                {
                    // Iterate through all the external relationships in the slide part. 
                    foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
                    {
                        // If the relationship ID matches the link ID
                        if (relation.Id.Equals(link.Id))
                        {
                            // Add the URI of the external relationship to the list of strings.
                            ret.Add(relation.Uri.AbsoluteUri);
                        }
                    }
                }
            }

            // Return the list of strings.
            return ret;
        }

        // Delete comments by a specific author. Pass an empty string for the author to delete all comments, by all authors.
        public static bool DeleteComments(string fileName, string author)
        {
            bool isChanged = false;
            PresentationDocument doc = PresentationDocument.Open(fileName, true);

            // Get the authors part.
            CommentAuthorsPart authorsPart = doc.PresentationPart.GetPartsOfType<CommentAuthorsPart>().FirstOrDefault();

            if (authorsPart is null)
            {
                // There's no authors part, so just
                // fail. If no authors, there can't be any comments.
                return isChanged;
            }

            // Get the comment authors, or the specified author if supplied:
            var commentAuthors = authorsPart.CommentAuthorList.Elements<CommentAuthor>();
            if (!string.IsNullOrEmpty(author))
            {
                commentAuthors = commentAuthors.Where(e => e.Name.Value.Equals(author));
            }

            bool changed = false;
            foreach (var commentAuthor in commentAuthors.ToArray())
            {
                var authorId = commentAuthor.Id;

                // Iterate through all the slides and get the slide parts.
                foreach (var slide in doc.PresentationPart.GetPartsOfType<SlidePart>())
                {
                    // Iterate through the slide parts and find the slide comment parts.
                    var slideCommentParts = slide.GetPartsOfType<SlideCommentsPart>().ToArray();

                    foreach (var slideCommentsPart in slideCommentParts)
                    {
                        // Get the list of comments.
                        var commentList = slideCommentsPart.CommentList.Elements<Comment>().
                          Where(e => e.AuthorId.Value == authorId.Value);

                        foreach (var comment in commentList.ToArray())
                        {
                            // Delete all the comments by the specified author.
                            slideCommentsPart.CommentList.RemoveChild<Comment>(comment);
                            isChanged = true;
                        }

                        // No comments left? Delete the comments part for this slide.
                        if (slideCommentsPart.CommentList.Count() == 0)
                        {
                            slide.DeletePart(slideCommentsPart);
                        }
                        else
                        {
                            // Save the slide comments part.
                            slideCommentsPart.CommentList.Save();
                        }
                    }
                }

                // Delete the comment author from the comment authors part.
                authorsPart.CommentAuthorList.RemoveChild<CommentAuthor>(commentAuthor);

                changed = true;
            }

            // Changed will only be false if the caller requested comments
            // for a particular author, and that author has no comments.
            if (changed)
            {
                if (authorsPart.CommentAuthorList.Count() == 0)
                {
                    // No authors left, so delete the part.
                    doc.PresentationPart.DeletePart(authorsPart);
                }
                else
                {
                    // Save the comment authors part.
                    authorsPart.CommentAuthorList.Save();
                }
            }

            return isChanged;
        }

        public static int GetSlideIndexByTitle(string fileName, string slideTitle)
        {
            // Given a slide document and a slide title, retrieve the 0-based index of the 
            // first slide with a matching title. Return -1 if the title isn't found.

            // Assume that you won't find a match.
            int slideLocation = -1;

            using (var document = PresentationDocument.Open(fileName, true))
            {
                var presPart = document.PresentationPart;

                // No presentation part? Something's wrong with the document.
                if (presPart == null)
                {
                    throw new ArgumentException("fileName");
                }

                // If you're here, you know that presentationPart exists.
                var slideIdList = presPart.Presentation.SlideIdList;
                // Go through the slides in order.
                // This requires investigating the actual slide IDs, rather 
                // than just retrieving the slide parts.
                int index = 0;
                foreach (var slideId in slideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = (SlidePart)(presPart.GetPartById(slideId.RelationshipId.ToString()));

                    if (slidePart == null)
                    {
                        throw new ArgumentNullException("presentationDocument");
                    }

                    Slide theSlide = slidePart.Slide;
                    if (theSlide != null)
                    {

                        // Assume the first title shape you find contains the title.
                        var titleShape = slidePart.Slide.Descendants<A.Shape>().
                          Where(s => IsTitleShape(s)).FirstOrDefault();
                        if (titleShape != null)
                        {
                            // Compare the title, case-insensitively.
                            if (string.Compare(titleShape.InnerText, slideTitle, true) == 0)
                            {
                                slideLocation = index;
                                break;
                            }
                            else
                            {
                                index += 1;
                            }
                        }
                    }
                }
            }
            return slideLocation;
        }

        private static bool IsTitleShape(A.Shape shape)
        {
            bool isTitle = false;

            var placeholderShape = shape.NonVisualShapeProperties.NonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (((placeholderShape) != null) && (((placeholderShape.Type) != null) &&
              placeholderShape.Type.HasValue))
            {
                // Any title shape
                if (placeholderShape.Type.Value == PlaceholderValues.Title)
                {
                    isTitle = true;

                }
                // A centered title.
                else if (placeholderShape.Type.Value == PlaceholderValues.CenteredTitle)
                {
                    isTitle = true;
                }
            }
            return isTitle;
        }

        // Return the number of slides, including hidden slides.
        public static int GetSlideCount(string fileName, bool includeHidden)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart != null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.GetPartsOfType<SlidePart>().Count();
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden will contain the value "0".
                        // The Show property may not exist, and most likely will not, for non-hidden slides.
                        var slides = presentationPart.GetPartsOfType<SlidePart>().
                          Where((s) => (s.Slide != null) &&
                            ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }
    }
}
