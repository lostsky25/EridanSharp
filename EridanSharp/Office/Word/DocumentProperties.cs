using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EridanSharp
{
    namespace Office
    {
        public class DocumentProperties
        {
            private object fileName = System.Reflection.Missing.Value;
            private object fileFormat = System.Reflection.Missing.Value;
            private object lockComments = System.Reflection.Missing.Value;
            private object password = System.Reflection.Missing.Value;
            private object addToRecentFiles = System.Reflection.Missing.Value;
            private object writePassword = System.Reflection.Missing.Value;
            private object readOnlyRecommended = System.Reflection.Missing.Value;
            private object embedTrueTypeFonts = System.Reflection.Missing.Value;
            private object saveNativePictureFormat = System.Reflection.Missing.Value;
            private object saveFormsData = System.Reflection.Missing.Value;
            private object saveAsAOCELetter = System.Reflection.Missing.Value;
            private object encoding = System.Reflection.Missing.Value;
            private object insertLineBreaks = System.Reflection.Missing.Value;
            private object allowSubstitutions = System.Reflection.Missing.Value;
            private object lineEnding = System.Reflection.Missing.Value;
            private object addBiDiMarks = System.Reflection.Missing.Value;

            public enum LineEndingCode
            {
                /// <summary>
                /// Carriage return plus line feed.
                /// </summary>
                wdCRLF = 0,
                /// <summary>
                /// Carriage return only.
                /// </summary>
                wdCROnly = 1,
                /// <summary>
                /// Line feed plus carriage return.
                /// </summary>
                wdLFCR = 3,
                /// <summary>
                /// Line feed only.
                /// </summary>
                wdLFOnly = 2,
                /// <summary>
                /// Not supported.
                /// </summary>
                wdLSPS = 4,
            }
            public enum EncodingCode
            {
                /// <summary>
                /// Arabic.
                /// </summary>
                msoEncodingArabic = 1256,
                /// <summary>
                /// Arabic ASMO.
                /// </summary>
                msoEncodingArabicASMO = 708,
                /// <summary>
                /// Web browser auto-detects type of Arabic encoding to use.
                /// </summary>
                msoEncodingArabicAutoDetect = 51256,
                /// <summary>
                /// Transparent Arabic.
                /// </summary>
                msoEncodingArabicTransparentASMO = 720,
                /// <summary>
                /// Web browser auto-detects type of encoding to use.
                /// </summary>
                msoEncodingAutoDetect = 50001,
                /// <summary>
                /// Baltic.
                /// </summary>
                msoEncodingBaltic = 1257,
                /// <summary>
                /// Central European.
                /// </summary>
                msoEncodingCentralEuropean = 1250,
                /// <summary>
                /// Cyrillic.
                /// </summary>
                msoEncodingCyrillic = 1251,
                /// <summary>
                /// Web browser auto-detects type of Cyrillic encoding to use.
                /// </summary>
                msoEncodingCyrillicAutoDetect = 51251,
                /// <summary>
                /// Extended Binary Coded Decimal Interchange Code (EBCDIC) Arabic.
                /// </summary>
                msoEncodingEBCDICArabic = 20420,
                /// <summary>
                /// EBCDIC as used in Denmark and Norway.
                /// </summary>
                msoEncodingEBCDICDenmarkNorway = 20277,
                /// <summary>
                /// EBCDIC as used in Finland and Sweden.
                /// </summary>
                msoEncodingEBCDICFinlandSweden = 20278,
                /// <summary>
                /// EBCDIC as used in France.
                /// </summary>
                msoEncodingEBCDICFrance = 20297,
                /// <summary>
                /// EBCDIC as used in Germany.
                /// </summary>
                msoEncodingEBCDICGermany = 20273,
                /// <summary>
                /// EBCDIC as used in the Greek language.
                /// </summary>
                msoEncodingEBCDICGreek = 20423,
                /// <summary>
                /// EBCDIC as used in the Modern Greek language.
                /// </summary>
                msoEncodingEBCDICGreekModern = 875,
                /// <summary>
                /// EBCDIC as used in the Hebrew language.
                /// </summary>
                msoEncodingEBCDICHebrew = 20424,
                /// <summary>
                /// EBCDIC as used in Iceland.
                /// </summary>
                msoEncodingEBCDICIcelandic = 20871,
                /// <summary>
                /// International EBCDIC.
                /// </summary>
                msoEncodingEBCDICInternational = 500,
                /// <summary>
                /// EBCDIC as used in Italy.
                /// </summary>
                msoEncodingEBCDICItaly = 20280,
                /// <summary>
                /// EBCDIC as used with Japanese Katakana (extended).
                /// </summary>
                msoEncodingEBCDICJapaneseKatakanaExtended = 20290,
                /// <summary>
                /// EBCDIC as used with Japanese Katakana (extended) and Japanese.
                /// </summary>
                msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930,
                /// <summary>
                /// EBCDIC as used with Japanese Latin (extended) and Japanese.
                /// </summary>
                msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939,
                /// <summary>
                /// EBCDIC as used with Korean (extended).
                /// </summary>
                msoEncodingEBCDICKoreanExtended = 20833,
                /// <summary>
                /// EBCDIC as used with Korean (extended) and Korean.
                /// </summary>
                msoEncodingEBCDICKoreanExtendedAndKorean = 50933,
                /// <summary>
                /// EBCDIC as used in Latin America and Spain.
                /// </summary>
                msoEncodingEBCDICLatinAmericaSpain = 20284,
                /// <summary>
                /// EBCDIC Multilingual ROECE (Latin 2).
                /// </summary>
                msoEncodingEBCDICMultilingualROECELatin2 = 870,
                /// <summary>
                /// EBCDIC as used with Russian.
                /// </summary>
                msoEncodingEBCDICRussian = 20880,
                /// <summary>
                /// EBCDIC as used with Serbian and Bulgarian.
                /// </summary>
                msoEncodingEBCDICSerbianBulgarian = 21025,
                /// <summary>
                /// EBCDIC as used with Simplified Chinese(extended) and Simplified Chinese.
                /// </summary>
                msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935,
                /// <summary>
                /// EBCDIC as used with Thai.
                /// </summary>
                msoEncodingEBCDICThai = 20838,
                /// <summary>
                /// EBCDIC as used with Turkish.
                /// </summary>
                msoEncodingEBCDICTurkish = 20905,
                /// <summary>
                /// EBCDIC as used with Turkish(Latin 5).
                /// </summary>
                msoEncodingEBCDICTurkishLatin5 = 1026,
                /// <summary>
                /// EBCDIC as used in the United Kingdom.
                /// </summary>
                msoEncodingEBCDICUnitedKingdom = 20285,
                /// <summary>
                /// EBCDIC as used in the United States and Canada.
                /// </summary>
                msoEncodingEBCDICUSCanada = 37,
                /// <summary>
                /// EBCDIC as used in the United States and Canada, and with Japanese.
                /// </summary>
                msoEncodingEBCDICUSCanadaAndJapanese = 50931,
                /// <summary>
                /// EBCDIC as used in the United States and Canada, and with Traditional Chinese.
                /// </summary>
                msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937,
                /// <summary>
                /// Extended Unix Code (EUC) as used with Chinese and Simplified Chinese.
                /// </summary>
                msoEncodingEUCChineseSimplifiedChinese = 51936,
                /// <summary>
                /// EUC as used with Japanese.
                /// </summary>
                msoEncodingEUCJapanese = 51932,
                /// <summary>
                /// EUC as used with Korean.
                /// </summary>
                msoEncodingEUCKorean = 51949,
                /// <summary>
                /// EUC as used with Taiwanese and Traditional Chinese.
                /// </summary>
                msoEncodingEUCTaiwaneseTraditionalChinese = 51950,
                /// <summary>
                /// Europa.
                /// </summary>
                msoEncodingEuropa3 = 29001,
                /// <summary>
                /// Extended Alpha lowercase.
                /// </summary>
                msoEncodingExtAlphaLowercase = 21027,
                /// <summary>
                /// Greek.
                /// </summary>
                msoEncodingGreek = 1253,
                /// <summary>
                /// Web browser auto-detects type of Greek encoding to use.
                /// </summary>
                msoEncodingGreekAutoDetect = 51253,
                /// <summary>
                /// Hebrew.
                /// </summary>
                msoEncodingHebrew = 1255,
                /// <summary>
                /// Simplified Chinese (HZGB).
                /// </summary>
                msoEncodingHZGBSimplifiedChinese = 52936,
                /// <summary>
                /// German (International Alphabet No. 5, or IA5).
                /// </summary>
                msoEncodingIA5German = 20106,
                /// <summary>
                /// IA5, International Reference Version(IRV).
                /// </summary>
                msoEncodingIA5IRV = 20105,
                /// <summary>
                /// IA5 as used with Norwegian.
                /// </summary>
                msoEncodingIA5Norwegian = 20108,
                /// <summary>
                /// IA5 as used with Swedish.
                /// </summary>
                msoEncodingIA5Swedish = 20107,
                /// <summary>
                /// Indian Script Code for Information Interchange(ISCII) as used with Assamese.
                /// </summary>
                msoEncodingISCIIAssamese = 57006,
                /// <summary>
                /// ISCII as used with Bengali.
                /// </summary>
                msoEncodingISCIIBengali = 57003,
                /// <summary>
                /// ISCII as used with Devanagari.
                /// </summary>
                msoEncodingISCIIDevanagari = 57002,
                /// <summary>
                /// ISCII as used with Gujarati.
                /// </summary>
                msoEncodingISCIIGujarati = 57010,
                /// <summary>
                /// ISCII as used with Kannada.
                /// </summary>
                msoEncodingISCIIKannada = 57008,
                /// <summary>
                /// ISCII as used with Malayalam.
                /// </summary>
                msoEncodingISCIIMalayalam = 57009,
                /// <summary>
                /// ISCII as used with Oriya.
                /// </summary>
                msoEncodingISCIIOriya = 57007,
                /// <summary>
                /// ISCII as used with Punjabi.
                /// </summary>
                msoEncodingISCIIPunjabi = 57011,
                /// <summary>
                /// ISCII as used with Tamil.
                /// </summary>
                msoEncodingISCIITamil = 57004,
                /// <summary>
                /// ISCII as used with Telugu.
                /// </summary>
                msoEncodingISCIITelugu = 57005,
                /// <summary>
                /// ISO 2022-CN encoding as used with Simplified Chinese.
                /// </summary>
                msoEncodingISO2022CNSimplifiedChinese = 50229,
                /// <summary>
                /// ISO 2022-CN encoding as used with Traditional Chinese.
                /// </summary>
                msoEncodingISO2022CNTraditionalChinese = 50227,
                /// <summary>
                /// ISO 2022-JP
                /// </summary>
                msoEncodingISO2022JPJISX02011989 = 50222,
                /// <summary>
                /// ISO 2022-JP
                /// </summary>
                msoEncodingISO2022JPJISX02021984 = 50221,
                /// <summary>
                /// ISO 2022-JP with no half-width Katakana.
                /// </summary>
                msoEncodingISO2022JPNoHalfwidthKatakana = 50220,
                /// <summary>
                /// ISO 2022-KR.
                /// </summary>
                msoEncodingISO2022KR = 50225,
                /// <summary>
                /// ISO 6937 Non-Spacing Accent.
                /// </summary>
                msoEncodingISO6937NonSpacingAccent = 20269,
                /// <summary>
                /// ISO 8859-15 with Latin 9.
                /// </summary>
                msoEncodingISO885915Latin9 = 28605,
                /// <summary>
                /// ISO 8859-1 Latin 1.
                /// </summary>
                msoEncodingISO88591Latin1 = 28591,
                /// <summary>
                /// ISO 8859-2 Central Europe.
                /// </summary>
                msoEncodingISO88592CentralEurope = 28592,
                /// <summary>
                /// ISO 8859-3 Latin 3.
                /// </summary>
                msoEncodingISO88593Latin3 = 28593,
                /// <summary>
                /// ISO 8859-4 Baltic.
                /// </summary>
                msoEncodingISO88594Baltic = 28594,
                /// <summary>
                /// ISO 8859-5 Cyrillic.
                /// </summary>
                msoEncodingISO88595Cyrillic = 28595,
                /// <summary>
                /// ISA 8859-6 Arabic.
                /// </summary>
                msoEncodingISO88596Arabic = 28596,
                /// <summary>
                /// ISO 8859-7 Greek.
                /// </summary>
                msoEncodingISO88597Greek = 28597,
                /// <summary>
                /// ISO 8859-8 Hebrew.
                /// </summary>
                msoEncodingISO88598Hebrew = 28598,
                /// <summary>
                /// ISO 8859-8 Hebrew (Logical).
                /// </summary>
                msoEncodingISO88598HebrewLogical = 38598,
                /// <summary>
                /// ISO 8859-9 Turkish.
                /// </summary>
                msoEncodingISO88599Turkish = 28599,
                /// <summary>
                /// Web browser auto-detects type of Japanese encoding to use.
                /// </summary>
                msoEncodingJapaneseAutoDetect = 50932,
                /// <summary>
                /// Japanese (Shift-JIS).
                /// </summary>
                msoEncodingJapaneseShiftJIS = 932,
                /// <summary>
                /// KOI8-R.
                /// </summary>
                msoEncodingKOI8R = 20866,
                /// <summary>
                /// K0I8-U.
                /// </summary>
                msoEncodingKOI8U = 21866,
                /// <summary>
                /// Korean.
                /// </summary>
                msoEncodingKorean = 949,
                /// <summary>
                /// Web browser auto-detects type of Korean encoding to use.
                /// </summary>
                msoEncodingKoreanAutoDetect = 50949,
                /// <summary>
                /// Korean(Johab).
                /// </summary>
                msoEncodingKoreanJohab = 1361,
                /// <summary>
                /// Macintosh Arabic.
                /// </summary>
                msoEncodingMacArabic = 10004,
                /// <summary>
                /// Macintosh Croatian.
                /// </summary>
                msoEncodingMacCroatia = 10082,
                /// <summary>
                /// Macintosh Cyrillic.
                /// </summary>
                msoEncodingMacCyrillic = 10007,
                /// <summary>
                /// Macintosh Greek.
                /// </summary>
                msoEncodingMacGreek1 = 10006,
                /// <summary>
                /// Macintosh Hebrew.
                /// </summary>
                msoEncodingMacHebrew = 10005,
                /// <summary>
                /// Macintosh Icelandic.
                /// </summary>
                msoEncodingMacIcelandic = 10079,
                /// <summary>
                /// Macintosh Japanese.
                /// </summary>
                msoEncodingMacJapanese = 10001,
                /// <summary>
                /// Macintosh Korean.
                /// </summary>
                msoEncodingMacKorean = 10003,
                /// <summary>
                /// Macintosh Latin 2.
                /// </summary>
                msoEncodingMacLatin2 = 10029,
                /// <summary>
                /// Macintosh Roman.
                /// </summary>
                msoEncodingMacRoman = 10000,
                /// <summary>
                /// Macintosh Romanian.
                /// </summary>
                msoEncodingMacRomania = 10010,
                /// <summary>
                /// Macintosh Simplified Chinese (GB 2312).
                /// </summary>
                msoEncodingMacSimplifiedChineseGB2312 = 10008,
                /// <summary>
                /// Macintosh Traditional Chinese(Big 5).
                /// </summary>
                msoEncodingMacTraditionalChineseBig5 = 10002,
                /// <summary>
                /// Macintosh Turkish.
                /// </summary>
                msoEncodingMacTurkish = 10081,
                /// <summary>
                /// Macintosh Ukrainian.
                /// </summary>
                msoEncodingMacUkraine = 10017,
                /// <summary>
                /// OEM as used with Arabic.
                /// </summary>
                msoEncodingOEMArabic = 864,
                /// <summary>
                /// OEM as used with Baltic.
                /// </summary>
                msoEncodingOEMBaltic = 775,
                /// <summary>
                /// OEM as used with Canadian French.
                /// </summary>
                msoEncodingOEMCanadianFrench = 863,
                /// <summary>
                /// OEM as used with Cyrillic.
                /// </summary>
                msoEncodingOEMCyrillic = 855,
                /// <summary>
                /// OEM as used with Cyrillic II.
                /// </summary>
                msoEncodingOEMCyrillicII = 866,
                /// <summary>
                /// OEM as used with Greek 437G.
                /// </summary>
                msoEncodingOEMGreek437G = 737,
                /// <summary>
                /// OEM as used with Hebrew.
                /// </summary>
                msoEncodingOEMHebrew = 862,
                /// <summary>
                /// OEM as used with Icelandic.
                /// </summary>
                msoEncodingOEMIcelandic = 861,
                /// <summary>
                /// OEM as used with Modern Greek.
                /// </summary>
                msoEncodingOEMModernGreek = 869,
                /// <summary>
                /// OEM as used with multi-lingual Latin I.
                /// </summary>
                msoEncodingOEMMultilingualLatinI = 850,
                /// <summary>
                /// OEM as used with multi-lingual Latin II.
                /// </summary>
                msoEncodingOEMMultilingualLatinII = 852,
                /// <summary>
                /// OEM as used with Nordic languages.
                /// </summary>
                msoEncodingOEMNordic = 865,
                /// <summary>
                /// OEM as used with Portuguese.
                /// </summary>
                msoEncodingOEMPortuguese = 860,
                /// <summary>
                /// OEM as used with Turkish.
                /// </summary>
                msoEncodingOEMTurkish = 857,
                /// <summary>
                /// OEM as used in the United States.
                /// </summary>
                msoEncodingOEMUnitedStates = 437,
                /// <summary>
                /// Web browser auto-detects type of Simplified Chinese encoding to use.
                /// </summary>
                msoEncodingSimplifiedChineseAutoDetect = 50936,
                /// <summary>
                /// Simplified Chinese GB 18030.
                /// </summary>
                msoEncodingSimplifiedChineseGB18030 = 54936,
                /// <summary>
                /// Simplified Chinese GBK.
                /// </summary>
                msoEncodingSimplifiedChineseGBK = 936,
                /// <summary>
                /// T61.
                /// </summary>
                msoEncodingT61 = 20261,
                /// <summary>
                /// Taiwan CNS.
                /// </summary>
                msoEncodingTaiwanCNS = 20000,
                /// <summary>
                /// Taiwan Eten.
                /// </summary>
                msoEncodingTaiwanEten = 20002,
                /// <summary>
                /// Taiwan IBM 5550.
                /// </summary>
                msoEncodingTaiwanIBM5550 = 20003,
                /// <summary>
                /// Taiwan TCA.
                /// </summary>
                msoEncodingTaiwanTCA = 20001,
                /// <summary>
                /// Taiwan Teletext.
                /// </summary>
                msoEncodingTaiwanTeleText = 20004,
                /// <summary>
                /// Taiwan Wang.
                /// </summary>
                msoEncodingTaiwanWang = 20005,
                /// <summary>
                /// Thai.
                /// </summary>
                msoEncodingThai = 874,
                /// <summary>
                /// Web browser auto-detects type of Traditional Chinese encoding to use.
                /// </summary>
                msoEncodingTraditionalChineseAutoDetect = 50950,
                /// <summary>
                /// Traditional Chinese Big 5.
                /// </summary>
                msoEncodingTraditionalChineseBig5 = 950,
                /// <summary>
                /// Turkish.
                /// </summary>
                msoEncodingTurkish = 1254,
                /// <summary>
                /// Unicode big endian.
                /// </summary>
                msoEncodingUnicodeBigEndian = 1201,
                /// <summary>
                /// Unicode little endian.
                /// </summary>
                msoEncodingUnicodeLittleEndian = 1200,
                /// <summary>
                /// United States ASCII.
                /// </summary>
                msoEncodingUSASCII = 20127,
                /// <summary>
                /// UTF-7 encoding.
                /// </summary>
                msoEncodingUTF7 = 65000,
                /// <summary>
                /// UTF-8 encoding.
                /// </summary>
                msoEncodingUTF8 = 65001,
                /// <summary>
                /// Vietnamese.
                /// </summary>
                msoEncodingVietnamese = 1258,
                /// <summary>
                /// Western.
                /// </summary>
                msoEncodingWestern = 1252,
            }
            public enum FormatCode
            {
                /// <summary>
                /// Microsoft Word format.
                /// </summary>
                wdFormatDocument = 0,

                /// <summary>
                /// Word default document file format. For Microsoft Office Word 2007, this is the DOCX format.
                /// </summary>
                wdFormatDocumentDefault = 16,

                /// <summary>
                /// Microsoft DOS text format.
                /// </summary>
                wdFormatDOSText = 4,

                /// <summary>
                /// Microsoft DOS text with line breaks preserved.
                /// </summary>
                wdFormatDOSTextLineBreaks = 5,

                /// <summary>
                /// Encoded text format.
                /// </summary>
                wdFormatEncodedText = 7,

                /// <summary>
                /// Filtered HTML format.
                /// </summary>
                wdFormatFilteredHTML = 10,

                /// <summary>
                /// Reserved for internal use.
                /// </summary>
                wdFormatFlatXML = 19,

                /// <summary>
                /// Reserved for internal use.
                /// </summary>
                wdFormatFlatXMLMacroEnabled = 20,

                /// <summary>
                /// Reserved for internal use.
                /// </summary>
                wdFormatFlatXMLTemplate = 21,

                /// <summary>
                /// Reserved for internal use.
                /// </summary>
                wdFormatFlatXMLTemplateMacroEnabled = 22,

                /// <summary>
                /// Standard HTML format.
                /// </summary>
                wdFormatHTML = 8,

                /// <summary>
                /// 
                /// </summary>
                wdFormatOpenDocumentText = 23,

                /// <summary>
                /// PDF format.
                /// </summary>
                wdFormatPDF = 17,

                /// <summary>
                /// Rich text format (RTF).
                /// </summary>
                wdFormatRTF = 6,

                /// <summary>
                /// Strict Open XML document format.
                /// </summary>
                wdFormatStrictOpenXMLDocument = 24,

                /// <summary>
                /// Microsoft Word template format.
                /// </summary>
                wdFormatTemplate = 1,

                /// <summary>
                /// Word 97 template format.
                /// </summary>
                wdFormatTemplate97 = 1,

                /// <summary>
                /// Microsoft Windows text format.
                /// </summary>
                wdFormatText = 2,

                /// <summary>
                /// Microsoft Windows text format with line breaks preserved.
                /// </summary>
                wdFormatTextLineBreaks = 3,

                /// <summary>
                /// Unicode text format.
                /// </summary>
                wdFormatUnicodeText = 7,

                /// <summary>
                /// Web archive format.
                /// </summary>
                wdFormatWebArchive = 9,

                /// <summary>
                /// Extensible Markup Language (XML) format.
                /// </summary>
                wdFormatXML = 11,

                /// <summary>
                /// XML document format.
                /// </summary>
                wdFormatXMLDocument = 12,

                /// <summary>
                /// XML template format with macros enabled.
                /// </summary>
                wdFormatXMLDocumentMacroEnabled = 13,

                /// <summary>
                /// XML template format.
                /// </summary>
                wdFormatXMLTemplate = 14,

                /// <summary>
                /// XML template format with macros enabled.
                /// </summary>
                wdFormatXMLTemplateMacroEnabled = 15,

                /// <summary>
                /// XPS format.
                /// </summary>
                wdFormatXPS = 18,
            }

            public string FileName
            {
                get
                {
                    return (string)fileName;
                }
                set
                {
                    fileName = value;
                }
            }
            public FormatCode FileFormat
            {
                get
                {
                    return (FormatCode)fileFormat;
                }
                set
                {
                    fileFormat = value;
                }
            }
            public bool LockComments
            {
                get
                {
                    return (bool)lockComments;
                }
                set
                {
                    lockComments = value;
                }
            }
            public string Password
            {
                get
                {
                    return (string)password;
                }
                set
                {
                    password = value;
                }
            }
            public bool AddToRecentFiles
            {
                get
                {
                    return (bool)addToRecentFiles;
                }
                set
                {
                    addToRecentFiles = value;
                }
            }
            public string WritePassword
            {
                get
                {
                    return (string)writePassword;
                }
                set
                {
                    writePassword = value;
                }
            }
            public bool ReadOnlyRecommended
            {
                get
                {
                    return (bool)readOnlyRecommended;
                }
                set
                {
                    readOnlyRecommended = value;
                }
            }
            public bool EmbedTrueTypeFonts
            {
                get
                {
                    return (bool)embedTrueTypeFonts;
                }
                set
                {
                    embedTrueTypeFonts = value;
                }
            }
            public bool SaveNativePictureFormat
            {
                get
                {
                    return (bool)saveNativePictureFormat;
                }
                set
                {
                    saveNativePictureFormat = value;
                }
            }
            public bool SaveFormsData
            {
                get
                {
                    return (bool)saveFormsData;
                }
                set
                {
                    saveFormsData = value;
                }
            }
            public bool SaveAsAOCELetter
            {
                get
                {
                    return (bool)saveAsAOCELetter;
                }
                set
                {
                    saveAsAOCELetter = value;
                }
            }
            public EncodingCode Encoding
            {
                get
                {
                    return (EncodingCode)encoding;
                }
                set
                {
                    encoding = value;
                }
            }
            public bool InsertLineBreaks
            {
                get
                {
                    return (bool)insertLineBreaks;
                }
                set
                {
                    insertLineBreaks = value;
                }
            }
            public bool AllowSubstitutions
            {
                get
                {
                    return (bool)allowSubstitutions;
                }
                set
                {
                    allowSubstitutions = value;
                }
            }
            public LineEndingCode LineEnding
            {
                get
                {
                    return (LineEndingCode)lineEnding;
                }
                set
                {
                    lineEnding = value;
                }
            }
            public bool AddBiDiMarks
            {
                get
                {
                    return (bool)addBiDiMarks;
                }
                set
                {
                    addBiDiMarks = value;
                }
            }

        }
    }
}
