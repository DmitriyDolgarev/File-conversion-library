#if WINDOWS
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using FileConverterLib.Utils;

namespace FileConverterLib.MSOffice
{
    public partial class MSOfficeConverter
    {
        private static void ConvertWord(string input, string output, WdSaveFormat outputType)
        {
            Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application
            {
                Visible = false
            };

            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;
            object oInput = input;
            object oOutput = output;
            object oFormat = outputType;

            _Document oDoc = oWord.Documents.Open(
                ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            oDoc.Activate();

            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            oDoc.Close();
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }

        #region Word to PDF
        public static void DocxFileToPdfFile(string wordFileName, string pdfFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            wordFileName = FileConverterUtils.GetCorrectedPath(wordFileName, "docx");

            ConvertWord(wordFileName, pdfFileName, WdSaveFormat.wdFormatPDF);
        }

        public static void DocxFileToPdfFile(string wordFileName)
        {
            DocxFileToPdfFile(wordFileName, FileConverterUtils.GetFileNameInSameFolder(wordFileName));
        }
        #endregion

        #region PDF to Word
        public static void PdfFileToDocxFile(string pdfFileName, string wordFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            wordFileName = FileConverterUtils.GetCorrectedPath(wordFileName, "docx");

            ConvertWord(pdfFileName, wordFileName, WdSaveFormat.wdFormatDocumentDefault);
        }

        public static void PdfFileToDocxFile(string pdfFileName)
        {
            PdfFileToDocxFile(pdfFileName, FileConverterUtils.GetFileNameInSameFolder(pdfFileName));
        }
        #endregion

        #region Pptx to PDF
        public static void PptxFileToPdfFile(string pptxFileName, string pdfFileName)
        {
            pdfFileName = FileConverterUtils.GetCorrectedPath(pdfFileName, "pdf");
            pptxFileName = FileConverterUtils.GetCorrectedPath(pptxFileName, "pptx");

            object unknownType = Type.Missing;

            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

            Presentation pptPresentation = pptApplication.Presentations.Open(pptxFileName,
                    MsoTriState.msoTrue, MsoTriState.msoTrue,
                    MsoTriState.msoFalse);

            pptPresentation.ExportAsFixedFormat(pdfFileName,
                PpFixedFormatType.ppFixedFormatTypePDF,
                PpFixedFormatIntent.ppFixedFormatIntentPrint,
                MsoTriState.msoFalse,
                PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null,
                PpPrintRangeType.ppPrintAll, string.Empty, true, true, true,
                true, false, unknownType);

            pptPresentation.Close();
            pptApplication.Quit();
        }

        public static void PptxFileToPdfFile(string pptxFileName)
        {
            PptxFileToPdfFile(pptxFileName, FileConverterUtils.GetFileNameInSameFolder(pptxFileName));
        }
        #endregion
    }
}
#endif