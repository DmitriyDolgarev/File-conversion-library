using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;

namespace FileConverterLib.MSOffice
{
    public class MSOfficeConverter
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
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            if (Path.GetExtension(wordFileName).ToLower() != ".docx" && Path.GetExtension(wordFileName).ToLower() != ".doc")
                wordFileName = Path.ChangeExtension(wordFileName, "docx");

            ConvertWord(wordFileName, pdfFileName, WdSaveFormat.wdFormatPDF);
        }
        #endregion

        #region PDF to Word
        public static void PdfFileToDocxFile(string pdfFileName, string wordFileName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            if (Path.GetExtension(wordFileName).ToLower() != ".docx" && Path.GetExtension(wordFileName).ToLower() != ".doc")
                wordFileName = Path.ChangeExtension(wordFileName, "docx");

            ConvertWord(pdfFileName, wordFileName, WdSaveFormat.wdFormatDocumentDefault);
        }
        #endregion

        #region Pptx to PDF
        public static void PptxFileToPdfFile(string pptxFileName, string pdfFileName)
        {
            pdfFileName = Path.ChangeExtension(pdfFileName, "pdf");
            if (Path.GetExtension(pptxFileName).ToLower() != ".pptx" && Path.GetExtension(pptxFileName).ToLower() != ".ppt")
                pptxFileName = Path.ChangeExtension(pptxFileName, "pptx");

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
        #endregion
    }
}
