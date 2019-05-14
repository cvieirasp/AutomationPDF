using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace AutomationPDF
{
    public class DocumentPDF
    {
        Stream _filePDF = null;

        public DocumentPDF(MemoryStream oFilePDF)
        {
            _filePDF = oFilePDF;
        }

        public MemoryStream InsertTable(string[][] oMatrix, int iColumnsNumber, float fFontSize, float fTableSize)
        {
            //MemoryStream oResultStream = new MemoryStream();
            FileStream oResultStream = new FileStream(@"F:\Users\uni19541\Desktop\NEW.pdf", FileMode.Create, FileAccess.Write);
            try
            {
                // Valida parâmetros;
                if (oMatrix == null || oMatrix.Length == 0)
                    throw new ArgumentException("Matriz de Registros não contém valores.");

                PdfReader oReader = new PdfReader(_filePDF);
                Document oDocument = new Document(oReader.GetPageSizeWithRotation(1));
                PdfWriter oWriter = PdfWriter.GetInstance(oDocument, oResultStream);
                oDocument.Open();

                for (var oNumPage = 1; oNumPage <= oReader.NumberOfPages; oNumPage++)
                {
                    oDocument.NewPage();
                    var oImportedPage = oWriter.GetImportedPage(oReader, oNumPage);
                    PdfContentByte oPdfContentByte = oWriter.DirectContent;

                    PdfPTable oTable = new PdfPTable(iColumnsNumber)
                    {
                        TotalWidth = oDocument.GetRight(oDocument.RightMargin) - oDocument.GetLeft(oDocument.LeftMargin),
                        LockedWidth = true
                    };

                    for (int i = 0; i < oMatrix.Length; i++)
                    {
                        for (int j = 0; j < oMatrix[i].Length; j++)
                        {
                            Font oFontNormal = new Font(Font.FontFamily.HELVETICA, fFontSize, Font.NORMAL);
                            PdfPCell oCell = new PdfPCell(new Phrase(oMatrix[i][j], oFontNormal))
                            {
                                UseAscender = true,
                                HorizontalAlignment = Element.ALIGN_CENTER,
                                VerticalAlignment = Element.ALIGN_MIDDLE,
                                FixedHeight = fTableSize
                            };
                            oTable.AddCell(oCell);
                        }
                    }

                    oTable.WriteSelectedRows(
                        0, // Linha inicial
                        -1, // Linha final
                        oDocument.GetLeft(oDocument.LeftMargin),
                        oTable.TotalHeight + 10.0f,
                        oPdfContentByte
                    );

                    oPdfContentByte.AddTemplate(oImportedPage, 0, 0);
                }

                oDocument.Close();
                oWriter.Close();
                oReader.Close();



                /*
                // the pdf content
                PdfContentByte oPdfContentByte = oWriter.DirectContent;

                //for (int i = 1; i <= oReader.NumberOfPages; i++)
                //{
                //    oDocument.NewPage();
                //    oWriter.DirectContent.AddTemplate(oWriter.GetImportedPage(oReader, i), 1f, 0, 0, 1, 0, 0);
                //}

                PdfPTable oTable = new PdfPTable(iColumnsNumber)
                {
                    TotalWidth = oDocument.GetRight(oDocument.RightMargin) - oDocument.GetLeft(oDocument.LeftMargin),
                    LockedWidth = true
                };

                for (int i = 0; i < oMatrix.Length; i++)
                {
                    for (int j = 0; j < oMatrix[i].Length; j++)
                    {
                        Font oFontNormal = new Font(Font.FontFamily.HELVETICA, fFontSize, Font.NORMAL);
                        PdfPCell oCell = new PdfPCell(new Phrase(oMatrix[i][j], oFontNormal))
                        {
                            HorizontalAlignment = Element.ALIGN_CENTER,
                            VerticalAlignment = Element.ALIGN_MIDDLE,
                            FixedHeight = fTableSize
                        };
                        oTable.AddCell(oCell);
                    }
                }

                oTable.WriteSelectedRows(
                    0, // Linha inicial
                    -1, // Linha final
                    oDocument.GetLeft(oDocument.LeftMargin),
                    oTable.TotalHeight + 10.0f,
                    oPdfContentByte
                );

                // create the new page and add it to the pdf
                PdfImportedPage oPage = oWriter.GetImportedPage(oReader, 1);
                oPdfContentByte.AddTemplate(oPage, 0, 0);

                oDocument.Close();
                oWriter.Close();
                oReader.Close();
                */
            }
            catch (Exception ex)
            {
                //throw;
                var error = ex;
            }

            //oResultStream.Position = 0;
            return null;
        }
    }
}
