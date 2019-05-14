using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AutomationPDF.Test
{
    [TestClass]
    public class UnitTestPDF
    {
        [TestMethod]
        public void TestMethodDocument()
        {
            string sMyFile = @"F:\Users\uni19541\Desktop\TESTE Conversão PDF I.pdf";
            string sMyFileDest = @"F:\Users\uni19541\Desktop\NEW.pdf";
            using (MemoryStream oFileMemoryStream = new MemoryStream())
            {
                using (FileStream oFileStream = new FileStream(sMyFile, FileMode.Open, FileAccess.Read))
                {
                    oFileStream.CopyTo(oFileMemoryStream);
                    oFileMemoryStream.Position = 0;
                    DocumentPDF oPDF = new DocumentPDF(oFileMemoryStream);

                    string[][] oMatrix = new string[][] { new string[] { "Carlos Henrique Vieira Figueiredo", "ABD12345", "1000/2000" } };

                    var oResultStream = oPDF.InsertTable(oMatrix, 3, 6.0f, 12.0f);

                    //using (FileStream oDestFileStream = new FileStream(sMyFileDest, FileMode.Create, FileAccess.Write))
                    //    oResultStream.CopyTo(oDestFileStream);

                    oFileStream.Close();
                }

                oFileMemoryStream.Close();
            }
        }
    }
}
