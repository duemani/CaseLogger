using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace CaseLogger
{
    class pdfparser
    {

        PdfReader reader;

        public pdfparser()
        {

        }

        public pdfparser(string path)
        {
            reader = new PdfReader(path);
        }

        public string extracttext()
        {
           string txt = "";
           for (int i = 1; i <= reader.NumberOfPages; i++)
           {
               txt = txt + PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
               txt = txt + "\n";
           }
           return txt;
        }
    }
}
