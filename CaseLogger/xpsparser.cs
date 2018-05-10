using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Xps.Packaging;


namespace CaseLogger
{
    class xpsparser
    {
        XpsDocument newxpsdocument;

        public xpsparser(string path)
        {
            newxpsdocument = new XpsDocument(path, System.IO.FileAccess.Read);
        }

        public string extracttext()
        {
            IXpsFixedDocumentSequenceReader fixedDocSeqReader  = newxpsdocument.FixedDocumentSequenceReader;
            IXpsFixedDocumentReader _document = fixedDocSeqReader.FixedDocuments[0];
           // IXpsFixedPageReader _page  = _document.FixedPages[documentViewerElement.MasterPageNumber];
            string _fullPageText = "";

            for (int pagenumber = 0; pagenumber < _document.FixedPages.Count(); pagenumber++)
            {
                IXpsFixedPageReader _page = _document.FixedPages[pagenumber];
                StringBuilder _currentText = new StringBuilder();
                System.Xml.XmlReader _pageContentReader = _page.XmlReader;
         
                if (_pageContentReader != null)
                {
                    while (_pageContentReader.Read())
                    {
                        if (_pageContentReader.Name == "Glyphs")
                        {
                            if (_pageContentReader.HasAttributes)
                            {
                                if (_pageContentReader.GetAttribute("UnicodeString") != null)
                                {
                                    _currentText.
                                      Append(_pageContentReader.
                                      GetAttribute("UnicodeString"));
                                }
                            }
                        }
                    }
                }
                _fullPageText = _fullPageText + _currentText.ToString();
            }
            return _fullPageText;
        }
    }
}
