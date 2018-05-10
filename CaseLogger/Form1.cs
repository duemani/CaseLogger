using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Windows.Xps.Packaging;
using System.Timers;

namespace CaseLogger
{
    public partial class Form1 : Form
    {
        string whereami;
        string residentname;
        string residentrole;
        string residentpgyyear;
        excelfile myexcelfile;
        excelfile myexcelfile1;
        excelfile myexcelfile2;
        string addressofexcelfile;
        int j;
        int totalnumberofcases;
        Form2 nameform;
        bool issenmodeactivated;
      //  System.Timers.Timer newtimer;
        System.Windows.Forms.Timer newtimer;

        public Form1()
        {
            InitializeComponent();
            residentrole = "39";

            seniorToolStripMenuItem.ForeColor = Color.Red;
            leadToolStripMenuItem.ForeColor = Color.Black;
            assistantToolStripMenuItem.ForeColor = Color.Black;

            residentpgyyear = "7";

            toolStripMenuItem8.ForeColor = Color.Red;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;

            residentname = "Reddy";
            totalnumberofcases = 0;
            nameform = new Form2(this);
            
            issenmodeactivated = false;
        }

        private void acgmeWebBrowser_DocumentCompleted(object sender, EventArgs e)
        {
            if (whereami == "loggingon")
            {
                HtmlDocument acgmedocument = acgmeWebBrowser.Document;
                j = 0;
                whereami = "inputtingfields";
            }

            if (whereami == "cancontinue")
            {
                submitToolStripMenuItem_Click(this, EventArgs.Empty);
            }
        }

        private void logonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            acgmeWebBrowser.Navigate("https://apps.acgme.org/connect/login?ReturnUrl=%252fconnect");
            acgmeWebBrowser.DocumentCompleted +=new WebBrowserDocumentCompletedEventHandler(acgmeWebBrowser_DocumentCompleted);
            whereami = "loggingon";
        }

        private string getfirstword(string tempstring)
        {
            char[] delimiterch = { ' ' };
            string[] words = tempstring.Split(delimiterch);
            return words[0];
        }

        private void getCPTsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myexcelfile1 = new excelfile(addressofexcelfile);
            int l = 1;
            while (l<=myexcelfile1.getlastrow())
            {
                string orcode = myexcelfile1.GetCellValue(string.Concat("h",l));
                {
                    string cptcode = "";
                    switch(orcode)
                    {
                        case("1670"):
                            cptcode = "61140";
                            break;
                        case("1134"):
                            cptcode = "61313";
                            break;
                        case("1133"):
                            cptcode = "62140";
                            break;
                        case("1132"):
                            cptcode = "61697"; //this is Leurssen, Bollo not available
                            break;
                        case("1022"):
                            cptcode = "61312";
                            break;
                        case("1346"):
                            cptcode = "61510";
                            break;
                        case ("1200"):
                            cptcode = "63081";
                            break;
                        case ("1775"):
                            cptcode = "22548";
                            break;
                        case ("1469"):
                            cptcode = "22600";
                            break;
                        case ("1202"):
                            cptcode = "61548";
                            break;
                        case ("1279"):
                            cptcode = "61885";
                            break;
                        case ("1163"):
                            cptcode = "63001";
                            break;
                        case ("1434"):
                            cptcode = "22808";
                            break;
                        case ("2008"):
                            cptcode = "22612";
                            break;
                        case ("1199"):
                            cptcode = "63005";
                            break;
                        case ("1501"):
                            cptcode = "62161";
                            break;
                        case ("1164"):
                            cptcode = "62223";
                            break;
                        case ("1402"):
                            cptcode = "61312";
                            break;
                        case ("1158"):
                            cptcode = "31600";
                            break;
                        case ("1165"):
                            cptcode = "62200";
                            break;
                    }
                    myexcelfile1.WriteCell(string.Concat("j", l),cptcode);
                    l++;
            }
        }
        myexcelfile1.closefile();
        MessageBox.Show("CPT Codes updated"); 
        }

        public static string StringWordsRemove(string stringToClean)
        {
            List<string> wordstoremove = "and or but of with for the an in using eg".Split(' ').ToList();
            return string.Join(" ", stringToClean.Split(new[] { ' ', ',', '.', '?', '!', ';', ':', '(', ')' }, StringSplitOptions.RemoveEmptyEntries).Except(wordstoremove));
        }

        private void submitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //            HtmlDocument acgmedocument = acgmeWebBrowser.Document;
            //            HtmlElement submitbutton = acgmedocument.GetElementById("submitButton"); 
                if (j == 0)
                {
                    myexcelfile2 = new excelfile(addressofexcelfile);
                }

                else if (j == totalnumberofcases)
                {
                    myexcelfile2.closefile();
                    MessageBox.Show("Done");
                    return;
                }

                j = j + 1;
                //              submitbutton.InvokeMember("click");
                label1.Visible = true;
                label1.Text = j + " / " + totalnumberofcases.ToString();

                HtmlDocument acgmedocument = acgmeWebBrowser.Document;
                HtmlElement institution = acgmedocument.GetElementById("Institutions");
                HtmlElement attending = acgmedocument.GetElementById("Attendings");
                HtmlElement adultorped = acgmedocument.GetElementById("PatientTypes");
                HtmlElement gender = acgmedocument.GetElementById("Genders");
                HtmlElement date = acgmedocument.GetElementById("ProcedureDate");
                HtmlElement residentroles = acgmedocument.GetElementById("ResidentRoles");
                HtmlElement comments = acgmedocument.GetElementById("Comments");
                HtmlElement patientid = acgmedocument.GetElementById("PatientId");
                HtmlElement residentyear = acgmedocument.GetElementById("ProcedureYear");
                HtmlElement cptid = acgmedocument.GetElementById("Code");

                institution.SetAttribute("value", "13233");
                //              attending.SetAttribute("value", "26518");
                adultorped.SetAttribute("value", "69");
                residentyear.SetAttribute("value", residentpgyyear);
                residentroles.SetAttribute("value", residentrole);


                double d_date = double.Parse(myexcelfile2.GetCellValue("a" + j.ToString()));
                DateTime conv = DateTime.FromOADate(d_date);
                date.SetAttribute("value", conv.ToShortDateString());
                gender.SetAttribute("value", myexcelfile2.GetCellValue("f" + j.ToString()));
                comments.SetAttribute("value", myexcelfile2.GetCellValue("g" + j.ToString()));
                patientid.SetAttribute("value", myexcelfile2.GetCellValue("d" + j.ToString()));
                cptid.SetAttribute("value", myexcelfile2.GetCellValue("j" + j.ToString()));
                string attendingsurgeon = myexcelfile2.GetCellValue("i" + j.ToString());
                string attendingcode = "26518";
                switch (attendingsurgeon)
                {
                    case ("Jea"):
                        attendingcode = "6149";
                        break;
                    case ("Dauser"):
                        attendingcode = "49221";
                        break;
                    case ("Whitehead"):
                        attendingcode = "28279";
                        break;
                    case ("Bollo"):
                        attendingcode = "165100"; //this is Leurssen, Bollo not available
                        break;
                    case ("Curry"):
                        attendingcode = "19065";
                        break;
                    case ("Yoshor"):
                        attendingcode = "17692";
                        break;
                    case ("Ehni"):
                        attendingcode = "60345";
                        break;
                    case ("Gopinath"):
                        attendingcode = "26518";
                        break;
                    case ("Patel"):
                        attendingcode = "359394";
                        break;
                    case ("Fuentes"):
                        attendingcode = "353465";
                        break;
                    case ("Ropper"):
                        attendingcode = "385016";
                        break;
                }
                attending.SetAttribute("value", attendingcode);
                cptid.SetAttribute("value", myexcelfile2.GetCellValue("j" + j.ToString()));
                acgmedocument.GetElementById("searchByCodeButton").InvokeMember("click");
                findelementbytextname("Code", acgmedocument).InvokeMember("click");
            
            //Sen Mode Activated 
                if (issenmodeactivated)
                {
                    whereami = "addingcptcode";
                    newtimer = new System.Windows.Forms.Timer();
                    newtimer.Interval = 1000;
                    newtimer.Start();
                    newtimer.Tick += new EventHandler(newtimer_Tick);
                }
        }

        void newtimer_Tick(object sender, EventArgs e)
        {
            if (whereami == "addingcptcode")
            {
                findbuttonbytextname("Add", acgmeWebBrowser.Document).InvokeMember("click");
                //      whereami = "cancontinue";
                whereami = "cptcodeadded";

            }
            if (whereami == "cptcodeadded")
            {
                acgmeWebBrowser.Document.GetElementById("submitButton").InvokeMember("click");
                whereami = "next";
            }

            if (whereami == "next")
            {
                newtimer.Stop();
                whereami = "cancontinue";                            
            }
        }


        private void convertXPSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openXPSFileDialog.ShowDialog();
            string filename = openXPSFileDialog.FileName;
            xpsparser myparserx = new xpsparser(filename);

            
            myexcelfile = new excelfile(addressofexcelfile);
            
            
            string outputstring = myparserx.extracttext();
            int startofdata = outputstring.IndexOf("In") + 2;
            outputstring = outputstring.Substring(startofdata);
            string[] stringSeparators = new string[] {"Conf"};
            string[] lines = outputstring.Split(stringSeparators, StringSplitOptions.None);
            int i = 0;
            int excelindex = 0;
            string proceduredate;
            string age;
            string sex;
            string mrn;
            string procedurename;
            string procedurecode;
            string attending;
            string lastname;
            string firstname;
            while (i < lines.Length)
            {
                string temp = lines[i].Replace(" ","");
                if (temp.Contains("Printed"))
                {
                    int indexofpage = temp.IndexOf("Page");
                    int indexofend = temp.LastIndexOf(":") + 5;
                    temp = temp.Remove(indexofpage, indexofend - indexofpage);
                }
                if (temp.Contains(residentname))
                {
                    excelindex = excelindex + 1;
//                    int endofdate = temp.LastIndexOf("/")+5;
                    int endofdate = temp.IndexOf("/", 3) + 5;
                    int endofage = temp.IndexOf("y.o.");
                    proceduredate = temp.Substring(0,endofdate);
                    age = temp.Substring(endofdate,endofage-endofdate);
                    sex = temp.Substring(endofage+4,1);
                    attending = "";
                    int endnameindex = temp.IndexOf("&lt;");
                    int endmrnindex = temp.IndexOf("&gt;");
                    int endprocedureinex = temp.IndexOf("]")+1;
                    int begprocedurecode = temp.IndexOf("[");
                    int indexofcomma = temp.IndexOf(",");
                    lastname = temp.Substring(endofage+5, indexofcomma - endofage-5);
                    firstname = temp.Substring(indexofcomma + 1, endnameindex - indexofcomma - 1); 
                    mrn = temp.Substring(endnameindex + 4, endmrnindex - endnameindex - 4);
                    procedurename = temp.Substring(endmrnindex + 4, endprocedureinex - endmrnindex - 4);
                    procedurecode = temp.Substring(begprocedurecode+1, endprocedureinex - begprocedurecode - 2);
                    if (temp.Contains("Ehni"))
                        attending = "Ehni";
                    else if (temp.Contains("Ropper"))
                        attending = "Ropper";
                    else if (temp.Contains("Fuentes"))
                        attending = "Fuentes";
                    else if (temp.Contains("Patel"))
                        attending = "Patel";
                    else if (temp.Contains("Omeis"))
                        attending = "Omeis";
                    else if (temp.Contains("Duckworth"))
                        attending = "Duckworth";
                    else
                        attending = "Gopinath";
                    myexcelfile.WriteCell("a" + excelindex.ToString(), proceduredate);
                    myexcelfile.WriteCell("b" + excelindex.ToString(), lastname);
                    myexcelfile.WriteCell("c" + excelindex.ToString(), firstname);
                    myexcelfile.WriteCell("d" + excelindex.ToString(), mrn);
                    myexcelfile.WriteCell("e" + excelindex.ToString(), age);
                    myexcelfile.WriteCell("f" + excelindex.ToString(), sex);
                    myexcelfile.WriteCell("g" + excelindex.ToString(), procedurename);
                    myexcelfile.WriteCell("h" + excelindex.ToString(), procedurecode);
                    myexcelfile.WriteCell("i" + excelindex.ToString(), attending);
                    totalnumberofcases = totalnumberofcases + 1;
                }
                i++;
            }
            label1.Visible = true;
            label1.Text = "Total number of Cases: " + totalnumberofcases.ToString();
           // myexcelfile.savefile();
            myexcelfile.closefile();
        }
        private HtmlElement findelementbytextname(string text, HtmlDocument inputhtmldoc)
        {
            for (int linksindex = 0; linksindex < inputhtmldoc.Links.Count; linksindex++)
            {
                if (inputhtmldoc.Links[linksindex].InnerText.Contains(text))
                {
                    return inputhtmldoc.Links[linksindex];
                }
            }
            return null;
        }


        private HtmlElement findbuttonbytextname(string text, HtmlDocument inputhtmldoc)
        {
            for (int buttonindex = 0; buttonindex < inputhtmldoc.GetElementsByTagName("button").Count; buttonindex++)
            {
                    if (inputhtmldoc.GetElementsByTagName("button")[buttonindex].InnerHtml.Contains(text))
                    {
                        return inputhtmldoc.GetElementsByTagName("button")[buttonindex];
                    }
            }
            return null;
        }

        public string Name
        {
            get
            {
                return residentname;
            }
            set
            {
                residentname = value;
            }
        }

        private void openExcelStripMenuItem_Click_1(object sender, EventArgs e)
        {
            addressofexcelfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\cases" + DateTime.Now.ToShortDateString().Replace("/", "_") + ".xlsx";
            if (!File.Exists(addressofexcelfile))
            {
                excelfile myexcelfile2 = new excelfile(addressofexcelfile);
                //     totalnumberofcases = myexcelfile2.getlastrow();
                myexcelfile2.savefile();
                myexcelfile2.closefile();
            }
            else
            {
                excelfile myexcelfile2 = new excelfile(addressofexcelfile);
                totalnumberofcases = myexcelfile2.getlastrow();
                label1.Text = "Total number of Cases: " + totalnumberofcases.ToString();
                label1.Visible = true;
             //   myexcelfile2.savefile();
                myexcelfile2.closefile();
            }
        }

        private void nameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            nameform.Show();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            residentpgyyear = "1";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Red;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            residentpgyyear = "2";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Red;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            residentpgyyear = "3";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Red;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            residentpgyyear = "4";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Red;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            residentpgyyear = "5";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Red;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            residentpgyyear = "6";
            toolStripMenuItem8.ForeColor = Color.Black;
            toolStripMenuItem7.ForeColor = Color.Red;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            residentpgyyear = "7";
            toolStripMenuItem8.ForeColor = Color.Red;
            toolStripMenuItem7.ForeColor = Color.Black;
            toolStripMenuItem6.ForeColor = Color.Black;
            toolStripMenuItem5.ForeColor = Color.Black;
            toolStripMenuItem4.ForeColor = Color.Black;
            toolStripMenuItem3.ForeColor = Color.Black;
            toolStripMenuItem2.ForeColor = Color.Black;
        }

        private void seniorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            residentrole = "39";
            seniorToolStripMenuItem.ForeColor = Color.Red;
            leadToolStripMenuItem.ForeColor = Color.Black;
            assistantToolStripMenuItem.ForeColor = Color.Black;
        }

        private void leadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            residentrole = "38";
            seniorToolStripMenuItem.ForeColor = Color.Black;
            leadToolStripMenuItem.ForeColor = Color.Red;
            assistantToolStripMenuItem.ForeColor = Color.Black;
        }

        private void assistantToolStripMenuItem_Click(object sender, EventArgs e)
        {
            residentpgyyear = "37";
            seniorToolStripMenuItem.ForeColor = Color.Black;
            leadToolStripMenuItem.ForeColor = Color.Black;
            assistantToolStripMenuItem.ForeColor = Color.Red;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Created by GDR");
        }

        private void sENMODEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (issenmodeactivated == false)
            {
                issenmodeactivated = true;
                senmodelabel.Visible = true;
                submitToolStripMenuItem.Text = "Run";
            }
            else
            {
                issenmodeactivated = false;
                senmodelabel.Visible = false;
                submitToolStripMenuItem.Text = "Next";
            }
        }

    }
}
