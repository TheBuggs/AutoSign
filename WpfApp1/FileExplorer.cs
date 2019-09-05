using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Application = System.Windows.Forms.Application;
using MessageBox = System.Windows.MessageBox;
using Timer = System.Timers.Timer;

namespace WpfApp1
{
   
    class FileExplorer
    {
        private string path;

        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("USER32.DLL", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern long GetClassName(IntPtr hWnd, StringBuilder ipClassName, long nMaxCount);

        // This is Window Class for Word Application
        private const string wordClassName = "OpusApp";
        public FileExplorer(string path)
        {
            this.path = path;
        }

        public void GetDocumentsSign(string username, string email, string password, string position, double timer1, double timer2, double timer3)
        {
            
      
            string[] files = Directory.GetFiles(this.path);
            
            object oEndOfDoc = "\\endofdoc";
            object oMissing = System.Reflection.Missing.Value;

            int counter = 0;

            Microsoft.Office.Interop.Word._Application oWord;

            Microsoft.Office.Interop.Word._Document oDoc;

            oWord = new Microsoft.Office.Interop.Word.Application();

            oWord.Visible = true;
            oWord.WindowState = WdWindowState.wdWindowStateMaximize;

            IntPtr wordHandle = FindWindow(FileExplorer.wordClassName, oWord.Application.Caption);

            if (wordHandle == IntPtr.Zero)
            {
               
                MessageBox.Show("Word is not running!");
             
                CloseApplication(oWord);
                return;
            }
           
            foreach (var file in files)
            {
                
                oDoc = oWord.Documents.Open(file, ReadOnly: false);
                oDoc.ActiveWindow.WindowState = WdWindowState.wdWindowStateMaximize;
                
                oDoc.Final = false;

                if (wordHandle == IntPtr.Zero)
                {
                    MessageBox.Show("Word is not running!");

                    CloseApplication(oWord);
                    return;
                }
                else
                {
                    SetForegroundWindow(wordHandle);
                }

                
   
                object paramNextPage = Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
                oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertBreak(ref paramNextPage);
                object breakPage = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                object saveOption = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
                object routeDocument = false;

                object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine;
                object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToLast;
                object count = 4;

                Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
                para.Range.Text = "С уважение!";
                para.Range.Font.Size = 12;
                para.Range.Font.Name = "Times New Roman";
               

                para.Range.InsertParagraphAfter();

                oWord.Selection.GoTo(ref what, ref which, ref count, ref oMissing);

               // oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertBreak(ref paramNextPage);
               // oWord.Selection.GoTo(ref what, ref which, ref count, ref oMissing);

                object sigID = "{00000000-0000-0000-0000-000000000000}";

                Timer timer = new Timer();
                timer.Elapsed += (s, args) =>
                {

                    SendKeys.SendWait("{TAB}");

                    SendKeys.SendWait("{ENTER}");

                    timer.Stop();
                };

                timer.Interval = timer1;
                timer.Start();

                try
                {
                    oWord.Activate();

                    
                    Microsoft.Office.Core.SignatureSet signatureSet = oWord.ActiveDocument.Signatures;
                    signatureSet.ShowSignaturesPane = false;
                    Signature objSignature = signatureSet.AddSignatureLine(sigID);
                    objSignature.Setup.SuggestedSigner = username;
                    objSignature.Setup.SuggestedSignerLine2 = position;
                    objSignature.Setup.SuggestedSignerEmail = email;
                    objSignature.Setup.ShowSignDate = true;

                    

                    

                    oWord.Documents.Save();
                    
                    Timer t1 = new Timer();
                    Timer t2 = new Timer();
                    if (counter == 0)
                    {
                        t1.Elapsed += (st, args) =>
                        {
                            oWord.Activate();
                            SendKeys.SendWait(" ");
                            SendKeys.SendWait("{ENTER}");
                            t2.Start();
                            t1.Stop();
                        };

                        t2.Elapsed += (st1, args1) =>
                        {
                            oWord.Activate();
                            int i = 0;
                            while (i < password.Length)
                            {
                                oWord.Activate();
                                SendKeys.SendWait(password[i].ToString());
                                i++;
                            }

                            SendKeys.SendWait("~");

                            t2.Stop();
                        };

                        t1.Interval = timer2;
                        t2.Interval = t1.Interval + timer3;
                        t1.Start();
                    }
                    else
                    {
                        t1.Elapsed += (st, args) =>
                        {
                            oWord.Activate();
                            SendKeys.SendWait(" ");
                            SendKeys.SendWait("{ENTER}");
                            t1.Stop();
                        };
                        t1.Interval = timer2;
                        t1.Start();

                    }
                    try
                    {
                        objSignature.Sign();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Документ №{counter} не бе разписан!\n " + ex.ToString());
                        oDoc.Close();
                        oWord.Quit();
                        
                        if (oDoc != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                        if (oWord != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
                        MessageBox.Show($"Фатална грешка при обработка на документ №{counter}\n" + ex.ToString());
                       // GC.Collect();
                        return;
                    }

                }
                catch (Exception ex)
                {

                    CloseDocument(oDoc);
                    CloseApplication(oWord);
                    MessageBox.Show($"Фатална грешка при обработка на документ №{counter}\n" + ex.ToString());
                    return; 
                }
                oDoc.Close();
                if (oDoc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                counter++;
                oDoc = null;
               // GC.Collect();
            }

            CloseApplication(oWord);

            if (counter > 0)
            {
                string msg = String.Format("Всички {0} документа бяха успешно подписани!\n", counter);
                MessageBox.Show(msg);
            }          
        }

        private static void CloseDocument(_Document oDoc)
        {
            oDoc.Close();
            if (oDoc != null)
            {
                try
                {
                    oDoc.Activate();
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc) > 0) ;
                }
                catch { }
                finally
                {
                    oDoc = null;
                }
            }


        }
        private static void CloseApplication(_Application oWord)
        {
            try
            {
                oWord.Activate();
                oWord.Quit();
                if (oWord != null)
                {
                    try
                    {
                        oWord.Activate();
                        while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oWord) > 0) ;
                    }
                    catch { }
                    finally
                    {
                        oWord = null;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error to close WORD Application\n");
            }
        }
    }
}
