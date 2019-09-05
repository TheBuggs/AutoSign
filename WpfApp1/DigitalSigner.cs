using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using List = Microsoft.Office.Interop.Word.List;
using MessageBox = System.Windows.MessageBox;
using Timer = System.Timers.Timer;

namespace WpfApp1
{
    class DigitalSigner
    {
        private Microsoft.Office.Interop.Word._Application oWord;
        private Microsoft.Office.Interop.Word._Document oDoc;
        private string username;
        private string password;
        private string email;

        public DigitalSigner(ref Microsoft.Office.Interop.Word._Application oWord, string username, string password, string email)
        {
            this.oWord = oWord;
            this.username = username;
            this.password = password;
            this.email = email;
        }


        public void SignWord()
        {
           


            object sigID = "{00000000-0000-0000-0000-000000000000}";

            Timer timer = new Timer();
            timer.Elapsed += (s, args) =>
            {
                SendKeys.SendWait("{TAB}");
                SendKeys.SendWait("~");
                timer.Stop();
            };

            timer.Interval = 1000;
            timer.Start();

           
            try
            {
                oWord.Activate();

                SignatureSet signatureSet = oWord.ActiveDocument.Signatures;
                signatureSet.ShowSignaturesPane = false;
                Signature objSignature = signatureSet.AddSignatureLine(sigID);
                objSignature.Setup.SuggestedSigner = this.username;
                objSignature.Setup.SuggestedSignerEmail = this.email;
                objSignature.Setup.ShowSignDate = true;

                oWord.Documents.Save();

                Timer t1 = new Timer();
                Timer t2 = new Timer();
                t1.Elapsed += (st, args) =>
                {
                    SendKeys.SendWait(" ");
                    SendKeys.SendWait("{ENTER}");
                    t2.Start();
                    t1.Stop();
                };
                t1.Interval = 2000;
                
                t2.Elapsed += (st1, args1) =>
                {
                    int i = 0;
                    while (i < this.password.Length)
                    {
                        SendKeys.SendWait(this.password[i].ToString());
                        i++;
                    }

                    SendKeys.SendWait("~");

                    t2.Stop();
                };
                t2.Interval = 3000;
                t1.Start();

                objSignature.Sign();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //oWord.Documents.Save();
            //oWord.Quit();

            try
            {
                Marshal.ReleaseComObject(oWord);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
    }
}
