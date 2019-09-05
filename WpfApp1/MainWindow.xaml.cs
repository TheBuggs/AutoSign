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
using MessageBox = System.Windows.MessageBox;
using Timer = System.Timers.Timer;
using FORMS = System.Windows.Forms;
namespace WpfApp1
{
   
    public partial class MainWindow : System.Windows.Window
    {
        double t1 = 2000;
        double t2 = 2000;
        double t3 = 3000;
        public MainWindow()
        {
            InitializeComponent();

            txtTimer1.Text = t1.ToString();
            txtTimer2.Text = t2.ToString();
            txtTimer3.Text = t3.ToString();
        }

        private void Sign_OnClick(object sender, RoutedEventArgs e)
        {
            string position = this.position.Text;
            string email = this.email.Text;
            string username = this.username.Text;
            string path = String.Format(@"{0}", this.path.Text);
            string password = this.password.Password;
            string repeater = this.repeater.Password;


            if (System.IO.Directory.Exists(path))
            { 
                FileExplorer f = new FileExplorer(path);
                if (password.Equals(repeater))
                {
                    f.GetDocumentsSign(username, email, password, position, Double.Parse(txtTimer1.Text), Double.Parse(txtTimer2.Text), Double.Parse(txtTimer3.Text));
                }
                else
                {
                    MessageBox.Show("Паролите не съвпадат!\n");
                }
            }
            else
            {
                MessageBox.Show(String.Format("Няма такъв път: {0}\n", path));
            }
        }

        private void Path_search_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            FORMS.DialogResult result = dialog.ShowDialog();
            
            if (result == FORMS.DialogResult.OK)
            {
                path.Text = dialog.SelectedPath; 
            }
            return;
        }
    }
}
