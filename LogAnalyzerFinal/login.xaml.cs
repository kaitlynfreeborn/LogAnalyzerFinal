using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace LogAnalyzerFinal
{
    /// <summary>
    /// Interaction logic for login.xaml
    /// </summary>
    public partial class login : Window
    {
        public login()
        {
            InitializeComponent();
        }

        // compares username and password to allow user into program
        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            string userName, password;
            userName = txtUsername.Text;
            password = Password.Password;


            if (userName == "admin" && password == "admin")
            {
                MessageBox.Show("Welcome " + userName);
                this.Hide();
                this.Close();

            }
            else
            {
                MessageBox.Show("Username/Password does not exist.");
            }

        }
    }
}
