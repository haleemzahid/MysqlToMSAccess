using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        OdbcConnection con = new OdbcConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\DB.mdb");


        string myConnectionString = @"Driver={Microsoft Access Driver (*.mdb)};" + "Dbq=DB.mdb;";

        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnSyncedClicked(object sender, RoutedEventArgs e)
        {
                    InsertIntoMSAccess();

        }

        public OdbcConnection OpenMSAccessConnection()
        {
            
            var myConnection = new OdbcConnection();
                myConnection.ConnectionString = myConnectionString;
                myConnection.Open();
            return myConnection;
            

            }
        

        public void InsertIntoMSAccess()
        {
            var con = OpenMSAccessConnection();
            
            var a = new Donnees_Dossiers();
            a.dordre_adresse = "asdas";
            a.dordre_cp = "asdas";
            a.dordre_Entete = "asdas";
            a.dordre_nom = "asdas";
            a.dordre_tel = "asdas";

            OdbcCommand cmd = con.CreateCommand();
            cmd.CommandText = "INSERT INTO Donnees_Dossiers(dordre_adresse,dordre_cp,dordre_Entete,dordre_nom,dordre_tel)VALUES('" + a.dordre_adresse + "','" + a.dordre_cp + "','" + a.dordre_Entete + "','"+a.dordre_nom+"','"+a.dordre_tel+"')";
              
                
                cmd.ExecuteNonQuery();
            con.Close();
            

        }
    }
}
