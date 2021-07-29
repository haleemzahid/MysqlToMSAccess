using MySql.Data.MySqlClient;
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
        List<Donees> mysqllist = new List<Donees>();
        List<Donees> Checkedl = new List<Donees>();
       
        private  MySqlConnection connection;
        List<Donnees_Dossiers> DataToInsertList = new List<Donnees_Dossiers>();
        string myConnectionString = @"Driver={Microsoft Access Driver (*.mdb)};" + "Dbq=DB.mdb;";

        public MainWindow()
        {
            InitializeComponent();
            dg.ItemsSource =mysqllist = GetMySqlList();
            MsAccessDG.ItemsSource = GetMSAccessList();

        }
        private void btnSyncedClicked(object sender, RoutedEventArgs e)
        {
            FilterData();
            GetMySqlDataIntoMSAccess();
            if (Checkedl.Count==0&&mysqllist.Count>0)
            {

                MessageBox.Show("Please select data to sync");
                    return;
            }
            if (DataToInsertList.Count == 0&&mysqllist.Count==0)
            {
                MessageBox.Show("No data in MySql Database to insert into MS Access");
                    return;
            }
            else if(DataToInsertList.Count == 0)
            {
                MessageBox.Show("Data Is already synced");
                return;
            }
            MessageBox.Show(DataToInsertList.Count + " rows affected out of "+Checkedl.Count);

            MsAccessDG.ItemsSource = GetMSAccessList();

        }

        public OdbcConnection OpenMSAccessConnection()
        {
            
            var myConnection = new OdbcConnection();
                myConnection.ConnectionString = myConnectionString;
                
            return myConnection;
            

            }

        public List<Donees> GetMySqlList()
        {
            List<Donees> l = new List<Donees>();
            var con = OpenMySqlConnection();
            string str = "SELECT * FROM Donnees_Dossiers";

            MySqlDataReader mySqlDataReader = (new MySqlCommand(str, con)).ExecuteReader();
            while (mySqlDataReader.Read())
            {
                
                l.Add(new Donees()
                {
                    IsChecked = false,
                    donnees_Dossiers = new Donnees_Dossiers()
                    {



                        Id = mySqlDataReader.GetInt32(0),
                        Num_devis_numero = mySqlDataReader.GetString(1),
                        Num_dossier = mySqlDataReader.GetString(2),
                        Num_dossier_lié = mySqlDataReader.GetString(3),
                        dordre_type = mySqlDataReader.GetString(4),
                        dordre_Entete = mySqlDataReader.GetString(5),
                        dordre_nom = mySqlDataReader.GetString(6),
                        dordre_adresse = mySqlDataReader.GetString(7),
                        dordre_cp = mySqlDataReader.GetString(8),
                        dordre_ville = mySqlDataReader.GetString(9),
                        dordre_tel = mySqlDataReader.GetString(10),
                        dordre_fax = mySqlDataReader.GetString(11),
                        dordre_mail = mySqlDataReader.GetString(12),
                        proprietaire_Entete = mySqlDataReader.GetString(13),
                        proprietaire_nom = mySqlDataReader.GetString(14),
                        proprietaire_adresse = mySqlDataReader.GetString(15),
                        proprietaire_cp = mySqlDataReader.GetString(16),
                        proprietaire_ville = mySqlDataReader.GetString(17),
                        proprietaire_tel = mySqlDataReader.GetString(18),
                        proprietaire_fax = mySqlDataReader.GetString(19),
                        proprietaire_mail = mySqlDataReader.GetString(20),
                        bien_adresse = mySqlDataReader.GetString(21),
                        bien_cp = mySqlDataReader.GetString(22),
                        bien_ville = mySqlDataReader.GetString(23),
                        bien_lieu_interne = mySqlDataReader.GetString(24),
                        bien_cadastre = mySqlDataReader.GetString(25),
                        bien_lot = mySqlDataReader.GetString(26),
                        bien_lot_cave_cellier = mySqlDataReader.GetString(27),
                        bien_lot_parking_garage = mySqlDataReader.GetString(28),
                        bien_lot_autre = mySqlDataReader.GetString(29),
                        bien_surface_terrain = mySqlDataReader.GetString(30),
                        bien_année_construction = mySqlDataReader.GetString(31),
                        bien_parcelle = mySqlDataReader.GetString(32),
                        bien_nature = mySqlDataReader.GetString(33),
                        bien_IGH_ERP = mySqlDataReader.GetString(34),
                        bien_description = mySqlDataReader.GetString(35),
                        rdv_date = mySqlDataReader.GetString(36),
                        rdv_heure = mySqlDataReader.GetString(37),
                        rdv_duree = mySqlDataReader.GetString(38),
                        rdv_contact_nom_tel = mySqlDataReader.GetString(39),
                        rdv_precisions = mySqlDataReader.GetString(40),
                        rdv_clefs = mySqlDataReader.GetString(41),
                        dossier_Acces = mySqlDataReader.GetString(42),
                        dossier_Nom = mySqlDataReader.GetString(43),
                        dossier_Acces_relatif = mySqlDataReader.GetString(44),
                        dossier_Archive = mySqlDataReader.GetString(45),
                        dossier_clot = mySqlDataReader.GetInt32(46),
                        dossier_etat_rapport = mySqlDataReader.GetInt32(47),
                        dossier_etat_paie = mySqlDataReader.GetString(48),
                        dossier_observations = mySqlDataReader.GetString(49),
                        rapport_date = mySqlDataReader.GetString(50),
                        rapport_date_envoyee = mySqlDataReader.GetString(51),
                        rapport_destinataires = mySqlDataReader.GetString(52),
                        rapport_facturation = mySqlDataReader.GetString(53),
                        rapport_type = mySqlDataReader.GetString(54),
                        rapport_amiante_FCFP = mySqlDataReader.GetInt32(55),
                        rapport_amiante_Autres = mySqlDataReader.GetInt32(56),
                        rapport_termites_resultat = mySqlDataReader.GetInt32(57),
                        notaire_Entete = mySqlDataReader.GetString(58),
                        notaire_nom = mySqlDataReader.GetString(59),
                        notaire_adresse = mySqlDataReader.GetString(60),
                        notaire_cp = mySqlDataReader.GetString(61),
                        notaire_ville = mySqlDataReader.GetString(62),
                        notaire_tel = mySqlDataReader.GetString(63),
                        notaire_fax = mySqlDataReader.GetString(64),
                        notaire_mail = mySqlDataReader.GetString(65),
                        bien_description_cases = mySqlDataReader.GetString(66),
                        bien_perimetre = mySqlDataReader.GetString(67),
                        rapport_destinataires_mail = mySqlDataReader.GetString(68),

                        dossier_etat = mySqlDataReader.GetString(69),
                        complement_visite = mySqlDataReader.GetString(70),
                        operateur_reperage = mySqlDataReader.GetString(71),
                        photo_de_presentation = mySqlDataReader.GetString(72),
                        facturation_restante = mySqlDataReader.GetInt32(73),
                        facturation_compte_client = mySqlDataReader.GetString(74),
                        facturation_remise_globale = mySqlDataReader.GetInt32(75),
                        facturation_date = mySqlDataReader.GetString(76),
                        facturation_date_fin = mySqlDataReader.GetString(77),
                        Donnee1 = mySqlDataReader.GetString(78),
                        Donnee2 = mySqlDataReader.GetString(79),
                        Donnee3 = mySqlDataReader.GetString(80),
                        Donnee4 = mySqlDataReader.GetString(81),
                        Donnee5 = mySqlDataReader.GetString(82),
                        Donnee6 = mySqlDataReader.GetString(83),
                        Donnee7 = mySqlDataReader.GetString(84),
                        Donnee8 = mySqlDataReader.GetString(85),
                        Donnee9 = mySqlDataReader.GetString(86),
                        Donnee10 = mySqlDataReader.GetString(87),
                        Donnee11 = mySqlDataReader.GetString(88),
                        Donnee12 = mySqlDataReader.GetString(89),
                        Donnee13 = mySqlDataReader.GetString(90),
                        Donnee14 = mySqlDataReader.GetString(91),
                        Donnee15 = mySqlDataReader.GetString(92),
                        Donnee16 = mySqlDataReader.GetString(93),
                        Donnee17 = mySqlDataReader.GetString(94),
                        Donnee18 = mySqlDataReader.GetString(95),
                        Donnee19 = mySqlDataReader.GetString(96),
                        Mission_Memo = mySqlDataReader.GetString(97),
                        operateur_certif_num = mySqlDataReader.GetString(98),
                        operateur_certif_societe = mySqlDataReader.GetString(99),
                        operateur_certif_date = mySqlDataReader.GetString(100),
                        Mode_Access = mySqlDataReader.GetString(101),
                        Date_commande = mySqlDataReader.GetString(102),
                        Signature_Opérateur = mySqlDataReader.GetString(103),
                        Id_facturation = mySqlDataReader.GetString(104),
                        Appareil_CREP = mySqlDataReader.GetString(105),
                        Date_RDV = mySqlDataReader.GetInt32(106),
                        Facture_validation = mySqlDataReader.GetInt32(107),
                        Paiement_validation = mySqlDataReader.GetInt32(108),
                        rapport_plus = mySqlDataReader.GetString(109),
                        Commerciaux = mySqlDataReader.GetString(110),
                        Commerciaux_autre = mySqlDataReader.GetString(111),
                        Certif_obtention = mySqlDataReader.GetString(112),
                        Date_1er_paiement = mySqlDataReader.GetString(113),
                        id_dossier_liciweb = mySqlDataReader.GetString(114),
                        id_donneur_ordre = mySqlDataReader.GetString(115),
                        conclusion = mySqlDataReader.GetString(116),
                        valeur_bien = mySqlDataReader.GetString(117),
                        rapport_type_expertise = mySqlDataReader.GetString(118),
                        rapport_plus_expertise = mySqlDataReader.GetString(119),
                        Adresse_web_dossier = mySqlDataReader.GetString(120),
                        id_dossier_licielweb = mySqlDataReader.GetString(121),
                        type_de_dossier = mySqlDataReader.GetString(122),
                        etat_licielweb = mySqlDataReader.GetString(123),
                        DATE_RDV_facture = mySqlDataReader.GetInt32(124),
                        Date_paiement_codee = mySqlDataReader.GetInt32(125),
                        AR_Amiante = mySqlDataReader.GetString(126),
                        Envoie_ADEME = mySqlDataReader.GetString(127),
                        backoffice = mySqlDataReader.GetString(128),
                        date_modification = mySqlDataReader.GetString(129),
                        date_derniere_sauvegarde = mySqlDataReader.GetInt32(130),
                    }
                }


                   );
            }                                                 
            mySqlDataReader.Close();
            con.Close();
            return l;
        }
        public List<Donnees_Dossiers> GetMSAccessList()
        {
            List<Donnees_Dossiers> l = new List<Donnees_Dossiers>();
            var con = OpenMSAccessConnection();
            con.Open();
            OdbcCommand cmd = con.CreateCommand();
            cmd.CommandText = "SELECT * FROM `Donnees_Dossiers`";
            OdbcDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection); // close conn after complete

            while (reader.Read())
            {
                string Num_devis_numero="", Num_dossier = "", Num_dossier_lié = "", dordre_nom = "", dordre_mail = "";
                int a = -1;
                Num_devis_numero = (!Convert.IsDBNull(1) ? reader.GetString(1) : "");
                Num_dossier = (!Convert.IsDBNull(2) ? reader.GetString(2) : "");
                Num_dossier_lié = (!Convert.IsDBNull(3) ? reader.GetString(3) : "");
                dordre_nom = (!Convert.IsDBNull(6) ? reader.GetString(6) : "");
                dordre_mail = (!Convert.IsDBNull(1) ? reader.GetString(12) : "");
                a = (!Convert.IsDBNull(1) ? reader.GetInt16(131) : -1);
                l.Add(new Donnees_Dossiers()
                {


                    Id = Convert.ToInt32(reader.GetInt32(0)),

                    Num_devis_numero = Num_devis_numero,
                    Num_dossier = Num_dossier,
                    Num_dossier_lié = Num_dossier_lié,

                    dordre_nom = dordre_nom,
                    dordre_mail = dordre_mail,
                    MySqlid = a

                }

                   ); ;
            }
            reader.Close();
            return l;
        }





        public MySqlConnection OpenMySqlConnection()
        {
            var builder = new MySqlConnectionStringBuilder
            {
                Server = "localhost",
                Database = "db",
                UserID = "root",
                Password = "abc123",
                SslMode = MySqlSslMode.None,
            };
               connection = new MySqlConnection(builder.ConnectionString);
            connection.Open();
            return connection;
        }             
                                                                    
     
        public void FilterData()
        {

            DataToInsertList = new List<Donnees_Dossiers>();
            int count=0;
            Checkedl = mysqllist.Where(x => x.IsChecked == true).ToList() ;
            var MSAccessData = GetMSAccessList();
            foreach (var item in Checkedl)
            {
                foreach (var item2 in MSAccessData)
                {
                    if (item2.MySqlid == item.donnees_Dossiers.Id)
                    {

                        count = 1;
                        break;
                    }

                }
                if (count != 1)
                {
                    DataToInsertList.Add(item.donnees_Dossiers);
                    
                        }
                else
                {
                    count = 0;
                }
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var a = ((sender as dynamic).DataContext as Donees);
            a.IsChecked = true;

        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            var a = ((sender as dynamic).DataContext as Donees);
            a.IsChecked = false;
        }
        public void GetMySqlDataIntoMSAccess()
        {
                var con = OpenMSAccessConnection();
            con.Open();
            foreach (var item in DataToInsertList)
            {

              


                OdbcCommand cmd = con.CreateCommand();
                cmd.CommandText = "INSERT INTO Donnees_Dossiers(Num_devis_numero,Num_dossier,Num_dossier_lié,dordre_type,dordre_Entete,dordre_nom,dordre_adresse,dordre_cp,dordre_ville,dordre_tel,dordre_fax,dordre_mail,proprietaire_Entete," +
                    "proprietaire_nom,proprietaire_adresse,proprietaire_cp,proprietaire_ville,proprietaire_tel,proprietaire_fax,proprietaire_mail,bien_adresse" +
                    ",bien_cp,bien_ville,bien_lieu_interne,bien_cadastre,bien_lot,bien_lot_cave_cellier,bien_lot_parking_garage," +
                    "bien_lot_autre,bien_surface_terrain,bien_année_construction,bien_parcelle,bien_nature,bien_IGH_ERP,bien_description" +
                    ",rdv_date,rdv_heure,rdv_duree,rdv_contact_nom_tel,rdv_precisions,rdv_clefs,dossier_Acces,dossier_Nom,dossier_Acces_relatif" +
                    ",dossier_Archive,dossier_clot,dossier_etat_rapport,dossier_etat_paie,dossier_observations,rapport_date,rapport_date_envoyee" +
                    ",rapport_destinataires,rapport_facturation,rapport_type,rapport_amiante_FCFP,rapport_amiante_Autres,rapport_termites_resultat" +
                    ",notaire_Entete,notaire_nom,notaire_adresse,notaire_cp,notaire_ville,notaire_tel,notaire_fax,notaire_mail,bien_description_cases" +
                    ",bien_perimetre,rapport_destinataires_mail,dossier_etat,complement_visite,operateur_reperage,photo_de_presentation,facturation_restante" +
                    ",facturation_compte_client,facturation_remise_globale,facturation_date,facturation_date_fin,Donnee1,Donnee2,Donnee3,Donnee4" +
                    ",Donnee5,Donnee6,Donnee7,Donnee8,Donnee9,Donnee10,Donnee11,Donnee12,Donnee13,Donnee14,Donnee15,Donnee16,Donnee17,Donnee18,Donnee19" +
                    ",Mission_Memo,operateur_certif_num,operateur_certif_societe,operateur_certif_date,Mode_Access,Date_commande,Signature_Opérateur" +
                    ",Id_facturation,Appareil_CREP,Date_RDV,Facture_validation,Paiement_validation,rapport_plus,Commerciaux,Commerciaux_autre,Certif_obtention" +
                    ",Date_1er_paiement,id_dossier_liciweb,id_donneur_ordre,conclusion,valeur_bien,rapport_type_expertise,rapport_plus_expertise" +
                    ",Adresse_web_dossier,id_dossier_licielweb,type_de_dossier,etat_licielweb,DATE_RDV_facture,Date_paiement_codee,AR_Amiante" +
                    ",Envoie_ADEME,backoffice,date_modification,date_derniere_sauvegarde,MySqlid) VALUES('" + item.Num_devis_numero+"','"+item.Num_dossier+"','"+item.Num_dossier_lié+"','"+item.dordre_type+"','"+item.dordre_Entete+"','"+item.dordre_nom+"','"+item.dordre_adresse+"','"+item.dordre_cp+"','"+item.dordre_ville+"','"+item.dordre_tel+"','"+item.dordre_fax+"','"+item.dordre_mail+"','"+item.proprietaire_Entete+
                    "','"+item.proprietaire_nom+"','"+item.proprietaire_adresse+"','"+item.proprietaire_cp+"','"+item.proprietaire_ville+"','"+item.proprietaire_tel+"','"+item.proprietaire_fax+"','"+item.proprietaire_mail+"','"+item.bien_adresse+
                    "','"+item.bien_cp+"','"+item.bien_ville+"','"+item.bien_lieu_interne+"','"+item.bien_cadastre+"','"+item.bien_lot
                    +"','"+item.bien_lot_cave_cellier+"','"+item.bien_lot_parking_garage
                    +"','"+item.bien_lot_autre+"','" + item.bien_surface_terrain + "','" + item.bien_année_construction
                    + "','" + item.bien_parcelle + "','" + item.bien_nature + "','" + item.bien_IGH_ERP + "','" + item.bien_description
                    + "','" + item.rdv_date + "','" + item.rdv_heure + "','" + item.rdv_duree + "','" + item.rdv_contact_nom_tel
                    + "','" + item.rdv_precisions + "','" + item.rdv_clefs + "','" + item.dossier_Acces + "','" + item.dossier_Nom
                    + "','" + item.dossier_Acces_relatif + "','" + item.dossier_Archive + "','" + item.dossier_clot
                    + "','" + item.dossier_etat_rapport + "','" + item.dossier_etat_paie + "','" + item.dossier_observations
                    + "','" + item.rapport_date + "','" + item.rapport_date_envoyee + "','" + item.rapport_destinataires
                    + "','" + item.rapport_facturation + "','" + item.rapport_type + "','" + item.rapport_amiante_FCFP
                    + "','" + item.rapport_amiante_Autres + "','" + item.rapport_termites_resultat
                    + "','" + item.notaire_Entete + "','" + item.notaire_nom + "','" + item.notaire_adresse + "','" + item.notaire_cp
                    + "','" + item.notaire_ville + "','" + item.notaire_tel + "','" + item.notaire_fax + "','" + item.notaire_mail
                    + "','" + item.bien_description_cases + "','" + item.bien_perimetre + "','" + item.rapport_destinataires_mail
                    + "','" + item.dossier_etat + "','" + item.complement_visite + "','" + item.operateur_reperage + "','" + item.photo_de_presentation
                    + "','" + item.facturation_restante + "','" + item.facturation_compte_client + "','" + item.facturation_remise_globale
                    + "','" + item.facturation_date + "','" + item.facturation_date_fin + "','" + item.Donnee1 + "','" + item.Donnee2
                    + "','" + item.Donnee3 + "','" + item.Donnee4 + "','" + item.Donnee5+ "','" + item.Donnee6 + "','" + item.Donnee7
                    + "','" + item.Donnee8 + "','" + item.Donnee9 + "','" + item.Donnee10 + "','" + item.Donnee11 + "','" + item.Donnee12
                    + "','" + item.Donnee13 + "','" + item.Donnee14 + "','" + item.Donnee15 + "','" + item.Donnee16 + "','" + item.Donnee17
                    + "','" + item.Donnee18 + "','" + item.Donnee19 + "','" + item.Mission_Memo + "','" + item.operateur_certif_num
                    + "','" + item.operateur_certif_societe + "','" + item.operateur_certif_date + "','" + item.Mode_Access
                    + "','" + item.Date_commande + "','" + item.Signature_Opérateur + "','" + item.Id_facturation
                    + "','" + item.Appareil_CREP + "','" + item.Date_RDV + "','" + item.Facture_validation + "','" + item.Paiement_validation
                    + "','" + item.rapport_plus + "','" + item.Commerciaux + "','" + item.Commerciaux_autre + "','" + item.Certif_obtention
                    + "','" + item.Date_1er_paiement + "','" + item.id_dossier_liciweb + "','" + item.id_donneur_ordre + "','" + item.conclusion
                    + "','" + item.valeur_bien + "','" + item.rapport_type_expertise + "','" + item.rapport_plus_expertise
                    + "','" + item.Adresse_web_dossier + "','" + item.id_dossier_licielweb + "','" + item.type_de_dossier
                    + "','" + item.etat_licielweb + "','" + item.DATE_RDV_facture + "','" + item.Date_paiement_codee + "','" + item.AR_Amiante
                    + "','" + item.Envoie_ADEME + "','" + item.backoffice + "','" + item.date_modification + "','" + item.date_derniere_sauvegarde+ "','" + item.Id+"')";

                cmd.ExecuteNonQuery();





            }
                con.Close();



        }
    }                                                               
}                                                                   
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    