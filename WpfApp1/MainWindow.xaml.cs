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
        private  string connectionString;

        private  MySqlConnection connection;
        string myConnectionString = @"Driver={Microsoft Access Driver (*.mdb)};" + "Dbq=DB.mdb;";

        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnSyncedClicked(object sender, RoutedEventArgs e)
        {
            //       InsertIntoMSAccess();
            //OpenMySqlConnection();
            var b = GetMSAccessList();
        var a =   GetMySqlList();
            dg.ItemsSource = a;
        }

        public OdbcConnection OpenMSAccessConnection()
        {
            
            var myConnection = new OdbcConnection();
                myConnection.ConnectionString = myConnectionString;
                
            return myConnection;
            

            }

        public List<Donnees_Dossiers> GetMySqlList()
        {
            List<Donnees_Dossiers> l = new List<Donnees_Dossiers>();
            var con = OpenMySqlConnection();
            string str = "SELECT * FROM Donnees_Dossiers";

            MySqlDataReader mySqlDataReader = (new MySqlCommand(str, con)).ExecuteReader();
            while (mySqlDataReader.Read())
            {
                l.Add(new Donnees_Dossiers()
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
                    facturation_date_fin= mySqlDataReader.GetString(77),
                    Donnee1= mySqlDataReader.GetString(78),
                    Donnee2= mySqlDataReader.GetString(79),
                    Donnee3= mySqlDataReader.GetString(80),
                    Donnee4= mySqlDataReader.GetString(81),
                    Donnee5= mySqlDataReader.GetString(82),
                    Donnee6= mySqlDataReader.GetString(83),
                    Donnee7= mySqlDataReader.GetString(84),
                    Donnee8= mySqlDataReader.GetString(85),
                    Donnee9= mySqlDataReader.GetString(86),
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
                    date_derniere_sauvegarde = mySqlDataReader.GetInt32(130)
                }
                                                                                                                                                                                                                                        
                   ) ;                                                 
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
                l.Add(new Donnees_Dossiers()
                {


                    Id =Convert.ToInt32( reader.GetInt32(0)),
                    Num_devis_numero =Convert.ToString(reader.GetString(1)),
                    Num_dossier = Convert.ToString(reader.GetString(2)),
                    Num_dossier_lié = Convert.ToString(reader.GetString(3)),
                    dordre_type = Convert.ToString(reader.GetString(4)),
                    //dordre_Entete = Convert.ToString(reader.GetString(5)),
                    //dordre_nom = Convert.ToString(reader.GetString(6)),
                    //dordre_adresse = Convert.ToString(reader.GetString(7)),
                    //dordre_cp = Convert.ToString(reader.GetString(8)),
                    //dordre_ville = Convert.ToString(reader.GetString(9)),
                    //dordre_tel = Convert.ToString(reader.GetString(10)),
                    //dordre_fax = Convert.ToString(reader.GetString(11)),
                    //dordre_mail = reader.GetString(12),
                    //proprietaire_Entete = reader.GetString(13),
                    //proprietaire_nom = reader.GetString(14),
                    //proprietaire_adresse = reader.GetString(15),
                    //proprietaire_cp = reader.GetString(16),
                    //proprietaire_ville = reader.GetString(17),
                    //proprietaire_tel = reader.GetString(18),
                    //proprietaire_fax = reader.GetString(19),
                    //proprietaire_mail = reader.GetString(20),
                    //bien_adresse = reader.GetString(21),
                    //bien_cp = reader.GetString(22),
                    //bien_ville = reader.GetString(23),
                    //bien_lieu_interne = reader.GetString(24),
                    //bien_cadastre = reader.GetString(25),
                    //bien_lot = reader.GetString(26),
                    //bien_lot_cave_cellier = reader.GetString(27),
                    //bien_lot_parking_garage = reader.GetString(28),
                    //bien_lot_autre = reader.GetString(29),
                    //bien_surface_terrain = reader.GetString(30),
                    //bien_année_construction = reader.GetString(31),
                    //bien_parcelle = reader.GetString(32),
                    //bien_nature = reader.GetString(33),
                    //bien_IGH_ERP = reader.GetString(34),
                    //bien_description = reader.GetString(35),
                    //rdv_date = reader.GetString(36),
                    //rdv_heure = reader.GetString(37),
                    //rdv_duree = reader.GetString(38),
                    //rdv_contact_nom_tel = reader.GetString(39),
                    //rdv_precisions = reader.GetString(40),
                    //rdv_clefs = reader.GetString(41),
                    //dossier_Acces = reader.GetString(42),
                    //dossier_Nom = reader.GetString(43),
                    //dossier_Acces_relatif = reader.GetString(44),
                    //dossier_Archive = reader.GetString(45),
                    //dossier_clot = reader.GetInt32(46),
                    //dossier_etat_rapport = reader.GetInt32(47),
                    //dossier_etat_paie = reader.GetString(48),
                    //dossier_observations = reader.GetString(49),
                    //rapport_date = reader.GetString(50),
                    //rapport_date_envoyee = reader.GetString(51),
                    //rapport_destinataires = reader.GetString(52),
                    //rapport_facturation = reader.GetString(53),
                    //rapport_type = reader.GetString(54),
                    //rapport_amiante_FCFP = reader.GetInt32(55),
                    //rapport_amiante_Autres = reader.GetInt32(56),
                    //rapport_termites_resultat = reader.GetInt32(57),
                    //notaire_Entete = reader.GetString(58),
                    //notaire_nom = reader.GetString(59),
                    //notaire_adresse = reader.GetString(60),
                    //notaire_cp = reader.GetString(61),
                    //notaire_ville = reader.GetString(62),
                    //notaire_tel = reader.GetString(63),
                    //notaire_fax = reader.GetString(64),
                    //notaire_mail = reader.GetString(65),
                    //bien_description_cases = reader.GetString(66),
                    //bien_perimetre = reader.GetString(67),
                    //rapport_destinataires_mail = reader.GetString(68),

                    //dossier_etat = reader.GetString(69),
                    //complement_visite = reader.GetString(70),
                    //operateur_reperage = reader.GetString(71),
                    //photo_de_presentation = reader.GetString(72),
                    //facturation_restante = reader.GetInt32(73),
                    //facturation_compte_client = reader.GetString(74),
                    //facturation_remise_globale = reader.GetInt32(75),
                    //facturation_date = reader.GetString(76),
                    //facturation_date_fin = reader.GetString(77),
                    //Donnee1 = reader.GetString(78),
                    //Donnee2 = reader.GetString(79),
                    //Donnee3 = reader.GetString(80),
                    //Donnee4 = reader.GetString(81),
                    //Donnee5 = reader.GetString(82),
                    //Donnee6 = reader.GetString(83),
                    //Donnee7 = reader.GetString(84),
                    //Donnee8 = reader.GetString(85),
                    //Donnee9 = reader.GetString(86),
                    //Donnee10 = reader.GetString(87),
                    //Donnee11 = reader.GetString(88),
                    //Donnee12 = reader.GetString(89),
                    //Donnee13 = reader.GetString(90),
                    //Donnee14 = reader.GetString(91),
                    //Donnee15 = reader.GetString(92),
                    //Donnee16 = reader.GetString(93),
                    //Donnee17 = reader.GetString(94),
                    //Donnee18 = reader.GetString(95),
                    //Donnee19 = reader.GetString(96),
                    //Mission_Memo = reader.GetString(97),
                    //operateur_certif_num = reader.GetString(98),
                    //operateur_certif_societe = reader.GetString(99),
                    //operateur_certif_date = reader.GetString(100),
                    //Mode_Access = reader.GetString(101),
                    //Date_commande = reader.GetString(102),
                    //Signature_Opérateur = reader.GetString(103),
                    //Id_facturation = reader.GetString(104),
                    //Appareil_CREP = reader.GetString(105),
                    //Date_RDV = reader.GetInt32(106),
                    //Facture_validation = reader.GetInt32(107),
                    //Paiement_validation = reader.GetInt32(108),
                    //rapport_plus = reader.GetString(109),
                    //Commerciaux = reader.GetString(110),
                    //Commerciaux_autre = reader.GetString(111),
                    //Certif_obtention = reader.GetString(112),
                    //Date_1er_paiement = reader.GetString(113),
                    //id_dossier_liciweb = reader.GetString(114),
                    //id_donneur_ordre = reader.GetString(115),
                    //conclusion = reader.GetString(116),
                    //valeur_bien = reader.GetString(117),
                    //rapport_type_expertise = reader.GetString(118),
                    //rapport_plus_expertise = reader.GetString(119),
                    //Adresse_web_dossier = reader.GetString(120),
                    //id_dossier_licielweb = reader.GetString(121),
                    //type_de_dossier = reader.GetString(122),
                    //etat_licielweb = reader.GetString(123),
                    //DATE_RDV_facture = reader.GetInt32(124),
                    //Date_paiement_codee = reader.GetInt32(125),
                    //AR_Amiante = reader.GetString(126),
                    //Envoie_ADEME = reader.GetString(127),
                    //backoffice = reader.GetString(128),
                    //date_modification = reader.GetString(129),
                    //date_derniere_sauvegarde = reader.GetInt32(130)
                }

                   );
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
            cmd.CommandText = "INSERT INTO Donnees_Dossiers(dordre_aDonnee9 { get; set; }dresse,dordre_cp,dordre_Entete,dordre_nom,dordre_tel)VALUES('" + a.dordre_adresse + "','" + a.dordre_cp + "','" + a.dordre_Entete + "','"+a.dordre_nom+"','"+a.dordre_tel+"')";
                                                                    
                                                                    
                cmd.ExecuteNonQuery();                              
            con.Close();                                            
                                                                    
                                                                    
        }                                                           
    }                                                               
}                                                                   
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    
                                                                    