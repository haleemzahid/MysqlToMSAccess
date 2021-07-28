using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
  public  class Donnees_Dossiers
    {
        public string Num_devis_numero { get; set; }
        public string Num_dossier { get; set; }
        public string Num_dossier_lié { get; set; }
        public string dordre_type { get; set; }
        public string dordre_Entete { get; set; }
        public string dordre_nom { get; set; }
        public string dordre_adresse { get; set; }
        public string dordre_cp { get; set; }
        public string dordre_ville { get; set; }
        public string dordre_tel { get; set; }
        public string dordre_fax { get; set; }
        public string dordre_mail { get; set; }
        public string proprietaire_Entete { get; set; }
        public string proprietaire_nom { get; set; }
        public string proprietaire_adresse { get; set; }
        public string proprietaire_cp { get; set; }
        public string proprietaire_ville { get; set; }
        public string proprietaire_tel { get; set; }
        public string proprietaire_fax { get; set; }
        public string proprietaire_mail { get; set; }
        public string bien_adresse { get; set; }
        public string bien_cp { get; set; }
        public string bien_ville { get; set; }
        public string bien_lieu_interne { get; set; }
        public string bien_cadastre { get; set; }
        public string bien_lot { get; set; }
        public string bien_lot_cave_cellier { get; set; }
        public string bien_lot_parking_garage { get; set; }
        public string bien_lot_autre { get; set; }
        public string bien_surface_terrain { get; set; }
        public string bien_année_construction { get; set; }
        public string bien_parcelle { get; set; }
        public string bien_nature { get; set; }
        public string bien_IGH_ERP { get; set; }
        public string bien_description { get; set; }
        public string rdv_date { get; set; }
        public string rdv_heure { get; set; }
        public string rdv_duree { get; set; }
        public string rdv_contact_nom_tel { get; set; }
        public string rdv_precisions { get; set; }
        public string rdv_clefs { get; set; }
        public string dossier_Acces { get; set; }
        public string dossier_Nom { get; set; }
        public string dossier_Acces_relatif { get; set; }
        public string dossier_Archive { get; set; }
        public int dossier_clot { get; set; }
        public int dossier_etat_rapport { get; set; }
        public string dossier_etat_paie { get; set; }
        public string dossier_observations { get; set; }
        public string rapport_date { get; set; }
        public string rapport_date_envoyee { get; set; }
        public string rapport_destinataires { get; set; }
        public string rapport_facturation { get; set; }
        public string rapport_type { get; set; }
        public int rapport_amiante_FCFP { get; set; }
        public int rapport_amiante_Autres { get; set; }
        public int rapport_termites_resultat { get; set; }
        public string notaire_Entete { get; set; }
        public string notaire_nom { get; set; }
        public string notaire_adresse { get; set; }
        public string notaire_cp { get; set; }
        public string notaire_ville { get; set; }
        public string notaire_tel { get; set; }
        public string notaire_fax { get; set; }
        public string notaire_mail { get; set; }
        public string bien_description_cases { get; set; }
        public string bien_perimetre { get; set; }
        public string rapport_destinataires_mail { get; set; }
        public string dossier_etat { get; set; }
        public string complement_visite { get; set; }
        public string operateur_reperage { get; set; }
        public string photo_de_presentation { get; set; }
        public int facturation_restante { get; set; }
        public string facturation_compte_client { get; set; }
        public int facturation_remise_globale { get; set; }
        public string facturation_date { get; set; }
        public string facturation_date_fin { get; set; }
        public string Donnee1 { get; set; }
        public string Donnee2 { get; set; }
        public string Donnee3 { get; set; }
        public string Donnee4 { get; set; }
        public string Donnee5 { get; set; }
        public string Donnee6 { get; set; }
        public string Donnee7 { get; set; }
        public string Donnee8 { get; set; }
        public string Donnee9 { get; set; }
        public string Donnee10 { get; set; }
        public string Donnee11 { get; set; }
        public string Donnee12 { get; set; }
        public string Donnee13 { get; set; }
        public string Donnee14 { get; set; }
        public string Donnee15 { get; set; }
        public string Donnee16 { get; set; }
        public string Donnee17 { get; set; }
        public string Donnee18 { get; set; }
        public string Donnee19 { get; set; }
        public string Mission_Memo { get; set; }
        public string operateur_certif_num { get; set; }
        public string operateur_certif_societe { get; set; }
        public string operateur_certif_date { get; set; }
        public string Mode_Access { get; set; }
        public string Date_commande { get; set; }
        public string Signature_Opérateur { get; set; }
        public string Id_facturation { get; set; }
        public string Appareil_CREP { get; set; }
        public int Date_RDV { get; set; }
        public int Facture_validation { get; set; }
        public int Paiement_validation { get; set; }
        public string rapport_plus { get; set; }
        public string Commerciaux { get; set; }
        public string Commerciaux_autre { get; set; }
        public string Certif_obtention { get; set; }
        public string Date_1er_paiement { get; set; }
        public string id_dossier_liciweb { get; set; }
        public string id_donneur_ordre { get; set; }
        public string conclusion { get; set; }
        public string valeur_bien { get; set; }
        public string rapport_type_expertise { get; set; }
        public string rapport_plus_expertise { get; set; }
        public string Adresse_web_dossier { get; set; }
        public string id_dossier_licielweb { get; set; }
        public string type_de_dossier { get; set; }
        public string etat_licielweb { get; set; }
        public int DATE_RDV_facture { get; set; }
        public int Date_paiement_codee { get; set; }
        public string AR_Amiante { get; set; }
        public string Envoie_ADEME { get; set; }
        public string backoffice { get; set; }
        public string date_modification { get; set; }
        public int date_derniere_sauvegarde { get; set; }
            }
}








