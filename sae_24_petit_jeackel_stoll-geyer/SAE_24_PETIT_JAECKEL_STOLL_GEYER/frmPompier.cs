using System;
using System.Data.SQLite;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using SAE_24_Tableau_de_bord;
using SAE24_Mission_Resume;
using UCCaserne;
using mission;
using TestUCStats;
using System.Configuration;
using System.Globalization;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.IO;
using iText.Kernel.Pdf.Canvas.Draw;
using System.Collections.Generic;

namespace SAE_24_PETIT_JAECKEL_STOLL_GEYER
{
    public partial class frmPompier : Form
    {
        private SQLiteConnection cx = Connexion.Connec;
        private DataSet ds = MesDatas.DsGlobal;
        private SAE_24_Tableau_de_bord.UC_Tab_de_bord Volet1;
        private mission.USNouvMission Volet2;
        private UCCaserne.EnginUC Volet3;
        private UCGestionPompier.GestionPompier Volet4;
        private TestUCStats.UC_Stats Volet5;

        public frmPompier()
        {
            InitializeComponent();
        } 

        private void Form1_Load(object sender, EventArgs e)
        {
            BDD_Load();
            TabDeBord_Load();
            NouvMission_Load();
            GestEngin_Load();
            GestionPompier_Load();
            Stats_Load();
        }

        private void BDD_Load()
        {
            try
            {
                string req;
                DataTable schemaTable = Connexion.Connec.GetSchema("Tables");
                //string liste = "";
                for (int i = 0; i<schemaTable.Rows.Count; i++)
                {
                    string nomTable = schemaTable.Rows[i][2].ToString();
                    req = "SELECT * FROM " + nomTable;
                    SQLiteDataAdapter da = new SQLiteDataAdapter(req, cx);
                    da.Fill(MesDatas.DsGlobal, nomTable);
                    //liste += nomTable + "\n";
                }
                //MessageBox.Show(liste + "\n" + MesDatas.DsGlobal.Tables.Count.ToString());
            }
            catch (SQLiteException ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }

        private void TabDeBord_Load()
        {
            //UC Tableau de Bord
            SAE_24_Tableau_de_bord.UC_Tab_de_bord Tab = new UC_Tab_de_bord();
            Tab.Location = new System.Drawing.Point(320, 10); //245, 10
            Tab.TabDeBord_chkEnCours += chkEnCours;
            //Ajout des Missions
            ChargerMission(Tab.PanelMissions);
            Volet1 = Tab;
            Controls.Add(Tab);
        }

        private void ChargerMission(FlowLayoutPanel pnl)
        {
            for (int i = ds.Tables["Mission"].Rows.Count - 1; i>=0; i--)
            {
                DataRow dr = ds.Tables["Mission"].Rows[i];
                int id = Convert.ToInt32(dr["id"]);
                int idC = Convert.ToInt32(dr["idCaserne"]);
                int idN = Convert.ToInt32(dr["idNatureSinistre"]);
                string dateDeb = dr["dateHeureDepart"].ToString().Trim();
                string format = "yyyy-MM-dd HH:mm";

                DateTime date = DateTime.ParseExact(dateDeb, format, CultureInfo.InvariantCulture);
                string formattedDate = date.ToString("dd/MM/yyyy");

                String desc = dr["motifAppel"].ToString();
                String caserne = ds.Tables["Caserne"].Rows[idC-1]["nom"].ToString();
                String type = ds.Tables["NatureSinistre"].Rows[idN-1]["libelle"].ToString();
                int terminee = Convert.ToInt32(dr["terminee"]);
                SAE24_Mission_Resume.UC_Mission_Resume mr = new UC_Mission_Resume(id, caserne, desc, type, idN, formattedDate, terminee);
                mr.Tag = id;
                mr.finMission = finirMission;
                mr.genererPDF = genPDF;
                mr.pbtnFinirMission.Tag = id;
                mr.pbtnGenPDF.Tag = id;
                pnl.Controls.Add(mr);
            }
        }

        private void NouvMission_Load()
        {
            //Charge le volet pour les nouvelles Mission
            mission.USNouvMission nouvMission = new mission.USNouvMission(ds);
            nouvMission.Location = new System.Drawing.Point(320, 10);
            nouvMission.Visible = false;
            Volet2 = nouvMission;
            Controls.Add(nouvMission);
        }

        private void GestEngin_Load()
        {
            //Charge le volet du Gestionnaire d'engin
            UCCaserne.EnginUC gestEngin = new UCCaserne.EnginUC(ds);
            gestEngin.Location = new System.Drawing.Point(320, 10); //245, 10
            gestEngin.Visible = false;
            Volet3 = gestEngin;
            Controls.Add(gestEngin);
        }

        private void GestionPompier_Load()
        {
            UCGestionPompier.GestionPompier gestPomp = new UCGestionPompier.GestionPompier(cx);
            gestPomp.Location = new System.Drawing.Point(320, 10); //245, 10
            gestPomp.Visible = false;
            Volet4 = gestPomp;
            Controls.Add(gestPomp);
        }
        public void Stats_Load()
        {
            //Charge le volet des statistiques
            TestUCStats.UC_Stats stats = new TestUCStats.UC_Stats(cx);
            stats.Location = new System.Drawing.Point(320, 10);
            stats.Visible = false;
            Volet5 = stats;
            Controls.Add(stats);
        }

        private void Exit_App(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void MenuClick(object sender, EventArgs e)
        {

            int tag = Convert.ToInt32(((Control)sender).Tag);
            //Gère la couleur des panels de navigation
            foreach (Panel panel in pnlMenu.Controls.OfType<Panel>())
            {
                //Si le tag du sender (panel cliqué) est le même que celui du panel alors on mets la couleur spéciale
                if (Convert.ToUInt32(panel.Tag) == tag) panel.BackColor = System.Drawing.Color.LightCoral;
                //Sinon on mets la couleur classique
                else panel.BackColor = System.Drawing.Color.IndianRed;
            }
            //On rend tout les volets invisible
            Volet1.Visible = false; Volet2.Visible = false; Volet3.Visible = false; Volet4.Visible = false; Volet5.Visible = false;
            //On rend visible seulement celui avec le bon tag
            switch (tag)
            {
                case 1:
                    FetchNouvMission();
                    Volet1.Visible = true;
                    break;
                case 2:
                    Volet2.Visible = true;
                    break;
                case 3:
                    Volet3.Visible = true;
                    break;
                case 4:
                    Volet4.Visible = true;
                    break;
                case 5:
                    Volet5.Visible = true;
                    break;
            }
        }

        //Fonction delegate pour le Tableau de Bord
        private void chkEnCours(object sender, EventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            UC_Tab_de_bord Tab = Controls.OfType<UC_Tab_de_bord>().FirstOrDefault();
            FlowLayoutPanel flpM = Tab.PanelMissions;
            if (Tab != null) 
            {
                if (chk.Checked)
                {
                    foreach (UC_Mission_Resume mr in flpM.Controls.OfType<UC_Mission_Resume>())
                    {
                        if (Convert.ToInt32(ds.Tables["Mission"].Rows[Convert.ToInt32(mr.Tag) -1]["terminee"]) == 1)
                        {
                            mr.Visible = false;
                        }
                    }
                }
                else
                {
                    foreach (UC_Mission_Resume mr in flpM.Controls.OfType<UC_Mission_Resume>())
                    {
                        mr.Visible = true;
                    }
                } 
            }
        }

        private void finirMission(object sender, EventArgs e) 
        {
            //Lors de la création de l'UC bloqué le bouton si la mission est déjà ok
            PictureBox ptrb = sender as PictureBox;
            ptrb.Enabled = false;
            ptrb.Image = Properties.Resources.mission_cloturer;
            int idMission = Convert.ToInt32(ptrb.Tag);
            String heureFin = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            // MAJ Local
            ds.Tables["Mission"].Rows[idMission-1]["terminee"] = 1;
            ds.Tables["Mission"].Rows[idMission-1]["dateHeureRetour"] = heureFin;

            //On vérifie si la mission est déjà dans la base (Sauvegarder En cours)
            try
            {
                string req = "SELECT count(*) FROM Mission WHERE id = @id;";
                SQLiteCommand cmd = new SQLiteCommand(req, cx);
                cmd.Parameters.AddWithValue("id", idMission);
                int res = Convert.ToInt32(cmd.ExecuteScalar());

                if (res == 0)
                {
                    // La Mission n'existe pas dans la base donc on la crée ainsi que ses dépendances
                    AjouterMissionBDD(idMission);
                }
                else
                {
                    // La Mission est déjà dans la base SQL et on la mets juste à jour
                    // MAJ sur la vrai base de données
                    try
                    {
                        string req2 = "UPDATE Mission SET terminee = 1 WHERE id = @id;";
                        req += "UPDATE Mission SET dateHeureRetour = @dateHeureRetour WHERE id = @id;";
                        SQLiteCommand cmd2 = new SQLiteCommand(req2, cx);
                        cmd2.Parameters.AddWithValue("id", idMission);
                        cmd2.Parameters.AddWithValue("dateHeureRetour", heureFin);
                        int res2 = cmd2.ExecuteNonQuery();

                        if (res2 > 0)
                            MessageBox.Show("Mise à jour réussie !");
                        else
                            MessageBox.Show("Aucune ligne mise à jour.");
                    }
                    catch (SQLiteException ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }       
            }
            catch (SQLiteException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void genPDF(object sender, EventArgs e)
        {
            PictureBox ptrb = sender as PictureBox;
            int idMission = Convert.ToInt32(ptrb.Tag);
            DataRow dr = ds.Tables["Mission"].Rows[idMission-1];
            int idC = Convert.ToInt32(dr["idCaserne"]);
            int idN = Convert.ToInt32(dr["idNatureSinistre"]);
            Console.WriteLine("Generer PDF");
            try
            {
                PdfWriter writer = new PdfWriter("Rapport_Mission_"+idMission+".pdf");
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf);

                //Titre Principale
                Paragraph header = new Paragraph("Rapport de mission\n\n")
                   .SetBold()
                   .SetFontSize(25);
                document.Add(header);

                //Date de début et de fin
                string Deb = dr["dateHeureDepart"].ToString().Trim();
                string Fin = dr["dateHeureRetour"].ToString().Trim();
                string format = "yyyy-MM-dd HH:mm";
                DateTime dateDeb = DateTime.ParseExact(Deb, format, CultureInfo.InvariantCulture);
                DateTime dateFin = DateTime.ParseExact(Fin, format, CultureInfo.InvariantCulture);
                string dateDebFormatee = dateDeb.ToString("dd-MM-yyyy 'à' HH'h'mm");
                string dateFinFormatee = dateFin.ToString("dd-MM-yyyy 'à' HH'h'mm");

                Paragraph date = new Paragraph("Déclenchée le " + dateDebFormatee + "\nRetour le " + dateFinFormatee)
                   .SetBold()
                   .SetFontSize(15);
                document.Add(date);
            
                //Ligne de séparation
                LineSeparator line = new LineSeparator(new SolidLine());
                document.Add(line);

                //Type sinistre
                String type = ds.Tables["NatureSinistre"].Rows[idN-1]["libelle"].ToString();
                Paragraph sinistre = new Paragraph("\nType de sinistre : " + type + "\n\n")
                    .SetBold()
                    .SetFontSize(20);
                document.Add(sinistre);

                //Détails
                String motif = "Motif : " + dr["motifAppel"].ToString() + "\n";
                String adresse = "Adresse : " + dr["adresse"].ToString() + " "
                    + dr["cp"] + " " + dr["ville"] + "\n\n";
                String rendu = "Compte-rendu : " + dr["compteRendu"].ToString();

                Paragraph details = new Paragraph(motif + adresse + rendu)
                    .SetBold()
                    .SetFontSize(15);
                document.Add(details);

                // Ligne de séparation
                document.Add(line);

                //Caserne
                Paragraph c = new Paragraph("\nCaserne : " + ds.Tables["Caserne"].Rows[idC-1]["nom"].ToString() + "\n")
                    .SetBold()
                    .SetFontSize(20);
                document.Add(c);

                Paragraph p = new Paragraph("Pompiers affectés :\n")
                    .SetBold()
                    .SetFontSize(15);
                document.Add(p);

                Paragraph pl = new Paragraph(PompiersMob(idMission))
                    .SetItalic()
                    .SetFontSize(15);
                document.Add(pl);

                // Ajouter une nouvelle page vide
                pdf.AddNewPage();
                document.SetRenderer(new iText.Layout.Renderer.DocumentRenderer(document));
                document.Add(new AreaBreak(iText.Layout.Properties.AreaBreakType.NEXT_PAGE));

                Paragraph ze = new Paragraph("Engins utilisés :\n")
                   .SetBold()
                   .SetFontSize(15);
                document.Add(ze);

                Paragraph el = new Paragraph(vehiculeMob(idMission))
                    .SetItalic()
                    .SetFontSize(15);
                document.Add(el);

                document.Close();
                MessageBox.Show("PDF généré avec succès !");
            }
            catch (iText.Kernel.Exceptions.PdfException ex)
            {
                MessageBox.Show("Erreur : " + ex.ToString());
            }
        }
        private String PompiersMob(int idM)
        {
            List<String> pompiers = new List<string>();
            DataRow[] drs = ds.Tables["Mobiliser"].Select("idMission = " + idM);
            foreach (DataRow dr in drs)
            {
              
                string pompier = "";
                int matricule = Convert.ToInt32(dr["matriculePompier"]);
                int idHab = Convert.ToInt16(dr["idHabilitation"]);
                pompier += " (" + ds.Tables["Habilitation"].Rows[idHab-1]["libelle"].ToString() + ")";
                DataRow[] res1 = ds.Tables["Pompier"].Select("matricule = " + matricule);
                pompier = " " + res1[0]["prenom"] + " " + res1[0]["nom"] + pompier;
                String f = "code = '" + res1[0]["codeGrade"] + "'";
                DataRow[] res2 = ds.Tables["Grade"].Select(f);
                pompier = res2[0]["libelle"] + pompier;
                pompiers.Add(pompier);
            }
            String pomps = "";
            foreach (String pomp in pompiers)
            {
                pomps += "--> " + pomp + "\n";
            }
            return pomps;
        }

        private String vehiculeMob(int idM)
        {
            List<String> vehicules = new List<string>();
            DataRow[] drs = ds.Tables["PartirAvec"].Select("idMission = " + idM);
            foreach (DataRow dr in drs)
            {
                string vehicule = "";
                DataRow[] res1 = ds.Tables["TypeEngin"].Select("code = '" + dr["codeTypeEngin"] + "'");
                vehicule += res1[0]["nom"] + " ";
                vehicule += dr["idCaserne"] + "-" + dr["codeTypeEngin"] + "-" + dr["numeroEngin"] + " ";
                if(dr["reparationsEventuelles"].ToString() != "") vehicule += "(" + dr["reparationsEventuelles"] + ")";
                else vehicule += "(pas de réparations prévues)";
                vehicules.Add(vehicule);
            }
            String vehics = "";
            foreach (String v in vehicules)
            {
                vehics += "--> " + v + "\n";
            }
            return vehics;
        }

        private void FetchNouvMission()
        {
            if (Volet1.PanelMissions.Controls.Count != ds.Tables["Mission"].Rows.Count)
            {
                Volet1.PanelMissions.Controls.Clear();
                ChargerMission(Volet1.PanelMissions);
            }
        }

        private void AjouterMissionBDD(int idM)
        {
            //Ajout de la Mission
            try
            {
                DataRow[] result = ds.Tables["Mission"].Select("id = " + idM);
                DataRow dr = result[0];
                string query = "INSERT INTO Mission (id, dateHeureDepart, dateHeureRetour, motifAppel, adresse, cp, ville, terminee, compteRendu) " +
                   "VALUES (@id, @dateHeureDepart, @dateHeureRetour, @motifAppel, @adresse, @cp, @ville, @terminee, @compteRendu);";
                SQLiteCommand cmd = new SQLiteCommand(query, cx);
                cmd.Parameters.AddWithValue("@id", idM);
                cmd.Parameters.AddWithValue("@dateHeureDepart", dr[1]);
                cmd.Parameters.AddWithValue("@dateHeureRetour", dr[2]);
                cmd.Parameters.AddWithValue("@motifAppel", dr[3]);
                cmd.Parameters.AddWithValue("@adresse", dr[4]);
                cmd.Parameters.AddWithValue("@cp", dr[5]);
                cmd.Parameters.AddWithValue("@ville", dr[6]);
                cmd.Parameters.AddWithValue("@terminee", 1);
                cmd.Parameters.AddWithValue("@compteRendu", dr[8]);
                cmd.ExecuteNonQuery();

            }
            catch (SQLiteException ex)
            {
                Console.WriteLine(ex.Message);
            }
            //Ajout des pompiers mobiliser
            DataRow[] result2 = ds.Tables["Mobiliser"].Select("idMission = " + idM);
            foreach (DataRow dr in result2)
            {
                try
                {
                    string req = "INSERT INTO Mobiliser (matriculePompier,idMission,idHabilitation) " +
                        "VALUES (@matriculePompier,@idMission,@idHabilitation);";
                    SQLiteCommand cmd2 = new SQLiteCommand(req, cx);
                    cmd2.Parameters.AddWithValue("@matriculePompier", dr[0]);
                    cmd2.Parameters.AddWithValue("@idMission", dr[1]);
                    cmd2.Parameters.AddWithValue("@idHabilitation", dr[2]);
                    cmd2.ExecuteNonQuery();
                }
                catch (SQLiteException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            //Ajout des véhicule mobiliser
            DataRow[] result3 = ds.Tables["PartirAvec"].Select("idMission = " + idM);
            foreach (DataRow dr in result3)
            {
                try
                {
                    string req2 = "INSERT INTO Mobiliser (idCaserne,codeTypeEngin,numeroEngin,idMission,reparationsEventuelles) " +
                        "VALUES (@idCaserne,@codeTypeEngin,@numeroEngin,@idMission,@reparationsEventuelles);";
                    SQLiteCommand cmd3 = new SQLiteCommand(req2, cx);
                    cmd3.Parameters.AddWithValue("@idCaserne", dr[0]);
                    cmd3.Parameters.AddWithValue("@codeTypeEngin", dr[1]);
                    cmd3.Parameters.AddWithValue("@numeroEngin", dr[2]);
                    cmd3.Parameters.AddWithValue("@idMission", dr[3]);
                    cmd3.Parameters.AddWithValue("@reparationsEventuelles", dr[4]);
                    cmd3.ExecuteNonQuery();
                }
                catch (SQLiteException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            MessageBox.Show("Mise à jour réussie !");
        }
    }
}
