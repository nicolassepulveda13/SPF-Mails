using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
namespace pruebamails
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (abrir.ShowDialog() == DialogResult.OK)
            {

                string leyenda = textBox2.Text;
                if (checkBox1.Checked == true)
                {
                    #region retirados
                    OleDbConnection Conexion = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elopez\Documents\Mail.mdb;");
                    Conexion.Open();
                    string Query = "";
                    Query = "SELECT * FROM MailsReti ";
                    OleDbDataAdapter DA = new OleDbDataAdapter(Query, Conexion);
                    DataTable DT = new DataTable();
                    DA.Fill(DT);
                    Conexion.Close();
                    Query = "SELECT * FROM CodigosR ORDER BY ID";       
                    OleDbDataAdapter DA1 = new OleDbDataAdapter(Query, Conexion);
                    DataTable DT1 = new DataTable();
                    DA1.Fill(DT1);
                    Conexion.Close();
                    
                    int cont = 0;

                    for (int i = 0; DT.Rows.Count > i; i++)
                    {
                        int contadoradj = 0;
                        string maildestin = Convert.ToString(DT.Rows[i]["Entidad"]);
                        //string maildestin1 = Convert.ToString(DT.Rows[i]["Entidad"]);
                        //string maildestin = "nico_xeneise93@hotmail.com";
                        int IDe = Convert.ToInt32(DT.Rows[i]["Id"]);

                        MailMessage email = new MailMessage();
                        email.To.Add(new MailAddress(maildestin));
                        email.From = new MailAddress("spfretirados@gmail.com");
                        email.Subject = "Cargo 0622 ";
                        email.Body = leyenda + ".\n \n \n \n \n" +
                        "Liquidaciones de Haberes. "  ;
                        email.IsBodyHtml = false;
                        for (int j = 0; DT1.Rows.Count > j; j++)
                        {
                            int ide2 = Convert.ToInt32(DT1.Rows[j]["Id"]);

                            if (IDe == ide2)
                            {

                                string codigos = Convert.ToString(DT1.Rows[j]["Codigos"]);
                                string codigo7 = "7" + codigos.Substring(1, 2);
                                string directorio = (abrir.SelectedPath + @"\RC00" + codigos + ".txt");
                                string directorio7 = (abrir.SelectedPath + @"\RC00" + codigo7+ ".txt");

                                string error = (abrir.SelectedPath + @"\REBOTADOS00" + codigos + ".txt");
                                string error7 = (abrir.SelectedPath + @"\REBOTADOS00" + codigo7 + ".txt");
                                if (File.Exists(directorio))
                                {
                                    if (codigos == "602")
                                    {

                                        //string altas = (abrir.SelectedPath + @"\RTOT0622.txt");
                                        ////string baja = (abrir.SelectedPath + @"\BAS00R.DBF");
                                        ////string mod = (abrir.SelectedPath + @"\BAS01R.DBF");
                                        ////string tot = (abrir.SelectedPath + @"\BAS26R.DBF");
                                        //email.Attachments.Add(new Attachment(altas));
                                        ////email.Attachments.Add(new Attachment(baja));
                                        ////email.Attachments.Add(new Attachment(mod));
                                        ////email.Attachments.Add(new Attachment(tot));
                                    
                                    }
                                    email.Attachments.Add(new Attachment(directorio));
                                    if (File.Exists(directorio7))
                                    {
                                        email.Attachments.Add(new Attachment(directorio7));
                                    }
                                    contadoradj += 1;
                                 
                                 
                                }
                                if (File.Exists(error))
                                {
                                    email.Attachments.Add(new Attachment(error));
                                }
                                if (File.Exists(error7))
                                {
                                    email.Attachments.Add(new Attachment(error7));
                                }

                            }
                        }
                        email.Priority = MailPriority.Normal;
                        string output = null;
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = "smtp.gmail.com";
                        smtp.Port = 587;
                        smtp.EnableSsl = true;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential("spfretirados@gmail.com", "dgareti1313");



                        try
                        {
                            if (contadoradj != 0)
                            {
                                cont += 1;
                                label1.Text = Convert.ToString(cont);
                                Application.DoEvents();
                                smtp.Send(email);
                                email.Dispose();
                                //output = "Corre electrónico fue enviado satisfactoriamente.";
                            }
                        }
                     
                        catch (Exception ex)
                           {
                               output = "Error enviando correo electrónico: " + ex.Message;

                        }

                        //MessageBox.Show(output);


                    }

                    MessageBox.Show("termino");
                }
                    #endregion
                else if (checkBox2.Checked == true)
                {
                    #region ACTIVAD
                    OleDbConnection Conexion = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elopez\Documents\Mail.mdb;");
                    Conexion.Open();
                    string Query = "";
                    Query = "SELECT * FROM Mailacti ";
                    OleDbDataAdapter DA = new OleDbDataAdapter(Query, Conexion);
                    DataTable DT = new DataTable();
                    DA.Fill(DT);
                    Conexion.Close();
                    Query = "SELECT * FROM CodigosA ORDER BY ID";
                    OleDbDataAdapter DA1 = new OleDbDataAdapter(Query, Conexion);
                    DataTable DT1 = new DataTable();
                    DA1.Fill(DT1);
                    Conexion.Close();

                    int cont = 0;
                    for (int i = 0; DT.Rows.Count > i; i++)
                    {
                        int contadoradj = 0;
                        string maildestin = Convert.ToString(DT.Rows[i]["MAIL"]);
                        //string maildestin = "nico_xeneise93@hotmail.com";
                        int IDe = Convert.ToInt32(DT.Rows[i]["Id"]);

                        MailMessage email = new MailMessage();
             
                        email.To.Add(new MailAddress(maildestin));
                        email.From = new MailAddress("spf_actividad@hotmail.com");
                        email.Subject = "Cargo 0622 ";
                        email.Body = leyenda + ".\n \n \n \n \n" +
                        "Liquidaciones de Haberes. ";
                        email.IsBodyHtml = false;
                        for (int j = 0; DT1.Rows.Count > j; j++)
                        {
                            int ide2 = Convert.ToInt32(DT1.Rows[j]["Id"]);

                            if (IDe == ide2)
                            {

                                string codigos = Convert.ToString(DT1.Rows[j]["Codigo"]);
                                string codigo7 = "7" + codigos.Substring(1, 2);
                                string directorio = (abrir.SelectedPath + @"\00" + codigos + ".txt");
                                string directorio7 = (abrir.SelectedPath + @"\00" + codigo7 + ".txt");
                                string encope = (abrir.SelectedPath + @"\E00" + codigos + ".txt");
                                string encope7 = (abrir.SelectedPath + @"\E00" + codigo7 + ".txt");
                                string directorio1 = (abrir.SelectedPath + @"\COD00" + codigos + ".txt");
                                string directorio17 = (abrir.SelectedPath + @"\COD00" + codigo7 + ".txt");
                                string encope1 = (abrir.SelectedPath + @"\EC00" + codigos + ".txt");
                                string encope17 = (abrir.SelectedPath + @"\EC00" + codigo7 + ".txt");
                                string fal = (abrir.SelectedPath + @"\FAL00" + codigos + ".txt");
                                string fal7 = (abrir.SelectedPath + @"\FAL00" + codigo7 + ".txt");

                                if (File.Exists(directorio))
                                {
                                    if (codigos == "602")
                                    {
                                    //    string tot = (abrir.SelectedPath + @"\ATOT0622");
                                    //    string etot = (abrir.SelectedPath + @"\ETOT0622");
                                    ////    //string baja = (abrir.SelectedPath + @"\BAS00.DBF");
                                    ////    //string mod = (abrir.SelectedPath + @"\Base01.dbf");
                                    ////    //string alta = (abrir.SelectedPath + @"\BASE26.DBF");
                                    //    email.Attachments.Add(new Attachment(tot));
                                    //    email.Attachments.Add(new Attachment(etot));
                                    ////    //email.Attachments.Add(new Attachment(mod));
                                    ////    //email.Attachments.Add(new Attachment(baja));
                                    ////    //email.Attachments.Add(new Attachment(alta));
                                    }
                                    email.Attachments.Add(new Attachment(directorio));
                                    contadoradj += 1;
                                    if (File.Exists(encope))
                                    {
                                        email.Attachments.Add(new Attachment(encope));
                                    }
                                    if (File.Exists(fal))
                                    {
                                        email.Attachments.Add(new Attachment(fal));
                                    }
                                    if (File.Exists(fal7))
                                    {
                                        email.Attachments.Add(new Attachment(fal7));
                                    }
                                    if (File.Exists(directorio7))
                                    {
                                        email.Attachments.Add(new Attachment(directorio7));
                                    }
                                    if (File.Exists(encope7))
                                    {
                                        email.Attachments.Add(new Attachment(encope7));
                                    }

                                }
                                if (File.Exists(directorio1))
                                {
                                    if (codigos == "602")
                                    {
                                    //    string tot = (abrir.SelectedPath + @"\ATOT0622.txt");
                                    //    string etot = (abrir.SelectedPath + @"\ETOT0622.txt");
                                    ////    //string baja = (abrir.SelectedPath + @"\BASE00.DBF");
                                    ////    //string mod = (abrir.SelectedPath + @"\Base01.dbf");
                                    ////    //string alta = (abrir.SelectedPath + @"\BASE26.DBF");
                                    //    email.Attachments.Add(new Attachment(tot));
                                    //    email.Attachments.Add(new Attachment(etot));
                                    ////    //email.Attachments.Add(new Attachment(mod));
                                    ////    //email.Attachments.Add(new Attachment(baja));
                                    ////    //email.Attachments.Add(new Attachment(alta));
                                    }
                                    email.Attachments.Add(new Attachment(directorio1));
                                    contadoradj += 1;
                                    if (File.Exists(encope1))
                                    {
                                        email.Attachments.Add(new Attachment(encope1));
                                    }
                                    if (File.Exists(fal))
                                    {
                                        email.Attachments.Add(new Attachment(fal));
                                    }
                                    if (File.Exists(directorio17))
                                    {
                                        email.Attachments.Add(new Attachment(directorio17));
                                    }
                                    if (File.Exists(encope17))
                                    {
                                        email.Attachments.Add(new Attachment(encope17));
                                    }
                                    if (File.Exists(fal7))
                                    {
                                        email.Attachments.Add(new Attachment(fal7));
                                    }
                                }
                            }
                        }
                        email.Priority = MailPriority.Normal;
                        string output = null;
                        SmtpClient smtp = new SmtpClient();
                        smtp.Host = "smtp.outlook.com";
                        smtp.Port = 587;
                        smtp.EnableSsl = true;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential("spf_actividad@hotmail.com", "dgaacti1313");
                        try
                        {
                            if (contadoradj != 0)
                            {
                                cont += 1;
                                label1.Text = Convert.ToString(cont);
                                Application.DoEvents();
                                smtp.Send(email);
                                email.Dispose();
                                //output = "Corre electrónico fue enviado satisfactoriamente.";
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show( "Error enviando correo electrónico: " + ex.Message);
                        }

                        //MessageBox.Show(output);


                    }

                    MessageBox.Show("termino");
                }
                    #endregion
            }
        }
        private void button2_Click(object sender, System.EventArgs e)
        {
       
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, System.EventArgs e)
        {

        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            #region ACTIVAD
            if (abrir.ShowDialog() == DialogResult.OK)
            {
                string leyenda = textBox2.Text;
                OleDbConnection Conexion = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\elopez\Documents\Mail.mdb;");

                Conexion.Open();
                string Query = "";
                Query = "SELECT * FROM MailEnti";
                OleDbDataAdapter DA = new OleDbDataAdapter(Query, Conexion);
                DataTable DT = new DataTable();
                DA.Fill(DT);
                Conexion.Close();
                Query = "SELECT * FROM CodigoEnt";
                OleDbDataAdapter DA1 = new OleDbDataAdapter(Query, Conexion);
                DataTable DT1 = new DataTable();
                DA1.Fill(DT1);
                Conexion.Close();

                int cont = 0;
                for (int i = 0; DT.Rows.Count > i; i++)
                {
                    int contadoradj = 0;
                    //string lala = Convert.ToString(DT.Rows[i]["MAIL"]);
                    string maildestin = Convert.ToString(DT.Rows[i]["MAIL"]);
                    //string maildestin = "nico_xeneise93@hotmail.com";
                    int IDe = Convert.ToInt32(DT.Rows[i]["Id"]);
                    //string maildestin = "economatocp1@yahoo.com.ar";
                    MailMessage email = new MailMessage();
                    email.To.Add(new MailAddress(maildestin));
                    email.From = new MailAddress("spf_actividad@hotmail.com");
                    email.Subject = "Cargo 0622 ";
                    email.Body = leyenda + ".\n \n \n \n \n" +
                "Liquidaciones de Haberes. ";
                    email.IsBodyHtml = false;
                    for (int j = 0; DT1.Rows.Count > j; j++)
                    {
                        int ide2 = Convert.ToInt32(DT1.Rows[j]["Id"]);

                        if (IDe == ide2)
                        {
                            string codigos = Convert.ToString(DT1.Rows[j]["Archivo"]);

                            string directorio = (abrir.SelectedPath + @"\Enti\" + codigos);

                            if (File.Exists(directorio))
                            {
                                email.Attachments.Add(new Attachment(directorio));
                                contadoradj += 1;
                            }

                        }
                    }
                    email.Priority = MailPriority.Normal;
                    string output = null;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.live.com";
                    smtp.Port = 587;                   
                    smtp.EnableSsl = true;
                  //  smtp.Timeout = 6000;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential("spf_actividad@hotmail.com", "dgaacti1313");



                    try
                    {
                        if (contadoradj != 0)
                        {
                            cont += 1;
                            label1.Text = Convert.ToString(cont);
                            Application.DoEvents();
                            smtp.Send(email);
                            email.Dispose();
                            //output = "Corre electrónico fue enviado satisfactoriamente.";
                        }
                    }
                    catch (Exception ex)
                    {
                        output = "Error enviando correo electrónico: " + ex.Message;
                    }

                    //MessageBox.Show(output);


                }

                MessageBox.Show("termino");

            #endregion
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}