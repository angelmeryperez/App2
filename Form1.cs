using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;// para conectarse a bases de datos

namespace App2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool Bucle = true;

        int D = 0, i = 0, Contador = 0;
        int F = 0;
        string BaseDatos;

        DataView ImportarDatos(string nombrearchivo) //COMO PARAMETROS OBTENEMOS EL NOMBRE DEL ARCHIVO A IMPORTAR
        {

            //UTILIZAMOS 12.0 DEPENDIENDO DE LA VERSION DEL EXCEL, EN CASO DE QUE LA VERSIÓN QUE TIENES ES INFERIOR AL DEL 2013, CAMBIAR A EXCEL 8.0 Y EN VEZ DE
            //ACE.OLEDB.12.0 UTILIZAR LO SIGUIENTE (Jet.Oledb.4.0)
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", nombrearchivo);
            nombrearchivo = "Regristro Actualizado";
            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            //DEPENDIENDO DEL NOMBRE QUE TIENE LA PESTAÑA EN TU ARCHIVO EXCEL COLOCAR DENTRO DE LOS []
            OleDbCommand consulta = new OleDbCommand("select * from [Hoja1$]", conector);

            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta
            };

            DataSet ds = new DataSet();

            adaptador.Fill(ds);

            conector.Close();

            return ds.Tables[0].DefaultView;


        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            
            dataGridView2.Rows.Add(99);

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //DE ESTA MANERA FILTRAMOS TODOS LOS ARCHIVOS EXCEL EN EL NAVEGADOR DE ARCHIVOS
                Filter = "Excel | *.xls;*.xlsx;",

                //AQUÍ INDICAMOS QUE NOMBRE TENDRÁ EL NAVEGADOR DE ARCHIVOS COMO TITULO
                Title = "Seleccionar Archivo"
            };

            //EN CASO DE SELECCIONAR EL ARCHIVO, ENTONCES PROCEDEMOS A ABRIR EL ARCHIVO CORRESPONDIENTE
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dataGridView1.DataSource = ImportarDatos(openFileDialog.FileName);
            }

            while (F < 100) // Este while imprime es la que se encargar de los numeros en la colunna 0
            {
                dataGridView2.Rows[F].Cells[0].Value = F;
                dataGridView2.Rows[F].Cells[1].Value = Frecuencia(F);
                Contador = 0;
                F++;
            }
            
        }

        int Frecuencia(int numero)
        {

            while (Bucle)
            {
                BaseDatos = dataGridView1.Rows[Contador].Cells[1].Value.ToString();
                i = int.Parse(BaseDatos);
                if (i == numero) { Contador++; }
                if (BaseDatos == "Final") { Bucle = false; }

            }
            Bucle = true;
            return Contador;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ////string conexion = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = C:/Users/Angel Mery Perez/Desktop/Loto/Programs con exel/Regristro Actualizado.xlsx ;Extended Properties = \"Exel 8.0;HDR = Yes\"";
            //String conexion = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = C:/Users/Angel Mery Perez/Desktop/Loto/Programs con exel/Regristro Actualizado.xlsx; Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            //OleDbConnection conector = default(OleDbConnection);
            //conector = new OleDbConnection(conexion);
            //conector.Open();

            //OleDbCommand consulta = default(OleDbCommand);
            //consulta = new OleDbCommand("selet * from[Hoja1$]",conector);

            //OleDbDataAdapter Adactador = new OleDbDataAdapter();
            //Adactador.SelectCommand = consulta;

            //DataSet DS = new DataSet();

            //Adactador.Fill(DS);
            //dataGridView1.DataSource = DS.Tables[0];
            //conector.Close();


        }
    }
}
