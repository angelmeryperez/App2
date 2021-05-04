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

        int Suma = 0;
        int i    = 0;
        int B    = 0;
        int F    = 0;
        string BaseDatos;
        string BaseD1;
        int fila = 0;
        int i1   = 2;
        int i2   = 1;
        string Colum;
        int Columna = 0;
        int V = 0;
        //string[] COLUM = new string[] { Column1, "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", };

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

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            dataGridView2.Rows.Add(80);
            dataGridView4.Rows.Add(99);

            for (int C = 0; C <= 80; C++)
            {
                dataGridView2.Rows[C].Cells[0].Value = "JUGADAS: " + C ; 
            }

            for (int K = 1; K < 21; K++)
            {
                for (int N = 0; N < 80; N++)
                {
                    dataGridView2.Rows[N].Cells[K].Value = 0;
                } 
            }

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
            F = 0;
            while (F < 100) // Este while imprime es la que se encargar de los numeros en la colunna 0
            {
                dataGridView4.Rows[F + 1].Cells[0].Value = F;
                F++;
            }

            BaseD1 = dataGridView1.Rows[0].Cells[0].Value.ToString();
            Suma = int.Parse(BaseD1);
            
            Colum = dataGridView1.Rows[0].Cells[0].Value.ToString();
            Columna = int.Parse(Colum);
            V = Columna - 1000;

            FuncionA();
            SumaTotal1();
            Colmna();

        }
        void FuncionA() 
        {
           
            F = 0;

            while (i1-1 < V+1)
            {
                while (i2 <= 100)
                {
                    dataGridView4.Rows[i2].Cells[i1].Value = Frecuencia(F++);
                    i2++;     
                }
                B++;
                F = 0;
                i2 = 1;
                i1++;
            }
            dataGridView4.Rows[1].Cells[1].Value = B; 
        }

        void SumaTotal1() 
        {
            i1 = 2;
            i2 = 1;
            int SumaTotal = 0;
            while (i2 <= 100)
            {
                while (i1 < V+2)
                {
                    try
                    {
                        BaseD1 = dataGridView4.Rows[i2].Cells[i1].Value.ToString();
                    }
                    catch (Exception)
                    {

                        continue;
                    }
                    Suma = int.Parse(BaseD1);
                    SumaTotal += Suma;
                    i1++;
                }
                F = 0;
                i1 = 2;
                dataGridView4.Rows[i2].Cells[1].Value = SumaTotal;
                SumaTotal = 0;
                i2++;
            }
            
        }

        int Frecuencia(int numero)
        {
            fila = 0;
            int contador = 0;
            while (Bucle)
            {
                try
                {
                    BaseDatos = dataGridView1.Rows[fila].Cells[B].Value.ToString();
                }
                catch (Exception)
                {

                    MessageBox.Show("Proceso Terminado");
                }
                if (BaseDatos == "") { return contador; }
                i = int.Parse(BaseDatos);
                if (i == numero) { contador++; }
                if (BaseDatos == "1000") { Bucle = false; }
                fila++;
            }
            Bucle = true;
            return contador;
        }

        void Combinaciones(int W) 
        {
                for (int M = 0; M < 81; M++)
                {
                    try
                    {
                        BaseDatos = dataGridView4.Rows[M].Cells[0].Value.ToString();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    if (BaseDatos == "") { continue; }
                    if (BaseDatos != "")
                    {
                        dataGridView2.Rows[M].Cells[B].Value = BaseDatos;
                    }
                } 
            
        }

        void Colmna()
        {
            B = 1;
            if (B <= V)
            {
                dataGridView4.Sort(Column2, ListSortDirection.Descending);
                Combinaciones(0); 
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column3, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column4, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column5, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column6, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column7, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column8, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column9, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column10, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column11, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column12, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column13, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column14, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column15, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column16, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column17, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column18, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column19, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column20, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column42, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
            if (B <= V)
            {
                dataGridView4.Sort(Column43, ListSortDirection.Descending);
                Combinaciones(0);
            }
            B++;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
    }
}
