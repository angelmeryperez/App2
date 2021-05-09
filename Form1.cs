using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;// para conectarse a bases de datos
using System.Windows.Forms;

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
        int i = 0;
        int G = 0;
        int B = 0;
        int F = 0;
        int fila = 0;
        int i1 = 2;
        int i2 = 1;
        int Columna = 0;
        int V = 0;
        int[] numbi = new int[22];

        string BaseDatos;
        string BaseD1;
        string Colum;
        string[] Jugada = new string[22];
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
            INICIO();
            for (int i = 0; i <= 7; i++)
            {
                FuncionA();
                ContadorBinario(i);
                //SumaTotal1(); ya no se necesita  
                Colmna();
                Aciertos();
                ExportarDatos(dataGridView2);
            }
        }
        void FuncionA()
        {

            F = 0;
            B = 0;
            while (i1 - 1 < V + 1)
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
                while (i1 < V + 2)
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
        void Combinaciones()
        {
            for (int M = 0; M < G; M++)
            {
                try
                {
                    BaseDatos = dataGridView4.Rows[M].Cells[0].Value.ToString();
                }
                catch (Exception)
                {
                    continue;
                }
                if (BaseDatos == "") { BaseDatos = "0"; }
                if (BaseDatos != "")
                {
                    dataGridView2.Rows[M].Cells[B].Value = BaseDatos;
                }
            }

        }
        void Colmna()
        {
            B = 1;
            for (int i = 2; i <= V + 2; i++)
            {
                if (numbi[i - 2] == 1)
                {
                    dataGridView4.Sort(dataGridView4.Columns[i], ListSortDirection.Descending);
                    Combinaciones(); B++;
                }
                else
                {
                    dataGridView4.Sort(dataGridView4.Columns[i], ListSortDirection.Ascending);
                    Combinaciones(); B++;
                }

            }
        }
        void ContadorBinario(int x)
        {
            int contador = 0;
            while (contador < 20)
            {
                numbi[contador++] = x % 2;
                x /= 2;
            }
        }
        void INICIO()
        {
            dataGridView2.Rows.Add(99);
            dataGridView4.Rows.Add(100);
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
                dataGridView4.Rows[F].Cells[0].Value = F;
                F++;
            }

            Colum = dataGridView1.Rows[0].Cells[0].Value.ToString();
            Columna = int.Parse(Colum);
            V = Columna - 1000;

            if (V == 20) { G = 80; } else { G = 100; }
            for (int C = 0; C <= G; C++)
            {
                if (C <= 99)
                {
                    dataGridView2.Rows[C].Cells[0].Value = "JUGADAS: " + C;
                }
            }

            for (int i = 0; i < 20; i++)
            {
                numbi[i] = 1;
            }

            for (int K = 1; K < 27; K++)
            {
                for (int N = 0; N < G; N++)
                {
                    if (K < 23)
                    {
                        dataGridView4.Rows[N].Cells[K].Value = 0;
                    }
                    dataGridView2.Rows[N].Cells[K].Value = 0;
                }
            }
        }
        void Aciertos()
        {
            F = 0;
            for (int Filas = 0; Filas < G; Filas++)
            {
                for (int Columnas = 1; Columnas < V + 1; Columnas++)
                {
                    Jugada[Columnas] = dataGridView2.Rows[Filas].Cells[Columnas].Value.ToString();
                }

                for (int H = 0; H < G; H++)
                {
                    for (int i = 1; i < V + 1; i++)
                    {
                        for (int B = 0; B <= V - 1; B++)
                        {
                            BaseD1 = dataGridView1.Rows[H].Cells[B].Value.ToString();
                            if (Jugada[i] == BaseD1)
                            {
                                F++;
                            }
                            dataGridView2.Rows[Filas].Cells[V + 1].Value = F;
                        }
                    }
                }
                F = 0;
            }
        }
        private void ExportarDatos(DataGridView datalistado)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); // Instancia a la libreria de Microsoft Office
            excel.Application.Workbooks.Add(true); //Con esto añadimos una hoja en el Excel para exportar los archivos
            int IndiceColumna = 0;
            foreach (DataGridViewColumn columna in datalistado.Columns) //Aquí empezamos a leer las columnas del listado a exportar
            {
                IndiceColumna++;
                excel.Cells[1, IndiceColumna] = columna.Name;
            }
            int IndiceFila = 0;
            foreach (DataGridViewRow fila in datalistado.Rows) //Aquí leemos las filas de las columnas leídas
            {
                IndiceFila++;
                IndiceColumna = 0;
                foreach (DataGridViewColumn columna in datalistado.Columns)
                {
                    IndiceColumna++;
                    excel.Cells[IndiceFila + 1, IndiceColumna] = fila.Cells[columna.Name].Value;
                }
            }
            excel.Visible = true;
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}