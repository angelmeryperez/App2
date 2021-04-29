﻿using System;
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

            dataGridView2.Rows.Add(100);
            dataGridView3.Rows.Add(100);

            dataGridView2.Rows[0].Cells[2].Value = "PRIMERA";
            dataGridView2.Rows[0].Cells[3].Value = "SEGUNDA";
            dataGridView2.Rows[0].Cells[4].Value = "TERCERA";
            dataGridView2.Rows[0].Cells[5].Value = "PRIMERA";
            dataGridView2.Rows[0].Cells[6].Value = "SEGUNDA";
            dataGridView2.Rows[0].Cells[7].Value = "TERCERA";
            dataGridView2.Rows[0].Cells[8].Value = "PRIMERA";
            dataGridView2.Rows[0].Cells[9].Value = "SEGUNDA";
            dataGridView2.Rows[0].Cells[10].Value = "TERCERA";
            dataGridView2.Rows[0].Cells[11].Value = "PRIMERA";
            dataGridView2.Rows[0].Cells[12].Value = "SEGUNDA";
            dataGridView2.Rows[0].Cells[13].Value = "TERCERA";

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
                dataGridView2.Rows[F + 1].Cells[0].Value = F;
                dataGridView3.Rows[F].Cells[1].Value = F;
                F++;
            }

            BaseD1 = dataGridView1.Rows[0].Cells[0].Value.ToString();
            Suma = int.Parse(BaseD1);

            switch (Suma)
            {
                case 1002:
                    FuncionA();
                    SumaTotal1();
                    break;

                case 1003:
                    F = 0;
                    B = 0;

                    while (F < 100)
                    {
                        dataGridView3.Rows[F].Cells[0].Value = Frecuencia(F);
                        F++;
                    }
                    F = 0;
                    break;




                default:
                    break;
            }
            Suma = 0;           

            dataGridView3.Sort(F22, ListSortDirection.Descending);
        }
        void FuncionA() 
        {
            F = 0;
            while (i1 < 14)
            {
                while (i2 <= 100)
                {
                    dataGridView2.Rows[i2].Cells[i1].Value = Frecuencia(F++);
                    i2++;
                }
                F = 0;
                i2 = 1;
                i1++;
                B++;
            }

        }

        void SumaTotal1() 
        {
            i1 = 2;
            i2 = 1;
            int SumaTotal = 0;
            while (i2 <= 100)
            {
                while (i1 < 14)
                {
                    BaseD1 = dataGridView2.Rows[i2].Cells[i1].Value.ToString();
                    Suma = int.Parse(BaseD1);
                    SumaTotal += Suma;

                    i1++;
                }
                F = 0;
                i1 = 2;
                dataGridView2.Rows[i2].Cells[1].Value = SumaTotal;
                dataGridView3.Rows[i2].Cells[0].Value = SumaTotal;
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
                BaseDatos = dataGridView1.Rows[fila].Cells[B].Value.ToString();
                if (BaseDatos == "") { return contador; }
                i = int.Parse(BaseDatos);
                if (i == numero) { contador++; }
                if (BaseDatos == "1001") { Bucle = false; }
                fila++;
            }
            Bucle = true;
            return contador;
        }
    }
}
