using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace NotepadToExcel
{
    public partial class Form1 : Form
    {
        TextReader notepad = new StreamReader("C:\\Users\\Randi\\Desktop\\Projects\\App Desktops\\WindowsForms\\NotepadToExcel\\source\\Prueba.txt");
        public Form1()
        {
            InitializeComponent();
        }

        string texto = "";
        string nuevotexto = "";
        string txt = "";
        int x = 0;
        bool bucle = true;

        int countnumber = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            
            texto = notepad.ReadToEnd();

            while (bucle)
            {

                //Compara
                for (char i = '0'; i <= '9'; i++)
                {

                    if (texto[x] == i)
                    {
                        if (texto[x - 1] == ' ' || texto[x + 1] == ' ') { }
                        else { nuevotexto += texto[x]; }
                    }

                }

                x++;
                if (texto[x] == 'x') { bucle = false; nuevotexto += "x"; }
            }

            bucle = true;
            x = 0;

            while (bucle)
            {
                if (nuevotexto[x] == 'x') { break; }
                txt += "" + nuevotexto[x];

                x++;
                if ((x) % 2 == 0) { txt += " "; }
                if (x % (countnumber * 2) == 0) { txt += "\n\n"; }
            }
            
            richTextBox1.Text = nuevotexto;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            countnumber = int.Parse(textBox1.Text);
        }
    }
}
