
using Ejercicio03Marzo.Algoritmos.MetodosAlgoritmos;
using Ejercicio03Marzo.Algoritmos.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ejercicio03Marzo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text.Equals(""))
                {
                    MessageBox.Show("Uno de los valores: semilla ó total de datos es vacío, favor de corregir");
                    return;
                }
                int semilla = Int32.Parse(textBox1.Text);
                int totalDatos = Int32.Parse(textBox3.Text);
                MetodoInicial algoritmo = new MetodoInicial();
                if (semilla > 0)
                {
                    List<int> listaValoresAleatorios = algoritmo.AlgoritmoGeneradorNumerosAleatorios(semilla, totalDatos);
                    textBox2.Text = algoritmo.AlgoritmoPrincipal().ToString();
                }
                else
                {
                    MessageBox.Show("Una de las condiciones es incorrecta, favor de revisar:semilla > 0");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido una excepción");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void llenarGrid(int numeroDatos, List<Demanda> listaInicial)
        {
            string numeroColumna1 = "1";
            string numeroColumna2 = "2";
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add(numeroColumna1, "ID");
            dataGridView1.Columns.Add(numeroColumna2, "Algoritmo");
            for (int i = 0; i < numeroDatos; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[Int32.Parse(numeroColumna1) - 1].Value = (i + 1).ToString();
                dataGridView1.Rows[i].Cells[Int32.Parse(numeroColumna2) - 1].Value = (listaInicial[i].CantidadRequerida).ToString();
            }
        }

        private void DescargaExcel(DataGridView data)
        {
            Microsoft.Office.Interop.Excel.Application exportarExcel = new Microsoft.Office.Interop.Excel.Application();
            exportarExcel.Application.Workbooks.Add(true);
            int indiceColumna = 0;
            foreach (DataGridViewColumn columna in data.Columns)
            {
                indiceColumna = indiceColumna + 1;
                exportarExcel.Cells[1, indiceColumna] = columna.HeaderText;
            }
            int indiceFila = 0;
            foreach (DataGridViewRow fila in data.Rows)
            {
                indiceFila++;
                indiceColumna = 0;
                foreach (DataGridViewColumn columna in data.Columns)
                {
                    indiceColumna++;
                    exportarExcel.Cells[indiceFila + 1, indiceColumna] = fila.Cells[columna.Name].Value;
                }
            }
            exportarExcel.Visible = true;
        }
    }
}