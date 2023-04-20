using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabWork3
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            
        }

        private void btnInflation_Click(object sender, EventArgs e)
        {

        }

        private void btnSalary_Click(object sender, EventArgs e)
        {

        }

        private void btnPopulation_Click(object sender, EventArgs e)
        {

        }

        // Смена языка на енгельский (Если идея - хуита, то можно потом делитнуть)
        private void btnLangEN_Click(object sender, EventArgs e)
        {
            if (btnInflation.Text != "Inflation" || btnPopulation.Text != "Population" || btnSalary.Text != "Median salary")
            {
                btnInflation.Text = "Inflation";
                btnPopulation.Text = "Population";
                btnSalary.Text = "Median Salary";
            }
        }

        // Смена языка на русский
        private void btnLangRU_Click(object sender, EventArgs e)
        {
            if (btnInflation.Text == "Inflation" || btnPopulation.Text == "Population" || btnSalary.Text == "Median salary")
            {
                btnInflation.Text = "Инфляция";
                btnPopulation.Text = "Население";
                btnSalary.Text = "Медианная зарплата";
            }
        }
    }
}
