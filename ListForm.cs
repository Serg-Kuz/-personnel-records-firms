using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kuzmich
{
    public partial class ListForm : Form
    {
        SqlConnection conn;
        Control c1, c2, c3, c4;
        SqlDataAdapter ad;
        DataSet ds;
        string table;
        bool mode = false;
        public int id;
        public ListForm()
        {
            InitializeComponent();
        }
        public ListForm(SqlConnection conn, Control c, string query, int id)
        {
            InitializeComponent();
            this.conn = conn;
            this.mode = true;
            this.c1 = c;
            ad = new SqlDataAdapter(query, conn);
            ds = new DataSet();
            ad.Fill(ds, "table");
            dataGridView1.DataSource = ds.Tables[0];
            this.id = id;
            dataGridView1.Columns[0].Visible = false;
        }
        public ListForm(SqlConnection conn, Control c1, Control c2, Control c3, Control c4, string query)
        {
            InitializeComponent();
            this.conn = conn;

            this.c1 = c1;
            this.c2 = c2;
            this.c3 = c3;
            this.c4 = c4;
            ad = new SqlDataAdapter(query, conn);
            ds = new DataSet();
            ad.Fill(ds, "table");
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Visible = false;
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[0].Value.ToString());
            if (mode)
            {
                c1.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            }
            else
            {
                c1.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
                c2.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                c3.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                if (dataGridView1.SelectedRows[0].Cells.Count > 4)
                    c4.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            }
            this.DialogResult = DialogResult.OK;

        }
    }
}
