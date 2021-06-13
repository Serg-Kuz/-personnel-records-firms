using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kuzmich
{
    public partial class Connect : Form
    {
        public Connect()
        {
            InitializeComponent();
        }
        private void Connect_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.this_pc)
            {
                checkBox1.Checked = true;
                comboBox1.Enabled = false;
            }
            else
            {
                checkBox1.Checked = true;
                comboBox1.Text = Properties.Settings.Default.ip_db;
            }
            checkBox2.Checked = Properties.Settings.Default.alter_login;
            if (checkBox2.Checked)
            {
                textBox1.ReadOnly = false;
                textBox2.ReadOnly = false;
                textBox1.Text = Properties.Settings.Default.login;
                textBox2.Text = Properties.Settings.Default.password;
            }
            checkBox1.CheckedChanged += CheckBox1_CheckedChanged;
            checkBox2.CheckedChanged += CheckBox2_CheckedChanged;
        }
        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox1.ReadOnly = !checkBox2.Checked;
            textBox2.ReadOnly = !checkBox2.Checked;
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                comboBox1.Enabled = true;
                this.Cursor = Cursors.WaitCursor;
                DirectoryEntry parent = new DirectoryEntry("WinNT:");
                foreach (DirectoryEntry dm in parent.Children)
                {
                    DirectoryEntry coParent = new DirectoryEntry("WinNT://" + dm.Name);
                    DirectoryEntries dent = coParent.Children;
                    dent.SchemaFilter.Add("Computer");
                    foreach (DirectoryEntry client in dent)
                    {
                        comboBox1.Items.Add(client.Name);
                    }
                }
                this.Cursor = Cursors.Default;

            }
            else
            {
                this.Cursor = Cursors.Default;
                comboBox1.Enabled = false;

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.this_pc = checkBox1.Checked;
            Properties.Settings.Default.ip_db = comboBox1.Text;
            Properties.Settings.Default.alter_login = checkBox2.Checked;
            Properties.Settings.Default.login = textBox1.Text;
            Properties.Settings.Default.password = textBox2.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
