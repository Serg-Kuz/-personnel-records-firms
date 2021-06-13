using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kuzmich
{
    public partial class HRForm : Form
    {
        SqlConnection conn;
        SqlDataAdapter ad;
        DataSet ds;
        int id1=0, id2=0, id_acc=0;
        bool admin = true;
        public HRForm()
        {
            InitializeComponent();
        }
        public HRForm(SqlConnection conn, bool admin)
        {
            InitializeComponent();
            this.conn = conn;
            ad = new SqlDataAdapter("select * from Персонал", conn);
            ds = new DataSet();
            ad.Fill(ds, "t");
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Visible = false;
            this.admin = admin;
            if (!admin)
            {
                tabControl1.TabPages.RemoveAt(7);
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count>0)
                switch (tabControl1.SelectedIndex)
                {
                    case 0:
                        {
                            textBox1.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
                            textBox2.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                            textBox3.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox4.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            textBox5.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            break;
                        }
                    case 1:
                        {
                            textBox6.Text= dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
                            break;
                        }
                    case 2:
                        {
                            id1 = (int)dataGridView1.SelectedRows[0].Cells[1].Value;
                            textBox7.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                            textBox8.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox9.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            break;
                        }
                    case 3:
                        {
                            id1 = (int)dataGridView1.SelectedRows[0].Cells[1].Value;
                            textBox10.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                            textBox11.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox12.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            textBox13.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            break;
                        }
                    case 4:
                        {
                            id1 = (int)dataGridView1.SelectedRows[0].Cells[1].Value;
                            textBox14.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                            textBox15.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox16.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            textBox17.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            textBox18.Text = dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
                            textBox19.Text = dataGridView1.SelectedRows[0].Cells[7].Value.ToString();
                            comboBox1.Text = dataGridView1.SelectedRows[0].Cells[8].Value.ToString();
                            comboBox2.Text = dataGridView1.SelectedRows[0].Cells[9].Value.ToString();
                            break;
                        }
                    case 5:
                        {
                            id1 = (int)dataGridView1.SelectedRows[0].Cells[1].Value;
                            id2 = (int)dataGridView1.SelectedRows[0].Cells[2].Value;
                            textBox20.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox21.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            dateTimePicker1.Text= dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            dateTimePicker2.Text= dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
                            comboBox3.Text= dataGridView1.SelectedRows[0].Cells[7].Value.ToString();
                            break;
                        }
                    case 6:
                        {
                            id1 = (int)dataGridView1.SelectedRows[0].Cells[1].Value;
                            textBox22.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                            textBox23.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                            textBox24.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                            textBox25.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            textBox26.Text = dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
                            break;
                        }
                    case 7:
                        {
                            break;
                        }
                    case 8:
                        {
                            break;
                        }
                }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text.Trim()=="")||(textBox2.Text.Trim()=="")||(textBox3.Text.Trim()=="")||(textBox4.Text.Trim()=="")||(textBox5.Text.Trim()==""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Персонал values('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "')";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button11.PerformClick();
                button12.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text.Trim() == "") || (textBox2.Text.Trim() == "") || (textBox3.Text.Trim() == "") || (textBox4.Text.Trim() == "") || (textBox5.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "update Персонал set Фамилия='" + textBox1.Text + "', Имя='" + textBox2.Text + "',Отчество='" + textBox3.Text + "',Адрес='" + textBox4.Text + "',Телефон='" + textBox5.Text + "' where ID_сотрудника=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button11.PerformClick();
                    button12.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text.Trim() == "") || (textBox2.Text.Trim() == "") || (textBox3.Text.Trim() == "") || (textBox4.Text.Trim() == "") || (textBox5.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Персонал where ID_сотрудника=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button11.PerformClick();
                    button12.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = "";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ds = new DataSet();
            ad.Fill(ds, "t");
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Trim() == "")
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Отделы values('" + textBox6.Text + "')";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button14.PerformClick();
                button13.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Trim() == "")
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "update Отделы set Название='" + textBox6.Text + "' where ID_отдела=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button14.PerformClick();
                    button13.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Trim() == "")
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Отделы where ID_отдела=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button14.PerformClick();
                    button13.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox7, "select * from Отделы", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if ((textBox7.Text.Trim() == "") || (textBox8.Text.Trim() == "") || (textBox9.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Должности values("+id1.ToString()+",'" + textBox8.Text + "',"+textBox9.Text+")";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button19.PerformClick();
                button18.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if ((textBox7.Text.Trim() == "")|| (textBox8.Text.Trim() == "")|| (textBox9.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "update Должности set ID_отдела=" + id1.ToString() + ",Название='"+textBox8.Text+"', Оклад="+textBox9.Text+" where ID_должности=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button19.PerformClick();
                    button18.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if ((textBox7.Text.Trim() == "") || (textBox8.Text.Trim() == "") || (textBox9.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Должности where ID_должности=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button19.PerformClick();
                    button18.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            textBox7.Text = textBox8.Text = textBox9.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox10, "select D.ID_должности, D.Название, O.Название Отдел, D.Оклад from Должности D, Отделы O where D.ID_отдела=O.ID_отдела", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button27_Click(object sender, EventArgs e)
        {
            if ((textBox10.Text.Trim()=="")|| (textBox11.Text.Trim() == "")|| (textBox12.Text.Trim() == "")|| (textBox13.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Вакансии values(" + id1.ToString() + ",'" + textBox11.Text + "'," + textBox12.Text+ ",'" + textBox13.Text + "')";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button24.PerformClick();
                button23.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            if ((textBox10.Text.Trim() == "") || (textBox11.Text.Trim() == "") || (textBox12.Text.Trim() == "") || (textBox13.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "update Вакансии set ID_должности=" + id1.ToString() + ",Название='" + textBox11.Text + "',Ставка=" + textBox12.Text + ",Требования='" + textBox13.Text + "' where ID_вакансии=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button24.PerformClick();
                    button23.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if ((textBox10.Text.Trim() == "") || (textBox11.Text.Trim() == "") || (textBox12.Text.Trim() == "") || (textBox13.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Вакансии where ID_вакансии=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString(); ;
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button24.PerformClick();
                    button23.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            textBox10.Text = textBox11.Text = textBox12.Text = textBox13.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox14, "select V.ID_вакансии, D.ID_должности Должность, V.Название, V.Ставка, V.Требования from Должности D, Вакансии V where D.ID_должности=V.ID_должности", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            if ((textBox14.Text.Trim() == "") || (textBox15.Text.Trim() == "") || (textBox16.Text.Trim() == "") || (textBox17.Text.Trim() == "") || (textBox18.Text.Trim() == "") || (textBox19.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Заявки values(" + id1.ToString() + ",'" + textBox15.Text + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox18.Text + "','" + textBox19.Text +"','"+comboBox1.Text+"','"+comboBox2.Text+ "')";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button29.PerformClick();
                button28.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            if ((textBox14.Text.Trim() == "") || (textBox15.Text.Trim() == "") || (textBox16.Text.Trim() == "") || (textBox17.Text.Trim() == "") || (textBox18.Text.Trim() == "") || (textBox19.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "update Заявки set ID_вакансии=" + id1.ToString() + ",Фамилия='" + textBox15.Text + "',Имя='" + textBox16.Text + "',Отчество='" + textBox17.Text + "',Адрес='" + textBox18.Text + "',Телефон='" + textBox19.Text + "',Образование='" + comboBox1.Text + "',[Статус заявки]='" + comboBox2.Text + "' where ID_заявки="+dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button29.PerformClick();
                    button28.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            if ((textBox14.Text.Trim() == "") || (textBox15.Text.Trim() == "") || (textBox16.Text.Trim() == "") || (textBox17.Text.Trim() == "") || (textBox18.Text.Trim() == "") || (textBox19.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Заявки where ID_заявки=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button29.PerformClick();
                    button28.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            textBox14.Text = textBox15.Text = textBox16.Text = textBox17.Text = textBox18.Text = textBox19.Text = "";
            comboBox1.SelectedIndex = comboBox2.SelectedIndex = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox20, "select * from Персонал", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox21, "select D.ID_должности, D.Название, O.Название Отдел, D.Оклад from Должности D, Отделы O where D.ID_отдела=O.ID_отдела", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id2 = form.id;
        }

        private void button37_Click(object sender, EventArgs e)
        {
            if ((textBox20.Text.Trim() == "") || (textBox21.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Контракты values(" + id1.ToString() + "," + id2.ToString()+",'" + dateTimePicker1.Text + "','" + dateTimePicker2.Text + "','" + comboBox3.Text + "')";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button34.PerformClick();
                button33.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            if ((textBox20.Text.Trim() == "") || (textBox21.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if(dataGridView1.SelectedRows.Count>0)
            {
                string query = "update Контракты set ID_сотрудника=" + id1.ToString() + ", ID_должности=" + id2.ToString() + ",Дата_заключения='" + dateTimePicker1.Text + "',Истекает='" + dateTimePicker2.Text + "',Статус_контракта='" + comboBox3.Text + "') where ID_контракта=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button34.PerformClick();
                    button33.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if ((textBox20.Text.Trim() == "") || (textBox21.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Контракты where ID_контракта=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button34.PerformClick();
                    button33.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            textBox20.Text = textBox21.Text = dateTimePicker1.Text = dateTimePicker2.Text = "";
            comboBox3.SelectedIndex = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox22, "select * from Персонал", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button42_Click(object sender, EventArgs e)
        {
            if ((textBox22.Text.Trim() == "") || (textBox23.Text.Trim() == "")|| (textBox24.Text.Trim() == "") || (textBox25.Text.Trim() == "")|| (textBox26.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            string query = "insert into Образование values('" + textBox23.Text + "','" +textBox24.Text + "'," + textBox25.Text + ",'" + textBox26.Text +"',"+ id1.ToString()+")";
            try
            {
                SqlCommand comm = new SqlCommand(query, conn);
                comm.ExecuteNonQuery();
                MessageBox.Show("Добавлено!");
                button39.PerformClick();
                button38.PerformClick();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            if ((textBox22.Text.Trim() == "") || (textBox23.Text.Trim() == "") || (textBox24.Text.Trim() == "") || (textBox25.Text.Trim() == "") || (textBox26.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count>0)
            {
                string query = "update Образование set [Учреждение образования]='" + textBox23.Text + "',[Серия номер диплома]='" + textBox24.Text + "',[Год окончания]=" + textBox25.Text + ",[Специальность по диплому]='" + textBox26.Text + "',ID_сотрудника=" + id1.ToString() + " where ID_образование="+dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Изменено!");
                    button39.PerformClick();
                    button38.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            if ((textBox22.Text.Trim() == "") || (textBox23.Text.Trim() == "") || (textBox24.Text.Trim() == "") || (textBox25.Text.Trim() == "") || (textBox26.Text.Trim() == ""))
            {
                MessageBox.Show("Присутствуют пустые поля");
                return;
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                string query = "delete from Образование where ID_образование=" + dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Удалено!");
                    button39.PerformClick();
                    button38.PerformClick();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            textBox22.Text = textBox23.Text = textBox24.Text = textBox25.Text = textBox26.Text = "";
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            TextBox tb = new TextBox();
            tb = (TextBox)sender;
            double number = 0;
            if (!Double.TryParse(tb.Text, out number))
            {
                tb.Text = "0";
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, textBox27, "select * from Персонал", id1);
            if (form.ShowDialog() == DialogResult.OK)
            {
                id1 = form.id;
                string query = "select * from [Учетные записи] where ID_сотрудника=" + id1.ToString();
                SqlDataAdapter a = new SqlDataAdapter(query, conn);
                DataSet d = new DataSet();
                a.Fill(d, "t");
                if (d.Tables[0].Rows.Count==0)
                {
                    label35.Visible = true;
                    textBox28.Text = textBox29.Text = "";
                }
                else
                {
                    label35.Visible = false;
                    id_acc = (int)d.Tables[0].Rows[0][0];
                    textBox28.Text = d.Tables[0].Rows[0][1].ToString();
                    textBox29.ForeColor = Color.Gray;
                    textBox29.Text = "без изменений";
                    textBox29.PasswordChar = '\0';
                    if (d.Tables[0].Rows[0][3].ToString() == "0")
                        checkBox1.Checked = true;
                    else
                        checkBox1.Checked = false;

                }

            }
            
        }

        private void textBox29_Enter(object sender, EventArgs e)
        {
            textBox29.ForeColor = Color.Black;
            textBox29.Text = "";
            textBox29.PasswordChar = '*';
        }

        private void textBox29_Leave(object sender, EventArgs e)
        {
            if (textBox29.Text.Trim()=="")
            {
                textBox29.ForeColor = Color.Gray;
                textBox29.Text = "без изменений";
                textBox29.PasswordChar = '\0';
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            if ((textBox29.ForeColor==Color.Gray)||(textBox28.Text.Trim()==""))
            {
                MessageBox.Show("Присутствуют пустые поля!");
                return;
            }
            if (label35.Visible)
            {
                string adm = checkBox1.Checked ? "0" : "1";
                string query = "insert into [Учетные записи] values('" + textBox28.Text + "','" + getMD5String(textBox29.Text) + "'," + adm + "," + id1.ToString() + ")";
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Выполнено!");
                    textBox28.Text = "";
                    textBox27.Text = "";
                    textBox29.Text = "";
                    checkBox1.Checked = false;
                    ds = new DataSet();
                    ad.Fill(ds, "t");
                    dataGridView1.DataSource = ds.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
            else
            {
                string adm = checkBox1.Checked ? "0" : "1";
                string query = "update [Учетные записи] set Логин='" + textBox28.Text + "', [Хэш пароля]='" + getMD5String(textBox29.Text) + "',Уровень=" + adm + ",ID_сотрудника=" + id1.ToString() + " where ID_УчетнойЗаписи="+id_acc.ToString();
                try
                {
                    SqlCommand comm = new SqlCommand(query, conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Выполнено!");
                    textBox28.Text = "";
                    textBox27.Text = "";
                    textBox29.Text = "";
                    checkBox1.Checked = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            ListForm form = new ListForm(conn, label36, "select V.ID_вакансии,  V.Название, V.Ставка, V.Требования from Должности D, Вакансии V where D.ID_должности=V.ID_должности", id1);
            if (form.ShowDialog() == DialogResult.OK)
                id1 = form.id;
        }

        private void button45_Click(object sender, EventArgs e)
        {
            if (label36.Text.Trim()=="")
            {
                MessageBox.Show("Не выбрана вакансия!");
                return;
            }
            string query = "select V.Название Вакансия, Z.Фамилия, Z.Имя, Z.Отчество, Z.Адрес, Z.Телефон, Z.Образование, Z.[Статус заявки] from Заявки Z, Вакансии V where Z.ID_вакансии=V.ID_вакансии and V.ID_вакансии="+id1.ToString();
            SqlDataAdapter a = new SqlDataAdapter(query, conn);
            DataSet d = new DataSet();
            a.Fill(d, "t");
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;

            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Exported from gridview";
            for (int i = 1; i < d.Tables[0].Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = d.Tables[0].Columns[i - 1].ColumnName;
            }
            for (int i = 0; i < d.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < d.Tables[0].Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = d.Tables[0].Rows[i][j].ToString();
                }
            }
            worksheet.Columns.AutoFit();
        }

        private void button46_Click(object sender, EventArgs e)
        {
            string query = "select P.Фамилия+' '+P.Имя+' '+P.Отчество Сотрудник, D.Название Должность, K.Дата_заключения, K.Истекает, K.Статус_контракта from Контракты K, Персонал P, Должности D where K.ID_сотрудника=P.ID_сотрудника and K.ID_должности=D.ID_должности and K.Истекает>='"+dateTimePicker3.Text+"'";
            SqlDataAdapter a = new SqlDataAdapter(query, conn);
            DataSet d = new DataSet();
            a.Fill(d, "t");
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;

            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Exported from gridview";
            for (int i = 1; i < d.Tables[0].Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = d.Tables[0].Columns[i - 1].ColumnName;
            }
            for (int i = 0; i < d.Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < d.Tables[0].Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = d.Tables[0].Rows[i][j].ToString();
                }
            }
            worksheet.Columns.AutoFit();
        }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void button47_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            xlWorkSheet.Columns.AutoFit();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    {
                        ad = new SqlDataAdapter("select * from Персонал", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        break;
                    }
                case 1:
                    {
                        ad = new SqlDataAdapter("select * from Отделы", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[0].Visible = false;
                        break;
                    }
                case 2:
                    {
                        ad = new SqlDataAdapter("select D.ID_должности, O.ID_отдела, O.Название Отдел, D.Название Должность, D.Оклад from Должности D, Отделы O where D.ID_отдела=O.ID_отдела", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Visible = false;
                        break;
                    }
                case 3:
                    {
                        ad = new SqlDataAdapter("select V.ID_вакансии, V.ID_должности, D.Название Должность, V.Название, V.Ставка, V.Требования from Вакансии V, Должности D where D.ID_должности=V.ID_должности", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Visible = false;
                        break;
                    }
                case 4:
                    {
                        ad = new SqlDataAdapter("select Z.ID_заявки, Z.ID_вакансии, V.Название Вакансия, Фамилия, Имя, Отчество, Адрес, Телефон, Образование, [Статус заявки] from Заявки Z, Вакансии V where Z.ID_вакансии=V.ID_вакансии", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Visible = false;
                        comboBox1.SelectedIndex = 0;
                        comboBox2.SelectedIndex = 0;
                        break;
                    }
                case 5:
                    {
                        ad = new SqlDataAdapter("select K.ID_контракта, P.ID_сотрудника, D.ID_должности, P.Фамилия+' '+P.Имя+' '+P.Отчество Сотрудник, D.Название Должность, K.Дата_заключения, K.Истекает, K.Статус_контракта from Контракты K, Персонал P, Должности D where K.ID_сотрудника=P.ID_сотрудника and K.ID_должности=D.ID_должности", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Visible = false;
                        dataGridView1.Columns[2].Visible = false;
                        comboBox3.SelectedIndex = 0;
                        break;
                    }
                case 6:
                    {
                        ad = new SqlDataAdapter("select O.ID_образование, P.ID_сотрудника, P.Фамилия+' '+P.Имя+' '+P.Отчество Сотрудник, O.[Учреждение образования],O.[Серия номер диплома], O.[Год окончания], O.[Специальность по диплому] Специальность from Образование O, Персонал P where O.ID_сотрудника=P.ID_Сотрудника", conn);
                        ds = new DataSet();
                        ad.Fill(ds, "t");
                        dataGridView1.DataSource = ds.Tables[0];
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Visible = false;
                        break;
                    }
                case 7:
                    {
                        if (admin)
                        {
                            ad = new SqlDataAdapter("select U.ID_УчетнойЗаписи, P.ID_сотрудника, P.Фамилия+' '+P.Имя+' '+P.Отчество Сотрудник, U.Логин, U.[Хэш пароля] from [Учетные записи] U, Персонал P where U.ID_сотрудника=P.ID_сотрудника", conn);
                            ds = new DataSet();
                            ad.Fill(ds, "t");
                            dataGridView1.DataSource = ds.Tables[0];
                            dataGridView1.Columns[0].Visible = false;
                        }
                        break;
                    }
                default:
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Rows.Clear();
                        break;
                    }
            }
        }
        string getMD5String(string inputString)
        {
            byte[] byteArray = Encoding.ASCII.GetBytes(inputString);
            MD5CryptoServiceProvider md5provider = new MD5CryptoServiceProvider();
            byte[] byteArrayHash = md5provider.ComputeHash(byteArray);
            return BitConverter.ToString(byteArrayHash);
        }
    }
}
