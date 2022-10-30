using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace HowToConnect
{
    public partial class Form1 : Form
    {
        ConectData conectData = new ConectData();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            conectData.Connect();
            DataTable dt = conectData.table("Select * from KhachH");
            dataGridView1.DataSource = dt;
            conectData.closeConnect();
        }
        private bool check()
        {
            if (textBox1.Text.Trim() == "")
            {
                MessageBox.Show("Mã khách hàng không được để trống", "Thông báo");
                return false;
            }
            if (textBox2.Text.Trim() == "")
            {
                MessageBox.Show("Tên khách hàng không được để trống", "Thông báo");
                return false;
            }
            if (textBox3.Text.Trim() == "")
            {
                MessageBox.Show("Địa chỉ khách hàng không được để trống", "Thông báo");
                return false;
            }
            if (textBox4.Text.Trim() == "")
            {
                MessageBox.Show("Số điện thoại khách hàng không được để trống", "Thông báo");
                return false;
            }
            return true;
        }
        private void resetForm()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (check())
            {
                string query = $"Insert into KhachH values(N'{textBox1.Text}',N'{textBox2.Text}',N'{textBox3.Text}',N'{textBox4.Text}')";
                if (MessageBox.Show("Bạn có muốn thêm khách hàng không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    try
                    {
                        conectData.excute(query);
                        MessageBox.Show("Thêm khách hàng thành công!", "Thông báo");
                        Form1_Load(sender, e);
                        resetForm();
                        conectData.closeConnect();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + "Mã khách hàng đã tồn tại vui lòng nhập mã khác!", "Thông báo");
                    }
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (check())
            {
                string query = $"Update KhachH set Ten = N'{textBox2.Text}',DiaChi = N'{textBox3.Text}',DienThoai = N'{textBox4.Text}' where Ma = N'{textBox1.Text.Trim()}'";
                if (MessageBox.Show("Bạn có muốn sửa khách hàng không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    try
                    {
                        conectData.excute(query);
                        MessageBox.Show("Sửa thành công!", "Thông báo");
                        Form1_Load(sender, e);
                        conectData.closeConnect();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex, "Thông báo");
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string s = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
         == System.Windows.Forms.DialogResult.Yes)
            {
                string query = "delete from KhachH where Ma = '" + s + "'";
                conectData.excute(query);
                Form1_Load(sender, e);
                conectData.closeConnect();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("Vui lòng nhập tên khách muốn tìm!", "Thông báo");
                Form1_Load(sender, e);
            }
            else
            {
                string query = $"Select * from KhachH where Ten = N'{textBox2.Text.Trim()}'";
                try
                {
                    dataGridView1.DataSource = conectData.table(query);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
            exSheet.get_Range("B3").Value = "Danh sách khách hàng";
            exSheet.get_Range("A4").Value = "STT";
            exSheet.get_Range("B4").Value = "Ma khách";
            exSheet.get_Range("C4").Value = "Tên khách";
            exSheet.get_Range("D4").Value = "Địa chỉ";
            exSheet.get_Range("E4").Value = "Số điện thoại";
            int n = dataGridView1.Rows.Count;
            for (int i = 0; i < n - 1; i++)
            {
                exSheet.get_Range("A" + (i + 5).ToString()).Value = (i + 1).ToString();
                exSheet.get_Range("B" + (i + 5).ToString()).Value = dataGridView1.Rows[i].Cells[0].Value;
                exSheet.get_Range("C" + (i + 5).ToString()).Value = dataGridView1.Rows[i].Cells[1].Value;
                exSheet.get_Range("D" + (i + 5).ToString()).Value = dataGridView1.Rows[i].Cells[2].Value;
                exSheet.get_Range("E" + (i + 5).ToString()).Value = dataGridView1.Rows[i].Cells[3].Value;
            }
            exBook.Activate();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    exBook.SaveAs(saveFileDialog.FileName.ToString());
                    MessageBox.Show("Xuất file thành công!", "Thông báo");
                    exApp.Quit();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo,
               MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                this.Close();
        }
    }
}
