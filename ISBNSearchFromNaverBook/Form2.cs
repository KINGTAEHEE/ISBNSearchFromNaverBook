using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISBNSearchFromNaverBook
{
    public partial class Form2 : Form
    {
        DataTable dt = new DataTable();
        string saveFile = string.Empty;

        public Form2(DataTable _dt)
        {
            InitializeComponent();

            dt = _dt;
            this.dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 60;
            dataGridView1.Columns[2].Width = 60;
            dataGridView1.Columns[3].Width = 60;
            dataGridView1.Columns[4].Width = 60;
            dataGridView1.Columns[5].Width = 60;
            dataGridView1.Columns[6].Width = 90;
            dataGridView1.Columns[7].Width = 340;
            dataGridView1.Columns[8].Width = 150;
            dataGridView1.Columns[9].Width = 150;
            dataGridView1.Columns[10].Width = 70;
            dataGridView1.Columns[11].Width = 60;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(dt);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ExportToExcel(DataTable _dt)
        {
            if (ShowFileSaveDialog())
            {
                button1.Enabled = false;
                button2.Enabled = false;
                float progressBarStep = 100 / (float)(_dt.Rows.Count);

                try
                {
                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Workbooks.Add();
                    Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelApp.ActiveSheet;

                    // 컬럼명 입력
                    for (var i = 0; i < _dt.Columns.Count; i++)
                    {
                        workSheet.Cells[1, i + 1] = _dt.Columns[i].ColumnName;
                    }

                    // DataTable 데이터를 Excel 셀에 입력
                    for (var i = 0; i < _dt.Rows.Count; i++)
                    {
                        for (var j = 0; j < _dt.Columns.Count; j++)
                        {
                            workSheet.Cells[i + 2, j + 1] = _dt.Rows[i][j];
                        }
                        progressBar1.Value = (int)((i + 1) * progressBarStep);
                    }

                    workSheet.SaveAs(saveFile);
                    excelApp.Quit();

                    progressBar1.Value = 0;
                    MessageBox.Show("저장 완료");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            button1.Enabled = true;
            button2.Enabled = true;
        }

        public bool ShowFileSaveDialog()
        {
            // 파일 오픈창 생성 및 설정
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "엑셀 파일 (*.xls)|*.xls";

            DialogResult dr = sfd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                saveFile = sfd.FileName;
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
