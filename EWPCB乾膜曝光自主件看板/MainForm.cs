using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace EWPCB乾膜曝光自主件看板
{
    public partial class MainForm : Form
    {
        string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";
        string strComm = "";
        string LogPath = Directory.GetCurrentDirectory() + @"\ErrLog.txt";
        DataTable srcData = new DataTable();
        DateTime date = DateTime.Now.AddDays(-2);
        DateTime updateTime = DateTime.Now;
        int setTime = 180;
        int clock = 0;
        StreamWriter writeLog;
        public MainForm()
        {
            InitializeComponent();
            dgvData.ReadOnly = true;
            DataRefresh(date.ToString("yyyy-MM-dd 00:00:00"));
            clock = setTime;
            Text = "乾膜曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            tmrRefresh.Interval = 1000;
            tmrRefresh.Start();
        }

        private void DataRefresh(string date)
        {
            srcData.Clear();
            writeLog = File.AppendText(LogPath);
            strComm = "select partnum as '料號', ISNULL(machineno,'0') + '-' + ISNULL(empname,'') as '曝光手', " +
                "workqnty as '曝光數', CONVERT(char(19), starttime, 120) as '曝光時間', " +
                "CONVERT(char(19), vrs, 120) as 'VRS時間', ISNULL(REPLACE(vrsman,' ',''), '') as 'VRS人員'," +
                "ISNULL(qcresult, '') as '結果' from drymcse where departname = 'FF' and process = '自主件' and todo = 1 and " +
                "starttime >= '" + date + "' order by starttime desc";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader read = sqlcomm.ExecuteReader();
                        srcData.Load(read);
                        /*
                        writeLog.WriteLine(updateTime + "     " + "Update OK！");
                        writeLog.Flush();
                        */
                        writeLog.Close();
                    }
                    catch (Exception ex)
                    {
                        writeLog.WriteLine(updateTime + "     " + ex.Message);
                        writeLog.Flush();
                        writeLog.Close();
                    }
                }
            }

            //先將DataTable[結果]欄位的唯讀取消
            srcData.Columns["結果"].ReadOnly = false;

            //檢查是否進2D工序
            for (int i = 0; i < srcData.Rows.Count; i++)
            {
                if (Check2D(srcData.Rows[i]["料號"].ToString()))
                {
                    srcData.Rows[i]["結果"] = "2D";
                }
            }

            //1920*1080 only 43" TV
            dgvData.DataSource = srcData;
            dgvData.RowHeadersWidth = 70;
            dgvData.Columns["料號"].Width = 285;
            dgvData.Columns["曝光手"].Width = 195;
            dgvData.Columns["曝光數"].Width = 165;
            dgvData.Columns["曝光時間"].Width = 400;
            dgvData.Columns["VRS時間"].Width = 400;
            dgvData.Columns["VRS人員"].Width = 190;
            dgvData.Columns["結果"].Width = 250;
            dgvData.DataBindingComplete += ChangRowColor;
        }

        //在資料繫結後，所觸發的事件
        private void ChangRowColor(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (string.IsNullOrWhiteSpace(row.Cells["結果"].Value.ToString()) & 
                    !string.IsNullOrWhiteSpace(row.Cells["VRS時間"].Value.ToString()))
                {
                    row.DefaultCellStyle.BackColor = Color.Lime;
                    row.DefaultCellStyle.SelectionBackColor = Color.Lime;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (row.Cells["結果"].Value.ToString().Trim().Contains("曝偏"))
                {
                    row.DefaultCellStyle.BackColor = Color.Magenta;
                    row.DefaultCellStyle.SelectionBackColor = Color.Magenta;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (row.Cells["結果"].Value.ToString().Trim().ToUpper().Contains("Mylar沒撕"))
                {
                    row.DefaultCellStyle.BackColor = Color.Red;
                    row.DefaultCellStyle.SelectionBackColor = Color.Red;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (row.Cells["結果"].Value.ToString().Trim().ToUpper().Contains("顯影不良"))
                {
                    row.DefaultCellStyle.BackColor = Color.MediumOrchid;
                    row.DefaultCellStyle.SelectionBackColor = Color.MediumOrchid;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.White;
                    row.DefaultCellStyle.SelectionBackColor = Color.White;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
            }
        }

        private void tmrRefresh_Tick(object sender, EventArgs e)
        {
            clock--;
            Text = "乾膜曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            if (clock == 0)
            {
                date = DateTime.Now;
                updateTime = DateTime.Now;
                DataRefresh(date.ToString("yyyy-MM-dd 00:00:00"));
                clock = setTime;
                Text = "乾膜曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }

        /// <summary>
        /// 檢查料號在二天內有無被申報過2D工序，若有傳回true，表示底片漲縮需進2D掃描
        /// </summary>
        /// <param name="PartNum">料號</param>
        /// <returns></returns>
        private bool Check2D(string PartNum)
        {
            var Result = false;
            var strComm = "select * from drymcse where departname='FF' and process='2D' and " +
                "partnum='" + PartNum + "' and starttime > DATEADD(DAY, -1, sysdatetime())";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader reader = sqlcomm.ExecuteReader();
                        if (reader.HasRows)
                        {
                            Result = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        writeLog.WriteLine(updateTime + "     " + ex.Message);
                        writeLog.Flush();
                        writeLog.Close();
                    }
                }
            }
            return Result;
        }
    }
}
