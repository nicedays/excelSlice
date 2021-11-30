using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using MSExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;






/***
 *  anthor:nicedays   ikevinsama@gmail.com  *
 *    **/
namespace 农邦分表器
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = GetValue("NUM");
         }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }


        private void label1_Click(object sender, EventArgs e)
        {
            
        }
        private void label1_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void label1_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            //其中label1.Text显示的就是拖进文件的文件名；

            label1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }
        private void label2_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void label2_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            //其中label1.Text显示的就是拖进文件的文件名；

            label2.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetValue("NUM", textBox1.Text.ToString());
            
            CutExcel(label1.Text.ToString(), label2.Text.ToString());
        }

        /// <summary>
        /// 分割excel
        /// 
        /// </summary>
        /// <param name="rule">规则表</param>
        /// <param name="date">总表</param>
        /// <returns></returns>
        public int CutExcel(string rule,string date)
        {
            FileInfo fi = new FileInfo(rule);
            FileInfo fi2 = new FileInfo(date);
            if (fi.Exists && fi2.Exists)
            {
                Console.WriteLine(label1.Text.ToString());
                //创建
                MSExcel.Application xlApp = new MSExcel.Application();
                xlApp.DisplayAlerts = false;
                xlApp.Visible = false;
                xlApp.ScreenUpdating = false;
                //打开Excel
                MSExcel.Workbook xlsWorkBook = xlApp.Workbooks.Open(rule, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

                //处理数据过程
                MSExcel.Worksheet m_Sheet = xlsWorkBook.Worksheets[1];//工作薄从1开始，不是0
                int Flag = m_Sheet.UsedRange.Rows.Count;
                int Flag2 = m_Sheet.UsedRange.Columns.Count;
                string[][] ruleSheet = new string[Flag2][];
                //读取规则excel表得内容存在数组当中
                for (int i = 0; i < Flag2; i++)
                {
                    ruleSheet[i] = new string[Flag];
                    for (int j = 0; j < Flag - 1; j++)
                    {
                        ruleSheet[i][j] = m_Sheet.Cells[j + 2, i + 1].Text;

                    }
                }
                button1.Text = "数据处理中，请等待";
                string Pathstr = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase+"xlsx";
                if (!Directory.Exists(Pathstr))
                {

                    DirectoryInfo directoryInfo = new DirectoryInfo(Pathstr);
                    directoryInfo.Create();
                }
                
                //分割表
                for (int i = 0; i < m_Sheet.UsedRange.Columns.Count; i++)
                {
                    //打开总表开始处理数据
                    MSExcel.Workbook xlsWorkBook2 = xlApp.Workbooks.Open(date, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                    System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                    System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                    MSExcel.Worksheet m_Sheet2 = xlsWorkBook2.Worksheets[1];//工作薄从1开始，不是0
                    int row_flag = m_Sheet2.UsedRange.Rows.Count ;//记录行号

                    for (int j = 2; j < row_flag+2; j++)
                    {
                        string PeisongAddress = m_Sheet2.Cells[j, int.Parse(textBox1.Text.ToString())].text;
                        
                        //判断配送点是否在地区当中,不在就删除行
                        if (((IList)ruleSheet[i]).Contains(PeisongAddress)&&(PeisongAddress!=""))
                        {

                            //Range mergeArea = m_Sheet2.Range[m_Sheet2.Cells[17,6], m_Sheet2.Cells[17, 6]].MergeArea;
                            //Console.WriteLine(mergeArea.Cells.Rows.Count);
                            //new_Sheet.Range[new_Sheet.Cells[row_flag, 1]
                            //    , new_Sheet.Cells[row_flag, m_Sheet2.UsedRange.Columns.Count]]
                            //    .Copy(m_Sheet2.Range[m_Sheet2.Cells[j + 1, 1], m_Sheet2.Cells[j + 1, m_Sheet2.UsedRange.Columns.Count]]);
                            //m_Sheet2.Range[m_Sheet2.Cells[j + 1, 1], m_Sheet2.Cells[j + 1, m_Sheet2.UsedRange.Columns.Count]].Copy();
                            //new_Sheet.Range[row_flag, 1].PasteSpecial();
                            //Range data_range = m_Sheet2.Range[m_Sheet2.Cells[j + 1, 1], m_Sheet2.Cells[j + 1, m_Sheet2.UsedRange.Columns.Count]];
                            // Range new_range = new_Sheet.Range[new_Sheet.Cells[row_flag, 1]
                            //    , new_Sheet.Cells[row_flag, m_Sheet2.UsedRange.Columns.Count]];
                            // data_range.Copy(Type.Missing);
                            // new_range.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            // Console.WriteLine(m_Sheet2.Cells[j + 1, 1].text);
                            // row_flag++;


                        }
                        else if((!((IList)ruleSheet[i]).Contains(PeisongAddress)) && (PeisongAddress != ""))
                        {
                            
                            m_Sheet2.Rows[j].Delete();
                            
                            j--;
                        }
                        else
                        {

                        }
                    }
                    string FilePath=Pathstr+"\\"+m_Sheet.Cells[1,i+1].Text+ DateTime.Now.ToString("yyyyMMddhhmm") + ".XLSX";
                    FilePath = FilePath.Replace("\\\\", "\\");
                    FileInfo fii = new FileInfo(FilePath);
                    if (fii.Exists)     //判断文件是否已经存在,如果存在就删除!
                    {
                        fii.Delete();
                    }
                    xlsWorkBook2.SaveAs(fii, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    xlsWorkBook2.Close(true, Type.Missing, Type.Missing);
                    // PrintSheet(m_Sheet2);
                    // Console.WriteLine("====");

                }



                button1.Text = "处理完毕";
                ClosePro(xlApp, xlsWorkBook);


            }
            return 0;
        }

        /// <summary>
        /// 打印一个表的所有信息
        /// </summary>
        /// <param name="m_Sheet">sheet</param>
        public void PrintSheet(MSExcel.Worksheet m_Sheet)
        {
            
            
                for (int i = 0; i < m_Sheet.UsedRange.Rows.Count; i++)
                {
                    for(int j = 0; j < m_Sheet.UsedRange.Columns.Count; j++)
                    {
                        Console.Write(m_Sheet.Cells[i+1,j+1].Text + "\t");
                    }
                    Console.WriteLine("");
                }
                
        }

        /// <summary>
        /// 配置文件键值对
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public static void SetValue(string key, string value)
        {
            //增加的内容写在appSettings段下 <add key="RegCode" value="0"/>  
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (config.AppSettings.Settings[key] == null)
            {
                config.AppSettings.Settings.Add(key, value);
            }
            else
            {
                config.AppSettings.Settings[key].Value = value;
            }
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");//重新加载新的配置文件   
        }
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        /// <summary>
        /// 关闭进程
        /// </summary>
        /// <param name="xlApp"></param>
        /// <param name="xlsWorkBook"></param>
        public void ClosePro(MSExcel.Application xlApp, MSExcel.Workbook xlsWorkBook)
        {
            if (xlsWorkBook != null)
                xlsWorkBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();
            // 安全回收进程
            System.GC.GetGeneration(xlApp);
            IntPtr t = new IntPtr(xlApp.Hwnd);   //获取句柄
            int k = 0;
            GetWindowThreadProcessId(t, out k);   //获取进程唯一标志
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();     //关闭进程
        }
        /// <summary>  
        /// 读取文件配置信息
        /// </summary>  
        /// <param name="key"></param>  
        /// <returns></returns>  
        public static string GetValue(string key)
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (config.AppSettings.Settings[key] == null)
                return "";
            else
                return config.AppSettings.Settings[key].Value;
        }
        
    }
}
