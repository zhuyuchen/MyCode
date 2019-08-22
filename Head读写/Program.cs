using System;
using System.Collections.Generic;
using System.IO;

namespace Head读写
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime startDT1 = System.DateTime.Now;
            List<List<string>> allhead = new List<List<string>>();
            TextListConverter tl = new TextListConverter();
            allhead = tl.ReadHeadFileToList(@"D:\head.dat");

            DateTime startDT2 = System.DateTime.Now;

            tl.SaveCSV(allhead, @"D:\OBSwells.txt");

            DateTime afterDT = System.DateTime.Now;
            TimeSpan ts1 = afterDT.Subtract(startDT1);
            TimeSpan ts2 = afterDT.Subtract(startDT2);
            Console.WriteLine("处理过程花费{0}ms.", ts1.TotalMilliseconds);
            Console.WriteLine("写文件花费{0}ms.", ts2.TotalMilliseconds);
        }

        /// 文本文件转换为List 
        /// </summary> 
        public class TextListConverter
        {
            //读取文本文件转换为List 
            public List<List<string>> ReadHeadFileToList(string fileName)
            {

                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                List<List<string>> bigList = new List<List<string>>();
                StreamReader sr = new StreamReader(fs);
                
                //使用StreamReader类来读取文件 
                sr.BaseStream.Seek(0, SeekOrigin.Begin);
                // 从数据流中读取每一行，直到文件的最后一行 
                List<string> smallList = new List<string>();
                string tmp = sr.ReadLine();

                while (tmp != null)
                {
                    if(tmp.Contains("TS"))
                    {
                        bigList.Add(smallList);
                        smallList = new List<string>();
                    }
                    smallList.Add(tmp);
                    tmp = sr.ReadLine();
                    //最后一次要在这加入
                    if(tmp==null)
                        bigList.Add(smallList);
                }
                //关闭此StreamReader对象 
                sr.Close();
                fs.Close();
                return bigList;
            }
            
            public void SaveCSV(List<List<string>> content,string CSVPath)
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(CSVPath);
                if (!fi.Directory.Exists)
                {
                    fi.Directory.Create();
                }
                System.IO.FileStream fs = new System.IO.FileStream(CSVPath, FileMode.Create,
                    System.IO.FileAccess.Write);
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.UTF8);
                int[] wells = {
                    26439, 19029, 26889, 27321, 27292, 29452, 31496, 31951, 31524, 28197,
                    29491, 28204, 32727, 33110, 33468, 42373, 38833, 39521, 40200, 36427,
                    37822, 41114, 38826, 39823, 42049, 31558, 43312, 33912, 41788, 42695,
                    30806, 35404, 30763, 32850, 29979, 29114, 45700, 41512, 36547, 39314,
                    39977, 28320, 28748, 24882, 30428, 31225, 25733, 22068, 32525, 26615,
                    32544, 23312, 19210, 22924, 29220, 16403, 19617, 26685, 16443, 35585,
                    21768, 17464, 30125, 32622, 22589, 27579, 27559, 31341, 29318, 22218
                };

                for(int i = 0; i < wells.Length;i++)
                {
                    for (int j = 1; j < content.Count; j++)//j从1开始是因为第0个list存储的说明
                    {
                        sw.Write(content[j][wells[i]]+" ");//这个地方是因为从0开始记录下标，还得去掉表头,0代表的表头
                    }
                    sw.WriteLine();
                }
                sw.Flush();
                sw.Close();
                fs.Close();
                    
                
            }

            ////写入excel
            //public void editExcel(string excelPath)
            //{
            //    Application app = new Application();
            //    object missing = Missing.Value;
            //    Workbook openwb = app.Workbooks.Open(excelPath, missing, false, missing, missing, missing,
            //        missing, missing, missing, false, missing, missing, missing, missing, missing);
            //    Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);
            //    ws.Cells[1, 1] = "朱玉晨";
            //    app.DisplayAlerts = true;//是否显示提示对话框
            //    openwb.Save();//保存工作表
            //    app.Visible = false;//显示Excel
            //    openwb.Close(false, missing, missing);//关闭工作表
            //    app.Quit();

            //}

        }

    }
}
