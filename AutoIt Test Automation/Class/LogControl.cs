using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;


namespace Golfzon_TVNX_Test
{
    class LogControl
    {
        static string logFileName;

        public void SetLogFileName()
        {
            // 로그파일이름 생성
            logFileName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";
            WriteLogWithoutBox(logFileName + " 로그 파일 이름 생성 완료");
        }

        public void Write(string logText)
        {
            // 로그 앞부분 날짜
            string date = DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss] ");


            // 로그 앞부분에 메소드 네임 기록
            string methodName = new StackFrame(1, true).GetMethod().Name + " :: ";                
            string text = date + methodName + logText;


            // 메인폼 로그텍스트박스에 기록
            RichTextBox TXTlog = FORMmain.mainForm.TXTlog;

            TXTlog.Text += text + "\r\n";
            TXTlog.SelectionStart = TXTlog.Text.Length;
            TXTlog.ScrollToCaret();


            // 파일 쓰기
            WriteFile(text);
        }


        public void WriteLog(object logBox, string logText)  // 텍스트박스와 파일쓰기를 동시에 하는 메소드
        {
            // 로그 텍스트 창에 로그 출력
            RichTextBox TXTlog = logBox as RichTextBox;

            TXTlog.Text += logText + "\r\n";
            TXTlog.SelectionStart = TXTlog.Text.Length;
            TXTlog.ScrollToCaret();
                        
            // 파일 쓰기
            WriteFile(logText);
        }

        public void WriteLogWithoutBox(string logText)   // 파일쓰기만 하는 메소드
        {
            WriteFile(logText);
        }

        private void WriteFile(string logText)    // 로그쓰는 메소드
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + @"\log");
            if (!di.Exists)
                di.Create();

            FileInfo fi = new FileInfo(di.FullName + @"\" + logFileName);
            FileStream fs = new FileStream(fi.FullName, FileMode.Append, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

            try
            {
                // 자신을 호출한 클래스를 받아옴
                // StackTrace st = new StackTrace();
                // StackFrame sf = st.GetFrame(st.FrameCount - 3);
                // MethodBase mb = sf.GetMethod();

                string date = DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss] "); // + "[" + mb.Name + "] ";
                sw.WriteLine(logText);
            }
            catch (Exception error)
            {
                MessageBox.Show("Class.LogControl 에러\r\n" + error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            sw.Flush();
            sw.Close();
            fs.Close();
        }
    }
}
