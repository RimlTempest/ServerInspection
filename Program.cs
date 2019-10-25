using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MultiTaskServerInspection.UI;
using System.Threading;

namespace MultiTaskServerInspection
{
    static class EntryPoint
    {
        public  static MainForm mf = null;
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            string MutexName = "ServerInspectionTool";
            Mutex  mutex = new Mutex(false, MutexName);

            bool hasHandle = false;

            try
            {
                try
                {
                    hasHandle = mutex.WaitOne(0, false);
                }
                catch (AbandonedMutexException)
                {
                    hasHandle = true;
                }

                if (hasHandle == false)
                {
                    MessageBox.Show("多重起動できません",
                            "エラー",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                            );
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            finally
            {
                if (hasHandle)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Close();
            }
        }
    }
}
