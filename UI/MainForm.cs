using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultiTaskServerInspection.UI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void InspectionStart_Click(object sender, EventArgs e)
        {
            InspectionForm inspection = null;
            try
            {
                if (inspection == null || inspection.IsDisposed)
                {
                    this.Cursor = Cursors.WaitCursor;
                    inspection = new InspectionForm();
                    inspection.Show();

                }
            }
            catch (Exception ex)
            {
                String errorMesssage;
                errorMesssage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                MessageBox.Show(errorMesssage, "例外エラー");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void LogViewer_Click(object sender, EventArgs e)
        {
           
        }

        private void AppplicationClose_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("アプリケーションを終了してよろしいですか？",
               "警告",
               MessageBoxButtons.YesNo,
               MessageBoxIcon.Warning,
               MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
