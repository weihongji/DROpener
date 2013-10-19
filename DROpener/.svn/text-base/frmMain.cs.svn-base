using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace OpenDailyReport
{
	public partial class frmMain : Form
	{
		String m_ReportName; // Path & Name

		public frmMain()
		{
			InitializeComponent();
			this.StartPosition = FormStartPosition.Manual;
			this.Location = new Point(-1000, -1000);
		}

		private void frmMain_Load(object sender, EventArgs e) {
			this.lblMessage.Text = "Openning daily report file ...";
			OpenReportFile();
			this.lblMessage.Text = "Daily report file is open.";
			Application.Exit();
		}

		private void OpenReportFile() {
			if (this.ReportName.Length > 0) {
				ProcessStartInfo pInfo = new ProcessStartInfo();
				pInfo.FileName = this.ReportName;
				pInfo.UseShellExecute = true;
				Process p = Process.Start(pInfo);
			}
			else {
				MessageBox.Show("Daily report file is not found. File to open:\n\n\t" + m_ReportName + "\n\nPlease check if the report main folder path is correct in the file \"dr.txt\".", "Daily Report Openner", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public String ReportPath {
			get {
				string reportPath;

				if (File.Exists("dr.txt")) {
					reportPath = File.ReadAllText("dr.txt").Trim();
				}
				else {
					reportPath = Directory.GetCurrentDirectory();
					File.WriteAllText("dr.txt", reportPath);
				}

				return reportPath;
			}
			set {
				File.WriteAllText("dr.txt", value);
			}
		}

		public String ReportName {
			get {
				String folderName, fileName;

				folderName = DateTime.Today.ToString("MMMM", CultureInfo.InvariantCulture);
				fileName = DateTime.Today.ToString("yyyyMMdd")+".xls";
				m_ReportName = this.ReportPath + "\\" + folderName + "\\" + fileName;

				if (File.Exists(m_ReportName)) {
					return m_ReportName;
				}
				else {
					return "";
				}
			}
		}
	}
}