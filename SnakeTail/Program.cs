#region License statement
/* SnakeTail is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
#endregion

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace SnakeTail
{
    static class Program
    {
        static volatile bool applicationCrashed = false;
        /// <summary>GitHub Releases API（用于检查更新），与 publish 工作流发布的 tag 一致。</summary>
        public static readonly string GitHubReleasesApiUrl = "https://api.github.com/repos/sunnybug/snaketail-net/releases/latest";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                ApplicationConfiguration.Initialize();

                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                Application.EnableVisualStyles();
                Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.Automatic);
                Application.SetCompatibleTextRenderingDefault(false);

                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                string path = Path.Combine(Path.GetTempPath(), "SnakeTail_startup_error.txt");
                try
                {
                    File.WriteAllText(path, ex.ToString());
                }
                catch { }
                MessageBox.Show(ex.ToString(), "SnakeTail 启动异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (applicationCrashed)
                return;
            applicationCrashed = true;

            if (e.ExceptionObject is Exception)
                SendCrashReport(e.ExceptionObject as Exception);
            else
                SendCrashReport(new Exception(string.Format("Unknown Exception - {0}", e.ExceptionObject)));

            applicationCrashed = false;
        }

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            if (applicationCrashed)
                return;
            applicationCrashed = true;

            SendCrashReport(e.Exception);

            applicationCrashed = false;
        }

        static void SendCrashReport(Exception ex)
        {
            try
            {
                if (ex == null)
                    ex = new Exception("Unknown exception (null reference)");

                ThreadExceptionDialogEx dlg = new ThreadExceptionDialogEx(ex);
                if (MainForm.Instance != null && !MainForm.Instance.IsDisposed)
                    dlg.ShowDialog(MainForm.Instance);
                else
                    dlg.ShowDialog();
            }
            catch (Exception dialogEx)
            {
                string path = Path.Combine(Path.GetTempPath(), "SnakeTail_crash_error.txt");
                try
                {
                    File.WriteAllText(path, ex?.ToString() + Environment.NewLine + "--- Dialog failed ---" + Environment.NewLine + dialogEx.ToString());
                }
                catch { }
                MessageBox.Show(
                    "无法显示错误报告对话框。" + Environment.NewLine + Environment.NewLine +
                    "原始异常: " + (ex?.Message ?? "null") + Environment.NewLine + Environment.NewLine +
                    "详细已写入: " + path,
                    Application.ProductName + " - Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
