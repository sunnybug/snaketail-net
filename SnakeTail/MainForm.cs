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
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

using SnakeTail.Storage;

namespace SnakeTail
{
    partial class MainForm : Form
    {
        private static MainForm _instance = null;
        public static MainForm Instance { get { return _instance; } }

        public string CurrenTailConfig { get { return _currenTailConfig != null ? _currenTailConfig : ""; } }

        private TailFileConfig _defaultTailConfig = null;
        private string _currenTailConfig = null;

        private JWC.MruStripMenu _mruMenu;
        private SnakeTailStorage _storage;

        public MainForm()
        {
            InitializeComponent();
            Icon = Properties.Resources.SnakeIcon;
            _trayIcon.Icon = Properties.Resources.SnakeIcon;
            _instance = this;

            _MDITabControl.ImageList = new ImageList();
            _MDITabControl.ImageList.ImageSize = new System.Drawing.Size(16, 16);
            _MDITabControl.ImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            _MDITabControl.ImageList.TransparentColor = System.Drawing.Color.Transparent;
            _MDITabControl.ImageList.Images.Add(new Bitmap(Properties.Resources.GreenBulletIcon.ToBitmap()));
            _MDITabControl.ImageList.Images.Add(new Bitmap(Properties.Resources.YellowBulletIcon.ToBitmap()));

            // 最近打开文件仅使用 xSnakeTail.db，不使用注册表
            _mruMenu = new JWC.MruStripMenuInline(recentFilesToolStripMenuItem, recentFile1ToolStripMenuItem, new JWC.MruStripMenu.ClickedHandler(OnMruFile), null, false, 10);
            try
            {
                _storage = new SnakeTailStorage(null);
                if (_storage.IsAvailable)
                {
                    List<string> recentFiles = _storage.GetRecentFiles(10);
                    foreach (string file in recentFiles)
                    {
                        _mruMenu.AddFile(file);
                    }
                }
            }
            catch (Exception ex)
            {
                _storage = null;
                MessageBox.Show(this, "无法初始化 SQLite 数据库，最近打开的文件列表将无法保存。\n\n错误: " + ex.Message,
                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateTitle()
        {
            if (IsDisposed || Disposing)
                return;
            try
            {
                Text = Application.ProductName;
                if (_currenTailConfig != null)
                    Text += " - " + Path.GetFileNameWithoutExtension(_currenTailConfig);
            }
            catch
            {
                // 忽略在窗口关闭时更新标题的错误
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                int filesOpened = 0;
                for (int i = 1; i < args.Length; ++i)
                {
                    if (args[1].EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (LoadSession(args[1]))
                            ++filesOpened;
                    }
                    else
                    {
                        filesOpened += OpenFileSelection(new string[] { args[i] });
                    }
                    if (filesOpened == 0 && i >= 2)
                        break;  // Stop attempting to open all arguements if the first two fails
                }
            }
            else
            {
                // 如果没有命令行参数，从 xSnakeTail.db 加载上次会话
                try
                {
                    LoadSessionFromDb();
                }
                catch
                {
                    // 忽略加载错误，继续正常启动
                }
            }
        }

        public void SetStatusBar(string text, int progressValue, int progressMax)
        {
            _statusProgressBar.Maximum = progressMax;
            _statusProgressBar.Value = progressValue;
            if (progressMax == 0 && progressValue == 0)
                _statusProgressBar.Visible = false;
            else
                _statusProgressBar.Visible = true;

            if (text == null)
                text = "Ready";

            if (_statusTextBar.Text != text || progressMax != 0 || progressValue != 0)
            {
                _statusTextBar.Text = text;
                _statusStrip.Invalidate();
                _statusStrip.Update();
            }
        }

        private void MainForm_MdiChildActivate(object sender, EventArgs e)
        {
            closeItemToolStripMenuItem.Enabled = this.ActiveMdiChild != null;

            // If no any child form, hide tabControl
            if (this.ActiveMdiChild == null)
            {
                if (_MDITabControl.TabCount==0)
                    _MDITabControl.Visible = false;
            }
            else
            {
                // If child form is new and no has tabPage, create new tabPage
                if (this.ActiveMdiChild.Tag == null)
                {
                    // Add a tabPage to tabControl with child form caption
                    AddMdiChildTab(this.ActiveMdiChild);

                    if (MdiChildren.Length > 1 && _MDITabControl.Visible == false)
                        return;

                    // Child form always maximized
                    this.ActiveMdiChild.WindowState = FormWindowState.Maximized;

                    _MDITabControl.SelectedTab = this.ActiveMdiChild.Tag as TabPage;
                }
                else
                {
                    if (_MDITabControl.Visible == false)
                        return;

                    TabPage tp = this.ActiveMdiChild.Tag as TabPage;
                    if (tp != null)
                    {
                        // Child form always maximized
                        this.ActiveMdiChild.WindowState = FormWindowState.Maximized;

                        _MDITabControl.SelectedTab = tp;
                    }
                }

                if (!_MDITabControl.Visible)
                    _MDITabControl.Visible = true;
            }
        }

        void AddMdiChildTab(Form mdiChild)
        {
            TabPage tp = new TabPage(mdiChild.Text);
            tp.Tag = mdiChild;
            tp.Parent = _MDITabControl;
            //AddOwnedForm(mdiChild);
            mdiChild.Tag = tp;
            mdiChild.FormClosed += new FormClosedEventHandler(ActiveMdiChild_FormClosed);
            mdiChild.SizeChanged += new EventHandler(ActiveMdiChild_SizeChanged);
            mdiChild.Shown += new EventHandler(ActiveMdiChild_Shown);
        }

        void ActiveMdiChild_Shown(object sender, EventArgs e)
        {
            // Fix the icon when starting MDI child in maximized state
            if ((sender as Form).WindowState == FormWindowState.Maximized)
            {
                ActivateMdiChild(null);
                ActivateMdiChild((sender as Form));
            }
        }

        void ActiveMdiChild_SizeChanged(object sender, EventArgs e)
        {
            // Disable tab-mode if the active MDI child changes WindowState
            if (this.ActiveMdiChild == sender && this.ActiveMdiChild.WindowState != FormWindowState.Maximized)
            {
                // Check if we are about to open / close a window
                if (MdiChildren.Length == _MDITabControl.TabCount)
                {
                    if (_MDITabControl.SelectedTab == null || this.ActiveMdiChild == _MDITabControl.SelectedTab.Tag)
                    {
                        _MDITabControl.Visible = false;
                        SetStatusBar(null, 0, 0);
                    }
                }
            }
        }

        private void ActiveMdiChild_FormClosed(object sender, FormClosedEventArgs e)
        {
            ((sender as Form).Tag as TabPage).Dispose();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "Open Log File";
            fileDialog.Filter = "Default Filter|*.txt;*.text;*.log*;*.xlog|Log Files|*.log*;*.xlog|Text Files|*.txt;*.text|All Files|*.*";
            if (fileDialog.ShowDialog(this) != DialogResult.OK)
                return;

            OpenFileSelection(fileDialog.FileNames);
        }

        private void OnMruFile(int number, String filename)
        {
            bool openedFile = false;
            if (filename.EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase))
                openedFile = LoadSession(filename);
            else
                openedFile = OpenFileSelection(new string[] { filename }) == 1;

            if (!openedFile)
            {
                MessageBox.Show(this, "The file '" + filename + "'cannot be opened and will be removed from the Recent list(s)", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (_mruMenu != null)
                {
                    _mruMenu.RemoveFile(number);
                }
                if (_storage != null)
                {
                    _storage.RemoveFile(filename);
                }
            }
        }

        public int OpenFileSelection(string[] filenames)
        {
            if (_defaultTailConfig == null)
            {
                TailConfig tailConfig = _storage?.LoadDefaultSession();
                if (tailConfig != null && tailConfig.TailFiles.Count > 0)
                {
                    _defaultTailConfig = tailConfig.TailFiles[0];
                    _defaultTailConfig.Title = null;
                }
                else
                {
                    _defaultTailConfig = new TailFileConfig();
                }
            }

            int filesOpened = 0;
            foreach (string filename in filenames)
            {
                string configPath = "";
                try
                {
                    if (string.IsNullOrEmpty(Path.GetDirectoryName(filename)))
                        configPath = Directory.GetCurrentDirectory();
                }
                catch
                {
                }

                TailForm mdiForm = new TailForm();
                TailFileConfig tailConfig = _defaultTailConfig;
                tailConfig.FilePath = filename;
                // Auto-detect encoding when opening a file
                if (File.Exists(filename))
                {
                    Encoding detectedEncoding = EncodingHelper.DetectFileEncoding(filename);
                    if (detectedEncoding != null)
                    {
                        tailConfig.EnumFileEncoding = detectedEncoding;
                    }
                }
                mdiForm.LoadConfig(tailConfig, configPath);
                if (mdiForm.IsDisposed)
                    continue;

                try
                {
                    string fullPath = filename;
                    if (string.IsNullOrEmpty(configPath))
                    {
                        new DirectoryInfo(Path.GetDirectoryName(filename));
                    }
                    else
                    {
                        fullPath = Path.Combine(configPath, filename);
                    }

                    // 添加到菜单
                    if (_mruMenu != null)
                    {
                        _mruMenu.AddFile(fullPath);
                    }

                    // 保存到 SQLite（如果可用）
                    if (_storage != null)
                    {
                        _storage.AddFile(fullPath);
                    }
                }
                catch
                {
                }

                mdiForm.MdiParent = this;
                mdiForm.Show();
                ++filesOpened;
                Application.DoEvents();
            }
            return filesOpened;
        }

        private void openEventLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenEventLogDialog openEventLogDlg = new OpenEventLogDialog();
            if (openEventLogDlg.ShowDialog(this) != DialogResult.OK)
                return;

            EventLogForm mdiForm = new EventLogForm();
            mdiForm.MdiParent = this;
            mdiForm.LoadFile(openEventLogDlg.EventLogFile);
            if (!mdiForm.IsDisposed)
                mdiForm.Show();
        }

        private void _MDITabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((_MDITabControl.SelectedTab != null) && (_MDITabControl.SelectedTab.Tag != null))
            {
                SuspendLayout();
                (_MDITabControl.SelectedTab.Tag as Form).SuspendLayout();
                Form activeMdiChild = this.ActiveMdiChild;
                if (activeMdiChild != null)
                    activeMdiChild.SuspendLayout();
                // Minimize flicker when switching between tabs, by changing to minimized state first
                if ((_MDITabControl.SelectedTab.Tag as Form).WindowState != FormWindowState.Maximized)
                    (_MDITabControl.SelectedTab.Tag as Form).WindowState = FormWindowState.Minimized;
                (_MDITabControl.SelectedTab.Tag as Form).Select();
                if (activeMdiChild != null && !activeMdiChild.IsDisposed)
                    activeMdiChild.ResumeLayout();
                (_MDITabControl.SelectedTab.Tag as Form).ResumeLayout();
                ResumeLayout();
                (_MDITabControl.SelectedTab.Tag as Form).Refresh();
            }
        }

        private void _MDITabControl_MouseClick(object sender, MouseEventArgs e)
        {
            var tabControl = sender as TabControl;
            TabPage tabPageCurrent = GetTabPageFromLocation(tabControl, e.Location);

            if (e.Button == MouseButtons.Middle)
            {
                if (tabPageCurrent != null)
                    (tabPageCurrent.Tag as Form).Close();
            }
            else if (e.Button == MouseButtons.Right)
            {
                var enablePath = tabPageCurrent.Tag is TailForm;
                _openFolderTabContext.Visible = enablePath;
                _copyPathTabContext.Visible = enablePath;
                _separatorTabContext.Visible = enablePath;

                _tabContextMenuStrip.Show(sender as TabControl, e.Location);
            }
        }

        private void cascadeWindowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void tileWindowsHorizontallyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void tileWindowsVerticallyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void minimizeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form activeChild = ActiveMdiChild;
            foreach (Form childForm in MdiChildren)
            {
                if (childForm.WindowState != FormWindowState.Minimized)
                    childForm.WindowState = FormWindowState.Minimized;
            }
            if (activeChild != null && activeChild != ActiveMdiChild)
                activeChild.Select();
        }

        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _MDITabControl.Visible = false;
            FormCollection forms = Application.OpenForms;
            for (int i = forms.Count - 1; i >= 0; i--)
            {
                ITailForm tailForm = forms[i] as ITailForm;
                if (tailForm != null)
                    tailForm.TailWindow.Close();
            }
            if (SearchForm.Instance.Visible)
                SearchForm.Instance.Close();
            _currenTailConfig = null;
            UpdateTitle();
        }

        private void enableTabsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_MDITabControl.Visible)
            {
                _MDITabControl.Visible = false;
            }
            else
            if (this.ActiveMdiChild != null)
            {
                this.ActiveMdiChild.WindowState = FormWindowState.Maximized;
                _MDITabControl.Visible = true;
                _MDITabControl.SelectedTab = this.ActiveMdiChild.Tag as TabPage;
            }
        }

        private void saveSessionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (!String.IsNullOrEmpty(_currenTailConfig))
            {
                saveFileDialog.FileName = Path.GetFileName(_currenTailConfig);
                saveFileDialog.InitialDirectory = Path.GetDirectoryName(_currenTailConfig);
            }
            saveFileDialog.Filter = "Xml files (*.xml)|*.xml|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                SaveSession(saveFileDialog.FileName);
            }
        }

        private void loadSessionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Xml files (*.xml)|*.xml|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                LoadSession(openFileDialog.FileName);
            }
        }

        private void SaveSession(string filepath)
        {
            TailConfig tailConfig = new TailConfig();
            if (_MDITabControl != null && _MDITabControl.Visible)
                tailConfig.SelectedTab = _MDITabControl.SelectedIndex;
            else
                tailConfig.SelectedTab = -1;
            tailConfig.WindowSize = Size;
            tailConfig.WindowPosition = DesktopLocation;
            tailConfig.MinimizedToTray = _trayIcon != null && _trayIcon.Visible;
            tailConfig.AlwaysOnTop = TopMost;

            List<Form> childForms = new List<Form>();

            // We first loop through the tabpages to store in proper TabPage order
            if (_MDITabControl != null && !_MDITabControl.IsDisposed)
            {
                foreach (TabPage tagPage in _MDITabControl.TabPages)
                {
                    Form tailForm = tagPage.Tag as Form;
                    if (tailForm != null)
                        childForms.Add(tailForm);
                }
            }

            // Then we loop through all forms (includes free floating)
            foreach (Form childForm in Application.OpenForms)
            {
                if (childForms.IndexOf(childForm) == -1)
                    childForms.Add(childForm);
            }

            // Save all forms and store in proper order
            foreach (Form childForm in childForms)
            {
                ITailForm tailForm = childForm as ITailForm;
                if (tailForm != null)
                {
                    TailFileConfig tailFile = new TailFileConfig();
                    tailForm.SaveConfig(tailFile);
                    tailConfig.TailFiles.Add(tailFile);
                }
            }

            SaveConfig(tailConfig, filepath);

            if (String.IsNullOrEmpty(_currenTailConfig))
            {
                if (_mruMenu != null)
                {
                    _mruMenu.AddFile(filepath);
                }
                if (_storage != null)
                {
                    _storage.AddFile(filepath);
                }
            }
            else if (_currenTailConfig != filepath)
            {
                if (_mruMenu != null)
                {
                    _mruMenu.RenameFile(_currenTailConfig, filepath);
                }
                if (_storage != null)
                {
                    _storage.RenameFile(_currenTailConfig, filepath);
                }
            }

            _currenTailConfig = filepath;

            UpdateTitle();
        }

        /// <summary>
        /// 将会话保存到指定 XML 文件（仅用于“另存为”）
        /// </summary>
        public void SaveConfig(TailConfig tailConfig, string filepath)
        {
            if (string.IsNullOrEmpty(filepath))
                return;

            XmlSerializer serializer = new XmlSerializer(typeof(TailConfig));
            try
            {
                using (XmlTextWriter writer = new XmlTextWriter(filepath, Encoding.UTF8))
                {
                    writer.Formatting = Formatting.Indented;
                    XmlSerializerNamespaces xmlnsEmpty = new XmlSerializerNamespaces();
                    xmlnsEmpty.Add("", "");
                    serializer.Serialize(writer, tailConfig, xmlnsEmpty);
                }

                _defaultTailConfig = null;
            }
            catch (System.Exception ex)
            {
                string errorMsg = ex.Message;
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                    errorMsg += "\n" + ex.Message;
                }
                MessageBox.Show(this, "Failed to save session xml file, please ensure it is valid location:\n\n   " + filepath + "\n\n" + errorMsg, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TailConfig LoadSessionFile(string filepath)
        {
            TailConfig tailConfig = null;
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(TailConfig));
                using (XmlTextReader reader = new XmlTextReader(filepath))
                {
                    _currenTailConfig = new Uri(reader.BaseURI).LocalPath;
                    tailConfig = serializer.Deserialize(reader) as TailConfig;
                }
                return tailConfig;
            }
            catch (Exception ex)
            {
                string errorMsg = ex.Message;
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                    errorMsg += "\n" + ex.Message;
                }
                MessageBox.Show(this, "Failed to open session xml file, please ensure it is valid file:\n\n   " + filepath + "\n\n" + errorMsg, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private bool LoadSession(string filepath, bool addToMru = true)
        {
            TailConfig tailConfig = LoadSessionFile(filepath);
            if (tailConfig == null)
                return false;

            if (addToMru)
            {
                if (_mruMenu != null)
                    _mruMenu.AddFile(filepath);
                if (_storage != null)
                    _storage.AddFile(filepath);
            }

            return LoadSessionWithConfig(tailConfig, Path.GetDirectoryName(filepath));
        }

        /// <summary>
        /// 从 xSnakeTail.db 加载上次会话
        /// </summary>
        private void LoadSessionFromDb()
        {
            TailConfig tailConfig = _storage?.LoadDefaultSession();
            if (tailConfig == null)
                return;

            string basePath = Path.GetDirectoryName(Application.ExecutablePath) ?? "";
            LoadSessionWithConfig(tailConfig, basePath);
        }

        /// <summary>
        /// 将会话保存到 xSnakeTail.db
        /// </summary>
        private void SaveSessionToDb()
        {
            TailConfig tailConfig = BuildCurrentTailConfig();
            if (tailConfig == null || tailConfig.TailFiles.Count == 0)
                return;

            _storage?.SaveDefaultSession(tailConfig);
        }

        private TailConfig BuildCurrentTailConfig()
        {
            TailConfig tailConfig = new TailConfig();
            if (_MDITabControl != null && _MDITabControl.Visible)
                tailConfig.SelectedTab = _MDITabControl.SelectedIndex;
            else
                tailConfig.SelectedTab = -1;
            tailConfig.WindowSize = Size;
            tailConfig.WindowPosition = DesktopLocation;
            tailConfig.MinimizedToTray = _trayIcon != null && _trayIcon.Visible;
            tailConfig.AlwaysOnTop = TopMost;

            List<Form> childForms = new List<Form>();
            if (_MDITabControl != null && !_MDITabControl.IsDisposed)
            {
                foreach (TabPage tagPage in _MDITabControl.TabPages)
                {
                    Form tailForm = tagPage.Tag as Form;
                    if (tailForm != null)
                        childForms.Add(tailForm);
                }
            }
            foreach (Form childForm in Application.OpenForms)
            {
                if (childForms.IndexOf(childForm) == -1)
                    childForms.Add(childForm);
            }
            foreach (Form childForm in childForms)
            {
                ITailForm tailForm = childForm as ITailForm;
                if (tailForm != null)
                {
                    TailFileConfig tailFile = new TailFileConfig();
                    tailForm.SaveConfig(tailFile);
                    tailConfig.TailFiles.Add(tailFile);
                }
            }
            return tailConfig;
        }

        private bool LoadSessionWithConfig(TailConfig tailConfig, string configBasePath)
        {
            if (tailConfig == null)
                return false;

            if (!tailConfig.MinimizedToTray)
            {
                Size = tailConfig.WindowSize;
                DesktopLocation = tailConfig.WindowPosition;
            }

            UpdateTitle();

            List<string> eventLogFiles = EventLogForm.GetEventLogFiles();
            Application.DoEvents();

            foreach (TailFileConfig tailFile in tailConfig.TailFiles)
            {
                Form mdiForm = null;
                int index = eventLogFiles.FindIndex(delegate (string arrItem) { return arrItem.Equals(tailFile.FilePath); });
                if (index >= 0)
                    mdiForm = new EventLogForm();
                else
                    mdiForm = new TailForm();

                if (mdiForm != null)
                {
                    ITailForm tailForm = mdiForm as ITailForm;
                    mdiForm.Text = tailFile.Title;
                    if (!tailFile.Modeless)
                    {
                        mdiForm.MdiParent = this;
                        mdiForm.ShowInTaskbar = false;
                        AddMdiChildTab(mdiForm);
                        if (tailForm != null)
                            tailForm.LoadConfig(tailFile, configBasePath);
                        if (mdiForm.IsDisposed)
                        {
                            _MDITabControl.TabPages.Remove(mdiForm.Tag as TabPage);
                            continue;
                        }
                    }
                    mdiForm.Show();

                    if (tailConfig.SelectedTab == -1 || tailFile.Modeless)
                    {
                        if (tailFile.WindowState != FormWindowState.Maximized)
                        {
                            mdiForm.DesktopLocation = tailFile.WindowPosition;
                            mdiForm.Size = tailFile.WindowSize;
                        }
                        if (mdiForm.WindowState != tailFile.WindowState)
                            mdiForm.WindowState = tailFile.WindowState;
                    }

                    if (tailFile.Modeless && tailForm != null)
                        tailForm.LoadConfig(tailFile, configBasePath);
                }
                Application.DoEvents();
            }

            if (tailConfig.SelectedTab != -1 && _MDITabControl.TabPages.Count > 0)
            {
                foreach (Form childForm in MdiChildren)
                    childForm.WindowState = FormWindowState.Minimized;
                _MDITabControl.SelectedIndex = tailConfig.SelectedTab;
                _MDITabControl.Visible = true;
                (_MDITabControl.SelectedTab.Tag as Form).WindowState = FormWindowState.Maximized;
            }

            if (tailConfig.MinimizedToTray)
            {
                _trayIcon.Visible = true;
                WindowState = FormWindowState.Minimized;
                Visible = false;
            }
            else if (tailConfig.AlwaysOnTop)
            {
                alwaysOnTopToolStripMenuItem.Checked = true;
                TopMost = true;
            }

            return true;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void windowToolStripMenuItem_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripMenuItem menuItem = e.ClickedItem as ToolStripMenuItem;
            if (menuItem != null && menuItem.IsMdiWindowListEntry)
            {
                // If a minimized window is chosen from the list, then it is restored to normal state
                this.windowToolStripMenuItem.DropDownItemClicked -= windowToolStripMenuItem_DropDownItemClicked;
                e.ClickedItem.PerformClick();
                if (ActiveMdiChild != null && ActiveMdiChild.WindowState == FormWindowState.Minimized)
                    ActiveMdiChild.WindowState = FormWindowState.Normal;
                this.windowToolStripMenuItem.DropDownItemClicked += windowToolStripMenuItem_DropDownItemClicked;
            }
        }

        private void minimizeToTrayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!_trayIcon.Visible)
            {
                _trayIcon.Visible = true;
                WindowState = FormWindowState.Minimized;
                Visible = false;
                minimizeToTrayToolStripMenuItem.Checked = true;
                _trayIcon.ShowBalloonTip(3, "Minimized to tray", "Double click the system tray icon to restore window", ToolTipIcon.Info);
            }
            else
            {
                Visible = true;
                WindowState = FormWindowState.Normal;
                _trayIcon.Visible = false;
                minimizeToTrayToolStripMenuItem.Checked = false;
            }
        }

        private void alwaysOnTopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TopMost = !TopMost;
            alwaysOnTopToolStripMenuItem.Checked = TopMost;
        }

        private void _trayIcon_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
         }

        private void windowToolStripMenuItem_DropDownOpening(object sender, EventArgs e)
        {
            enableTabsToolStripMenuItem.Checked = _MDITabControl.Visible;
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog(this);
        }

        private void checkForUpdateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using(new HourGlass(this))
                {
                    CheckForUpdates updateChecker = new CheckForUpdates();
                    updateChecker.PromptAlways = true;
                    updateChecker.Check();
                }
            }
            catch (Exception ex)
            {
                ThreadExceptionDialog dlg = new ThreadExceptionDialog(ex);
                dlg.Text = "Error checking for new updates";
                dlg.ShowDialog(this);
            }
        }

        private void _trayIconContextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            // We steal the items from the main menu (we restore them when closing again)
            ToolStripItem[] items = new ToolStripItem[fileToolStripMenuItem.DropDownItems.Count];
            fileToolStripMenuItem.DropDownItems.CopyTo(items, 0);
            _trayIconContextMenuStrip.Items.Clear();            // Clear the dummy item
            _trayIconContextMenuStrip.Items.AddRange(items);
            minimizeToTrayToolStripMenuItem.Checked = true;
            minimizeToTrayToolStripMenuItem.Font = new Font(minimizeToTrayToolStripMenuItem.Font, FontStyle.Bold);
        }

        private void _trayIconContextMenuStrip_Closed(object sender, ToolStripDropDownClosedEventArgs e)
        {
            // Restore the items back to the main menu when closing
            ToolStripItem[] items = new ToolStripItem[_trayIconContextMenuStrip.Items.Count];
            _trayIconContextMenuStrip.Items.CopyTo(items, 0);
            fileToolStripMenuItem.DropDownItems.AddRange(items);
            _trayIconContextMenuStrip.Items.Clear();
            _trayIconContextMenuStrip.Items.Add(new ToolStripSeparator());  // Dummy item so menu is shown the next time
            minimizeToTrayToolStripMenuItem.Checked = false;
            minimizeToTrayToolStripMenuItem.Font = new Font(minimizeToTrayToolStripMenuItem.Font, FontStyle.Regular);
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                try
                {
                    // 尝试获取文件列表以验证权限
                    Array array = e.Data.GetData(DataFormats.FileDrop) as Array;
                    if (array != null && array.Length > 0)
                    {
                        // 检查第一个文件是否可访问（不实际打开，只检查路径）
                        string firstFile = array.GetValue(0).ToString();
                        if (!string.IsNullOrEmpty(firstFile) && System.IO.Path.IsPathRooted(firstFile))
                        {
                            e.Effect = DragDropEffects.Copy;
                            return;
                        }
                    }
                }
                catch
                {
                    // 如果权限检查失败，仍然允许拖拽，让 DragDrop 处理错误
                }
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                Array array = e.Data.GetData(DataFormats.FileDrop) as Array;
                if (array == null)
                    return;

                // Extract strings from array
                List<string> filenames = new List<string>();
                foreach(object filename in array)
                {
                    filenames.Add(filename.ToString());
                }

                this.Activate();        // in the case Explorer overlaps this form

                // Call OpenFile asynchronously.
                // Explorer instance from which file is dropped is not responding
                // all the time when DragDrop handler is active, so we need to return
                // immidiately (especially if OpenFile shows MessageBox).
                System.Threading.ThreadPool.QueueUserWorkItem(worker_DoWork, filenames.ToArray());
            }
            catch (Exception ex)
            {
                // don't show MessageBox here - Explorer is waiting !
                System.Diagnostics.Debug.WriteLine("Drag Drop Failed: " + ex.Message);
            }
        }

        void worker_DoWork(object param)
        {
            // Discovered a strange problem where the Windows Explorer would lock, eventhough I deferred the actual DragDrop operation using BeginInvoke().
            // The solution was to create a thread, that slept for 100 ms and then invoked the wanted method. If I removed the sleep from the new thread,
            // then Windows Explorer would lock again. Very strange indeed.
            System.Threading.Thread.Sleep(100);
            this.BeginInvoke(new Action<string[]>(delegate(string[] filenames)
            {
                try
                {
                    // 验证文件路径和权限
                    List<string> validFiles = new List<string>();
                    List<string> invalidFiles = new List<string>();

                    foreach (string filename in filenames)
                    {
                        try
                        {
                            if (string.IsNullOrEmpty(filename))
                                continue;

                            // 检查路径是否有效
                            if (!System.IO.Path.IsPathRooted(filename))
                            {
                                invalidFiles.Add(filename);
                                continue;
                            }

                            // 尝试访问文件信息以验证权限
                            // 对于目录，检查是否存在
                            if (System.IO.Directory.Exists(filename))
                            {
                                // 目录暂时不支持，跳过
                                continue;
                            }

                            // 对于文件，检查是否存在或是否可以访问
                            if (System.IO.File.Exists(filename))
                            {
                                // 尝试获取文件信息以验证权限
                                System.IO.FileInfo fileInfo = new System.IO.FileInfo(filename);
                                // 如果文件存在但无法访问，会在 OpenFileSelection 中处理
                                validFiles.Add(filename);
                            }
                            else
                            {
                                // 文件不存在，但可能是新文件或需要特殊权限
                                // 尝试检查父目录权限
                                string dir = System.IO.Path.GetDirectoryName(filename);
                                if (!string.IsNullOrEmpty(dir) && System.IO.Directory.Exists(dir))
                                {
                                    validFiles.Add(filename);
                                }
                                else
                                {
                                    invalidFiles.Add(filename);
                                }
                            }
                        }
                        catch (System.UnauthorizedAccessException)
                        {
                            invalidFiles.Add(filename + " (权限不足)");
                        }
                        catch (System.Security.SecurityException)
                        {
                            invalidFiles.Add(filename + " (安全权限不足)");
                        }
                        catch (Exception ex)
                        {
                            invalidFiles.Add(filename + " (" + ex.Message + ")");
                        }
                    }

                    // 打开有效的文件
                    if (validFiles.Count > 0)
                    {
                        OpenFileSelection(validFiles.ToArray());
                    }

                    // 显示无效文件的错误消息
                    if (invalidFiles.Count > 0)
                    {
                        string errorMsg = "以下文件无法打开：\n\n" + string.Join("\n", invalidFiles.ToArray());
                        MessageBox.Show(this, errorMsg, "文件拖拽失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "打开文件时发生错误：\n\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }), new object[] { param });
        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            if (_trayIcon.Visible && WindowState == FormWindowState.Minimized)
                Visible = false;
        }

        private void _MDITabControl_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _MDITabControl.AllowDrop = true;
                _MDITabControl.DoDragDrop(_MDITabControl.SelectedTab, DragDropEffects.All);
                _MDITabControl.AllowDrop = false;
            }
        }

        private void _MDITabControl_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(TabPage)))
                e.Effect = DragDropEffects.Move;
            else
                e.Effect = DragDropEffects.None;
        }

        private void _MDITabControl_DragDrop(object sender, DragEventArgs e)
        {
            Point clientPoint = _MDITabControl.PointToClient(new Point(e.X, e.Y));
            for(int i = 0; i < _MDITabControl.TabPages.Count; ++i)
            {
                if (_MDITabControl.GetTabRect(i).Contains(clientPoint))
                {
                    if (_MDITabControl.TabPages[i] == _MDITabControl.SelectedTab)
                        break;  // No change

                    TabPage tabPage = _MDITabControl.SelectedTab;
                    _MDITabControl.TabPages.Remove(tabPage);
                    _MDITabControl.TabPages.Insert(i, tabPage);
                    _MDITabControl.SelectedIndex = i;
                    break;
                }
            }
        }

        private void clearListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_mruMenu != null)
            {
                _mruMenu.RemoveAll();
            }
            if (_storage != null)
            {
                _storage.ClearAllRecentFiles();
            }
        }


        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                // 自动保存当前会话到 xSnakeTail.db（不再写入 XML）
                try
                {
                    bool hasOpenFiles = false;
                    foreach (Form childForm in Application.OpenForms)
                    {
                        ITailForm tailForm = childForm as ITailForm;
                        if (tailForm != null)
                        {
                            hasOpenFiles = true;
                            break;
                        }
                    }
                    if (hasOpenFiles)
                        SaveSessionToDb();
                }
                catch
                {
                    // 忽略保存会话的错误，不影响程序关闭
                }

                // 清理不存在的文件记录并释放 SQLite 资源
                if (_storage != null)
                {
                    try
                    {
                        _storage.CleanupNonExistentFiles();
                    }
                    catch
                    {
                        // 忽略清理错误
                    }

                    try
                    {
                        _storage.Dispose();
                    }
                    catch
                    {
                        // 忽略释放错误
                    }
                    finally
                    {
                        _storage = null;
                    }
                }
            }
            catch(Exception ex)
            {
                // 使用 null 作为 owner，避免在窗口关闭时出现问题
                try
                {
                    MessageBox.Show(null, "Failed to save list of recently used files.\n\n" + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                    // 如果 MessageBox 也失败，则完全忽略
                }
            }
            finally
            {
                // 最后才将 _instance 设置为 null，确保异常处理可以访问它
                _instance = null;
            }
        }

        private void closeItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ActiveMdiChild != null)
            {
                ActiveMdiChild.Close();
            }
        }

        private TForm GetSelectedTabForm<TForm>(object sender) where TForm : Form
        {
            ToolStripItem item = (sender as ToolStripItem);
            if (item != null)
            {
                ContextMenuStrip owner = item.Owner as ContextMenuStrip;
                if (owner != null)
                {
                    var sourceControl = owner.SourceControl;
                    var tabControl = sourceControl as TabControl;
                    if (tabControl != null)
                    {
                        var relativeToScreen = tabControl.PointToClient(owner.Bounds.Location);
                        var tabPageCurrent = GetTabPageFromLocation(tabControl, relativeToScreen);
                        if (tabPageCurrent != null)
                        {
                            return tabPageCurrent.Tag as TForm;
                        }
                    }
                }
            }
            return null;
        }

        private TabPage GetTabPageFromLocation(TabControl tabControl, Point point)
        {
            for (var i = 0; i < tabControl.TabCount; i++)
            {
                if (!tabControl.GetTabRect(i).Contains(point))
                    continue;
                return tabControl.TabPages[i];
            }
            return null;
        }

        private void _copyFolderPathClick(object sender, EventArgs e)
        {
            TailForm tailForm = GetSelectedTabForm<TailForm>(sender);
            if (tailForm != null)
            {
                tailForm.CopyPath();
            }
        }

        private void _closeContextClick(object sender, EventArgs e)
        {
            Form form = GetSelectedTabForm<Form>(sender);
            if (form != null)
            {
                form.Close();
            }
        }

        private void _openContainingFolderClick(object sender, EventArgs e)
        {
            TailForm tailForm = GetSelectedTabForm<TailForm>(sender);
            if (tailForm != null)
            {
                tailForm.OpenExplorer();
            }
        }
    }
}
