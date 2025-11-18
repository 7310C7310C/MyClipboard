using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;

namespace MyClipboard
{
    public class ClipboardItem
    {
        public DateTime Time { get; set; }
        public string Text { get; set; }
        public string Format { get; set; }
        public byte[] Data { get; set; }
        
        public override string ToString()
        {
            if (!string.IsNullOrEmpty(Text))
            {
                string preview = Text.Length > 50 ? Text.Substring(0, 50) + "..." : Text;
                preview = preview.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                return string.Format("[{0}] {1}", Time.ToString("HH:mm:ss"), preview);
            }
            return string.Format("[{0}] {1}", Time.ToString("HH:mm:ss"), Format);
        }
    }

    public class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private ListView listView;
        private ContextMenuStrip listContextMenu;
        private List<ClipboardItem> clipboardHistory = new List<ClipboardItem>();
        private string dataPath = @"C:\ProgramData\MyClipboard\history.dat";
        private IDataObject lastClipboardData = null;
        private System.Windows.Forms.Timer clipboardTimer;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const int MOD_ALT = 0x0001;
        private const int MOD_CONTROL = 0x0002;
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 1;
        private const byte VK_X = 0x58;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        public MainForm()
        {
            InitializeComponents();
            LoadHistory();
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_X);
            
            // 启动剪贴板监控
            clipboardTimer = new System.Windows.Forms.Timer();
            clipboardTimer.Interval = 500;
            clipboardTimer.Tick += ClipboardTimer_Tick;
            clipboardTimer.Start();
        }

        private void InitializeComponents()
        {
            // 主窗体设置
            this.Text = "MyClipboard";
            this.Size = new Size(400, 600);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            
            // 添加边框效果
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.Padding = new Padding(1);

            // 内容面板
            Panel contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.BackColor = Color.FromArgb(30, 30, 30);
            this.Controls.Add(contentPanel);

            // ListView设置
            listView = new ListView();
            listView.Dock = DockStyle.Fill;
            listView.View = View.Details;
            listView.FullRowSelect = true;
            listView.GridLines = false;
            listView.HeaderStyle = ColumnHeaderStyle.None;
            listView.BackColor = Color.FromArgb(30, 30, 30);
            listView.ForeColor = Color.White;
            listView.BorderStyle = BorderStyle.None;
            listView.Columns.Add("Content", 380);
            listView.DoubleClick += ListView_DoubleClick;
            listView.MouseClick += ListView_MouseClick;
            contentPanel.Controls.Add(listView);

            // 右键菜单
            listContextMenu = new ContextMenuStrip();
            listContextMenu.Items.Add("复制", null, CopyItem_Click);
            listContextMenu.Items.Add("删除", null, DeleteItem_Click);
            listView.ContextMenuStrip = listContextMenu;

            // 托盘图标
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application; // 默认图标，如果有ico文件会在后面替换
            trayIcon.Visible = true;
            trayIcon.Text = "MyClipboard";
            trayIcon.MouseClick += TrayIcon_MouseClick;

            // 尝试加载自定义图标
            try
            {
                string exePath = Application.ExecutablePath;
                string exeDir = Path.GetDirectoryName(exePath);
                string icoPath = Path.Combine(exeDir, "icon.ico");
                
                if (File.Exists(icoPath))
                {
                    trayIcon.Icon = new Icon(icoPath);
                    this.Icon = new Icon(icoPath);
                }
                else
                {
                    // 从exe中提取图标
                    this.Icon = Icon.ExtractAssociatedIcon(exePath);
                    trayIcon.Icon = Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch
            {
                // 使用默认图标
            }

            // 窗体事件
            this.Load += MainForm_Load;
            this.Deactivate += MainForm_Deactivate;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // 启动时隐藏窗口
            this.Hide();
            
            // 设置窗口位置在屏幕右侧
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(workingArea.Right - this.Width - 10, workingArea.Top + 50);
        }

        private void MainForm_Deactivate(object sender, EventArgs e)
        {
            // 失去焦点时隐藏
            this.Hide();
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ToggleWindow();
            }
        }

        private void ToggleWindow()
        {
            if (this.Visible)
            {
                this.Hide();
            }
            else
            {
                ShowWindow();
            }
        }

        private void ShowWindow()
        {
            // 更新位置到鼠标附近或屏幕右侧
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            Point cursorPos = Cursor.Position;
            
            int x = Math.Min(cursorPos.X, workingArea.Right - this.Width);
            int y = Math.Min(cursorPos.Y, workingArea.Bottom - this.Height);
            
            // 优先在屏幕右侧显示
            x = workingArea.Right - this.Width - 10;
            y = Math.Max(workingArea.Top + 50, Math.Min(y, workingArea.Bottom - this.Height));
            
            this.Location = new Point(x, y);
            this.Show();
            this.Activate();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                ToggleWindow();
            }
            base.WndProc(ref m);
        }

        private void ClipboardTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Clipboard.ContainsData(DataFormats.Text) || 
                    Clipboard.ContainsData(DataFormats.UnicodeText) ||
                    Clipboard.ContainsData(DataFormats.Rtf) ||
                    Clipboard.ContainsData(DataFormats.Html) ||
                    Clipboard.ContainsImage())
                {
                    IDataObject currentData = Clipboard.GetDataObject();
                    
                    if (IsNewClipboardData(currentData))
                    {
                        lastClipboardData = currentData;
                        AddClipboardItem(currentData);
                    }
                }
            }
            catch
            {
                // 剪贴板访问失败时忽略
            }
        }

        private bool IsNewClipboardData(IDataObject newData)
        {
            if (lastClipboardData == null)
                return true;

            // 比较文本内容
            if (newData.GetDataPresent(DataFormats.UnicodeText) && 
                lastClipboardData.GetDataPresent(DataFormats.UnicodeText))
            {
                string newText = newData.GetData(DataFormats.UnicodeText) as string;
                string lastText = lastClipboardData.GetData(DataFormats.UnicodeText) as string;
                return newText != lastText;
            }

            return true;
        }

        private void AddClipboardItem(IDataObject data)
        {
            ClipboardItem item = new ClipboardItem();
            item.Time = DateTime.Now;

            // 优先处理文本
            if (data.GetDataPresent(DataFormats.UnicodeText))
            {
                item.Text = data.GetData(DataFormats.UnicodeText) as string;
                item.Format = "Text";
            }
            else if (data.GetDataPresent(DataFormats.Text))
            {
                item.Text = data.GetData(DataFormats.Text) as string;
                item.Format = "Text";
            }
            else if (data.GetDataPresent(DataFormats.Rtf))
            {
                string rtf = data.GetData(DataFormats.Rtf) as string;
                item.Data = Encoding.UTF8.GetBytes(rtf);
                item.Format = "RTF";
                item.Text = "富文本内容";
            }
            else if (data.GetDataPresent(DataFormats.Html))
            {
                string html = data.GetData(DataFormats.Html) as string;
                item.Data = Encoding.UTF8.GetBytes(html);
                item.Format = "HTML";
                item.Text = "HTML内容";
            }
            else if (Clipboard.ContainsImage())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    Image img = Clipboard.GetImage();
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    item.Data = ms.ToArray();
                }
                item.Format = "Image";
                item.Text = "图片";
            }
            else
            {
                return; // 不支持的格式
            }

            // 避免重复
            if (clipboardHistory.Count > 0)
            {
                var last = clipboardHistory[0];
                if (last.Text == item.Text && last.Format == item.Format)
                    return;
            }

            clipboardHistory.Insert(0, item);
            RefreshListView();
            SaveHistory();
        }

        private void RefreshListView()
        {
            listView.Items.Clear();
            foreach (var item in clipboardHistory)
            {
                ListViewItem lvi = new ListViewItem(item.ToString());
                lvi.Tag = item;
                listView.Items.Add(lvi);
            }
        }

        private void ListView_DoubleClick(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count > 0)
            {
                ClipboardItem item = listView.SelectedItems[0].Tag as ClipboardItem;
                PasteItem(item);
            }
        }

        private void PasteItem(ClipboardItem item)
        {
            try
            {
                // 设置剪贴板内容
                if (item.Format == "Text")
                {
                    Clipboard.SetText(item.Text);
                }
                else if (item.Format == "RTF")
                {
                    string rtf = Encoding.UTF8.GetString(item.Data);
                    Clipboard.SetData(DataFormats.Rtf, rtf);
                }
                else if (item.Format == "HTML")
                {
                    string html = Encoding.UTF8.GetString(item.Data);
                    Clipboard.SetData(DataFormats.Html, html);
                }
                else if (item.Format == "Image")
                {
                    using (MemoryStream ms = new MemoryStream(item.Data))
                    {
                        Image img = Image.FromStream(ms);
                        Clipboard.SetImage(img);
                    }
                }

                // 隐藏窗口
                this.Hide();

                // 等待一下让窗口完全隐藏
                Thread.Sleep(100);

                // 模拟 Ctrl+V 粘贴
                keybd_event(0x11, 0, 0, UIntPtr.Zero); // Ctrl down
                keybd_event(0x56, 0, 0, UIntPtr.Zero); // V down
                keybd_event(0x56, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // V up
                keybd_event(0x11, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // Ctrl up
            }
            catch (Exception ex)
            {
                MessageBox.Show("粘贴失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ListView_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (listView.SelectedItems.Count > 0)
                {
                    listContextMenu.Show(listView, e.Location);
                }
            }
        }

        private void CopyItem_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count > 0)
            {
                ClipboardItem item = listView.SelectedItems[0].Tag as ClipboardItem;
                try
                {
                    if (item.Format == "Text")
                    {
                        Clipboard.SetText(item.Text);
                    }
                    else if (item.Format == "RTF")
                    {
                        string rtf = Encoding.UTF8.GetString(item.Data);
                        Clipboard.SetData(DataFormats.Rtf, rtf);
                    }
                    else if (item.Format == "HTML")
                    {
                        string html = Encoding.UTF8.GetString(item.Data);
                        Clipboard.SetData(DataFormats.Html, html);
                    }
                    else if (item.Format == "Image")
                    {
                        using (MemoryStream ms = new MemoryStream(item.Data))
                        {
                            Image img = Image.FromStream(ms);
                            Clipboard.SetImage(img);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("复制失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void DeleteItem_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count > 0)
            {
                ClipboardItem item = listView.SelectedItems[0].Tag as ClipboardItem;
                clipboardHistory.Remove(item);
                RefreshListView();
                SaveHistory();
            }
        }

        private void SaveHistory()
        {
            try
            {
                string dir = Path.GetDirectoryName(dataPath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                using (FileStream fs = new FileStream(dataPath, FileMode.Create))
                using (BinaryWriter writer = new BinaryWriter(fs))
                {
                    writer.Write(clipboardHistory.Count);
                    foreach (var item in clipboardHistory)
                    {
                        writer.Write(item.Time.ToBinary());
                        writer.Write(item.Format ?? "");
                        writer.Write(item.Text ?? "");
                        
                        if (item.Data != null)
                        {
                            writer.Write(item.Data.Length);
                            writer.Write(item.Data);
                        }
                        else
                        {
                            writer.Write(0);
                        }
                    }
                }
            }
            catch
            {
                // 保存失败时忽略
            }
        }

        private void LoadHistory()
        {
            try
            {
                if (!File.Exists(dataPath))
                    return;

                using (FileStream fs = new FileStream(dataPath, FileMode.Open))
                using (BinaryReader reader = new BinaryReader(fs))
                {
                    int count = reader.ReadInt32();
                    for (int i = 0; i < count; i++)
                    {
                        ClipboardItem item = new ClipboardItem();
                        item.Time = DateTime.FromBinary(reader.ReadInt64());
                        item.Format = reader.ReadString();
                        item.Text = reader.ReadString();
                        
                        int dataLength = reader.ReadInt32();
                        if (dataLength > 0)
                        {
                            item.Data = reader.ReadBytes(dataLength);
                        }
                        
                        clipboardHistory.Add(item);
                    }
                }

                RefreshListView();
            }
            catch
            {
                // 加载失败时忽略
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveHistory();
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            trayIcon.Visible = false;
            clipboardTimer.Stop();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (trayIcon != null)
                {
                    trayIcon.Dispose();
                }
                if (clipboardTimer != null)
                {
                    clipboardTimer.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
