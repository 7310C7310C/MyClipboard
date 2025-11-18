using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace MyClipboard
{
    // 剪貼板監聽器
    public class ClipboardMonitor : NativeWindow, IDisposable
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
        
        public event EventHandler ClipboardChanged;
        
        public ClipboardMonitor()
        {
            CreateHandle(new CreateParams());
            AddClipboardFormatListener(this.Handle);
        }
        
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                if (ClipboardChanged != null)
                {
                    ClipboardChanged(this, EventArgs.Empty);
                }
            }
            base.WndProc(ref m);
        }
        
        public void Dispose()
        {
            RemoveClipboardFormatListener(this.Handle);
            DestroyHandle();
        }
    }

    public class ClipboardItem
    {
        public DateTime Time { get; set; }
        public string Text { get; set; }
        public string Format { get; set; }
        public byte[] Data { get; set; }
        
        public string GetPreview(int maxLength = 200)
        {
            if (!string.IsNullOrEmpty(Text))
            {
                string preview = Text.Length > maxLength ? Text.Substring(0, maxLength) + "..." : Text;
                return preview;
            }
            return Format;
        }
    }

    // 卡片式剪貼板項目控件
    public class ClipboardCard : Panel
    {
        private ClipboardItem item;
        private Label textLabel;
        private PictureBox imageBox;
        private Label timeLabel;
        private bool isHovered = false;
        private Color normalColor, hoverColor, textColor;
        
        public ClipboardItem Item { get { return item; } }
        public event EventHandler OnDoubleClickItem;
        public event EventHandler OnCopyRequested;
        public event EventHandler OnDeleteRequested;
        
        public ClipboardCard(ClipboardItem clipItem, bool isDark)
        {
            item = clipItem;
            InitializeCard(isDark);
        }
        
        private void InitializeCard(bool isDark)
        {
            // 設置卡片樣式
            this.Width = 410;
            this.Height = item.Format == "Image" ? 140 : 90;
            this.Margin = new Padding(5, 5, 5, 5);
            this.Padding = new Padding(12);
            this.Cursor = Cursors.Hand;
            
            if (isDark)
            {
                normalColor = Color.FromArgb(45, 45, 48);
                hoverColor = Color.FromArgb(60, 60, 63);
                textColor = Color.White;
            }
            else
            {
                normalColor = Color.FromArgb(250, 250, 250);
                hoverColor = Color.FromArgb(240, 240, 240);
                textColor = Color.Black;
            }
            
            this.BackColor = normalColor;
            
            // 時間標籤
            timeLabel = new Label();
            timeLabel.Text = item.Time.ToString("HH:mm:ss");
            timeLabel.ForeColor = isDark ? Color.FromArgb(180, 180, 180) : Color.FromArgb(120, 120, 120);
            timeLabel.Font = new Font("Microsoft YaHei UI", 8F);
            timeLabel.AutoSize = true;
            timeLabel.Location = new Point(12, 12);
            this.Controls.Add(timeLabel);
            
            // 文字標籤
            textLabel = new Label();
            textLabel.Text = item.GetPreview(150);
            textLabel.ForeColor = textColor;
            textLabel.Font = new Font("Microsoft YaHei UI", 10F);
            textLabel.AutoSize = false;
            textLabel.Size = new Size(item.Format == "Image" ? 300 : 380, item.Format == "Image" ? 50 : 40);
            textLabel.Location = new Point(12, 35);
            this.Controls.Add(textLabel);
            
            // 圖片預覽
            if (item.Format == "Image" && item.Data != null)
            {
                try
                {
                    using (MemoryStream ms = new MemoryStream(item.Data))
                    {
                        Image img = Image.FromStream(ms);
                        Image thumbnail = CreateThumbnail(img, 80, 80);
                        
                        imageBox = new PictureBox();
                        imageBox.Image = thumbnail;
                        imageBox.SizeMode = PictureBoxSizeMode.CenterImage;
                        imageBox.Size = new Size(80, 80);
                        imageBox.Location = new Point(320, 35);
                        imageBox.BackColor = Color.Transparent;
                        this.Controls.Add(imageBox);
                    }
                }
                catch { }
            }
            
            // 事件處理
            this.MouseEnter += (s, e) => { isHovered = true; this.BackColor = hoverColor; };
            this.MouseLeave += (s, e) => { isHovered = false; this.BackColor = normalColor; };
            this.DoubleClick += (s, e) => { if (OnDoubleClickItem != null) OnDoubleClickItem.Invoke(this, EventArgs.Empty); };
            
            textLabel.MouseEnter += (s, e) => { isHovered = true; this.BackColor = hoverColor; };
            textLabel.MouseLeave += (s, e) => { isHovered = false; this.BackColor = normalColor; };
            textLabel.DoubleClick += (s, e) => { if (OnDoubleClickItem != null) OnDoubleClickItem.Invoke(this, EventArgs.Empty); };
            
            timeLabel.MouseEnter += (s, e) => { isHovered = true; this.BackColor = hoverColor; };
            timeLabel.MouseLeave += (s, e) => { isHovered = false; this.BackColor = normalColor; };
            timeLabel.DoubleClick += (s, e) => { if (OnDoubleClickItem != null) OnDoubleClickItem.Invoke(this, EventArgs.Empty); };
            
            if (imageBox != null)
            {
                imageBox.MouseEnter += (s, e) => { isHovered = true; this.BackColor = hoverColor; };
                imageBox.MouseLeave += (s, e) => { isHovered = false; this.BackColor = normalColor; };
                imageBox.DoubleClick += (s, e) => { if (OnDoubleClickItem != null) OnDoubleClickItem.Invoke(this, EventArgs.Empty); };
            }
            
            // 右鍵菜單
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.BackColor = isDark ? Color.FromArgb(45, 45, 48) : Color.White;
            menu.ForeColor = textColor;
            
            ToolStripMenuItem copyItem = new ToolStripMenuItem("複製");
            copyItem.Click += (s, e) => { if (OnCopyRequested != null) OnCopyRequested.Invoke(this, EventArgs.Empty); };
            menu.Items.Add(copyItem);
            
            ToolStripMenuItem deleteItem = new ToolStripMenuItem("刪除");
            deleteItem.Click += (s, e) => { if (OnDeleteRequested != null) OnDeleteRequested.Invoke(this, EventArgs.Empty); };
            menu.Items.Add(deleteItem);
            
            this.ContextMenuStrip = menu;
        }
        
        private Image CreateThumbnail(Image image, int width, int height)
        {
            double ratioX = (double)width / image.Width;
            double ratioY = (double)height / image.Height;
            double ratio = Math.Min(ratioX, ratioY);

            int newWidth = (int)(image.Width * ratio);
            int newHeight = (int)(image.Height * ratio);

            Bitmap newImage = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.Clear(Color.Transparent);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                int x = (width - newWidth) / 2;
                int y = (height - newHeight) / 2;
                g.DrawImage(image, x, y, newWidth, newHeight);
            }
            return newImage;
        }
        
        public void UpdateTheme(bool isDark)
        {
            if (isDark)
            {
                normalColor = Color.FromArgb(45, 45, 48);
                hoverColor = Color.FromArgb(60, 60, 63);
                textColor = Color.White;
                timeLabel.ForeColor = Color.FromArgb(180, 180, 180);
            }
            else
            {
                normalColor = Color.FromArgb(250, 250, 250);
                hoverColor = Color.FromArgb(240, 240, 240);
                textColor = Color.Black;
                timeLabel.ForeColor = Color.FromArgb(120, 120, 120);
            }
            
            this.BackColor = isHovered ? hoverColor : normalColor;
            textLabel.ForeColor = textColor;
        }
    }

    public class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private Panel containerPanel;
        private FlowLayoutPanel flowPanel;
        private List<ClipboardItem> clipboardHistory = new List<ClipboardItem>();
        private string dataPath = @"C:\ProgramData\MyClipboard\history.dat";
        private string settingsPath = @"C:\ProgramData\MyClipboard\settings.dat";
        private ClipboardMonitor clipboardMonitor;
        private string lastClipboardText = null;
        private bool isDragging = false;
        private Point dragStartPoint;
        private bool isDarkTheme = true;
        private Color bgColor, titleBarColor, textColor, borderColor;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

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
            LoadSettings();
            ApplyTheme();
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, VK_X);
            
            // 啟動剪貼板監控
            clipboardMonitor = new ClipboardMonitor();
            clipboardMonitor.ClipboardChanged += ClipboardMonitor_ClipboardChanged;
            
            // 初始化時捕獲當前剪貼板內容
            CaptureCurrentClipboard();
        }

        private void InitializeComponents()
        {
            // 主窗體設置
            this.Text = "MyClipboard";
            this.Size = new Size(450, 650);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.Padding = new Padding(2);

            // 容器面板
            containerPanel = new Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = Color.FromArgb(30, 30, 30);
            this.Controls.Add(containerPanel);
            
            // 標題欄
            Panel titleBar = new Panel();
            titleBar.Dock = DockStyle.Top;
            titleBar.Height = 45;
            titleBar.BackColor = Color.FromArgb(37, 37, 38);
            titleBar.MouseDown += TitleBar_MouseDown;
            titleBar.MouseMove += TitleBar_MouseMove;
            titleBar.MouseUp += TitleBar_MouseUp;
            
            Label titleLabel = new Label();
            titleLabel.Text = "MyClipboard";
            titleLabel.ForeColor = Color.White;
            titleLabel.Font = new Font("Microsoft YaHei UI", 13F, FontStyle.Bold);
            titleLabel.AutoSize = false;
            titleLabel.Size = new Size(200, 45);
            titleLabel.Location = new Point(15, 0);
            titleLabel.TextAlign = ContentAlignment.MiddleLeft;
            titleLabel.MouseDown += TitleBar_MouseDown;
            titleLabel.MouseMove += TitleBar_MouseMove;
            titleLabel.MouseUp += TitleBar_MouseUp;
            titleBar.Controls.Add(titleLabel);
            
            // 清空按鈕
            Button clearBtn = new Button();
            clearBtn.Text = "清空";
            clearBtn.ForeColor = Color.White;
            clearBtn.BackColor = Color.FromArgb(60, 60, 60);
            clearBtn.FlatStyle = FlatStyle.Flat;
            clearBtn.FlatAppearance.BorderSize = 0;
            clearBtn.Size = new Size(60, 28);
            clearBtn.Location = new Point(365, 9);
            clearBtn.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            clearBtn.Font = new Font("Microsoft YaHei UI", 9F);
            clearBtn.Cursor = Cursors.Hand;
            clearBtn.Click += ClearAll_Click;
            clearBtn.MouseEnter += (s, e) => clearBtn.BackColor = Color.FromArgb(80, 80, 80);
            clearBtn.MouseLeave += (s, e) => clearBtn.BackColor = Color.FromArgb(60, 60, 60);
            titleBar.Controls.Add(clearBtn);
            
            containerPanel.Controls.Add(titleBar);

            // FlowLayoutPanel 用於卡片式布局
            flowPanel = new FlowLayoutPanel();
            flowPanel.Dock = DockStyle.Fill;
            flowPanel.AutoScroll = true;
            flowPanel.BackColor = Color.FromArgb(30, 30, 30);
            flowPanel.FlowDirection = FlowDirection.TopDown;
            flowPanel.WrapContents = false;
            flowPanel.Padding = new Padding(10);
            containerPanel.Controls.Add(flowPanel);

            // 托盤圖示
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Visible = true;
            trayIcon.Text = "MyClipboard";
            trayIcon.MouseClick += TrayIcon_MouseClick;
            
            // 托盤右鍵選單
            ContextMenuStrip trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("切換主題", null, (s, ev) => {
                isDarkTheme = !isDarkTheme;
                ApplyTheme();
                SaveSettings();
            });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("退出", null, (s, ev) => {
                Application.Exit();
            });
            trayIcon.ContextMenuStrip = trayMenu;

            // 嘗試加載自定義圖標
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
                    this.Icon = Icon.ExtractAssociatedIcon(exePath);
                    trayIcon.Icon = Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch { }

            // 窗體事件
            this.Load += MainForm_Load;
            this.Deactivate += MainForm_Deactivate;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void MainForm_Deactivate(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                dragStartPoint = e.Location;
            }
        }

        private void TitleBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentScreenPos = Control.MousePosition;
                this.Location = new Point(currentScreenPos.X - dragStartPoint.X, currentScreenPos.Y - dragStartPoint.Y);
            }
        }

        private void TitleBar_MouseUp(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                isDragging = false;
                SaveSettings();
            }
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.BeginInvoke(new Action(() => {
                    if (this.Visible)
                    {
                        this.Hide();
                    }
                    else
                    {
                        this.Show();
                        this.Activate();
                        this.Focus();
                    }
                }));
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                if (this.Visible)
                {
                    this.Hide();
                }
                else
                {
                    this.Show();
                    this.Activate();
                }
            }
            base.WndProc(ref m);
        }

        private void ClipboardMonitor_ClipboardChanged(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(50);
            CaptureCurrentClipboard();
        }

        private void CaptureCurrentClipboard()
        {
            try
            {
                if (!Clipboard.ContainsData(DataFormats.Text) && 
                    !Clipboard.ContainsData(DataFormats.UnicodeText) &&
                    !Clipboard.ContainsData(DataFormats.Rtf) &&
                    !Clipboard.ContainsData(DataFormats.Html) &&
                    !Clipboard.ContainsImage())
                {
                    return;
                }

                ClipboardItem item = new ClipboardItem();
                item.Time = DateTime.Now;
                bool hasData = false;

                if (Clipboard.ContainsData(DataFormats.UnicodeText))
                {
                    string text = Clipboard.GetText(TextDataFormat.UnicodeText);
                    if (text == lastClipboardText)
                        return;
                    
                    lastClipboardText = text;
                    item.Text = text;
                    item.Format = "Text";
                    hasData = true;
                }
                else if (Clipboard.ContainsData(DataFormats.Text))
                {
                    string text = Clipboard.GetText(TextDataFormat.Text);
                    if (text == lastClipboardText)
                        return;
                    
                    lastClipboardText = text;
                    item.Text = text;
                    item.Format = "Text";
                    hasData = true;
                }
                else if (Clipboard.ContainsData(DataFormats.Rtf))
                {
                    string rtf = Clipboard.GetData(DataFormats.Rtf) as string;
                    item.Data = Encoding.UTF8.GetBytes(rtf);
                    item.Format = "RTF";
                    item.Text = "富文本內容";
                    lastClipboardText = null;
                    hasData = true;
                }
                else if (Clipboard.ContainsData(DataFormats.Html))
                {
                    string html = Clipboard.GetData(DataFormats.Html) as string;
                    item.Data = Encoding.UTF8.GetBytes(html);
                    item.Format = "HTML";
                    item.Text = "HTML內容";
                    lastClipboardText = null;
                    hasData = true;
                }
                else if (Clipboard.ContainsImage())
                {
                    Image img = Clipboard.GetImage();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        item.Data = ms.ToArray();
                    }
                    item.Format = "Image";
                    item.Text = string.Format("圖片 ({0}x{1})", img.Width, img.Height);
                    lastClipboardText = null;
                    hasData = true;
                }

                if (!hasData)
                    return;

                if (clipboardHistory.Count > 0)
                {
                    var last = clipboardHistory[0];
                    if (last.Text == item.Text && last.Format == item.Format)
                        return;
                }

                clipboardHistory.Insert(0, item);
                
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => {
                        RefreshCards();
                        SaveHistory();
                    }));
                }
                else
                {
                    RefreshCards();
                    SaveHistory();
                }
            }
            catch { }
        }

        private void RefreshCards()
        {
            flowPanel.SuspendLayout();
            flowPanel.Controls.Clear();
            
            foreach (var item in clipboardHistory)
            {
                ClipboardCard card = new ClipboardCard(item, isDarkTheme);
                card.OnDoubleClickItem += (s, e) => PasteItem(item);
                card.OnCopyRequested += (s, e) => CopyItem(item);
                card.OnDeleteRequested += (s, e) => DeleteItem(item);
                flowPanel.Controls.Add(card);
            }
            
            flowPanel.ResumeLayout();
        }

        private void PasteItem(ClipboardItem item)
        {
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

                this.Hide();
                Thread.Sleep(100);

                keybd_event(0x11, 0, 0, UIntPtr.Zero);
                keybd_event(0x56, 0, 0, UIntPtr.Zero);
                keybd_event(0x56, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                keybd_event(0x11, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            catch (Exception ex)
            {
                MessageBox.Show("貼上失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopyItem(ClipboardItem item)
        {
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
                MessageBox.Show("複製失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteItem(ClipboardItem item)
        {
            clipboardHistory.Remove(item);
            RefreshCards();
            SaveHistory();
        }

        private void ClearAll_Click(object sender, EventArgs e)
        {
            if (clipboardHistory.Count == 0)
                return;
                
            DialogResult result = MessageBox.Show(
                "確定要清空所有記錄嗎？",
                "確認清空",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                clipboardHistory.Clear();
                RefreshCards();
                SaveHistory();
            }
        }

        private void ApplyTheme()
        {
            if (isDarkTheme)
            {
                bgColor = Color.FromArgb(30, 30, 30);
                titleBarColor = Color.FromArgb(37, 37, 38);
                textColor = Color.White;
                borderColor = Color.FromArgb(63, 63, 70);
                this.BackColor = borderColor;
                containerPanel.BackColor = bgColor;
                flowPanel.BackColor = bgColor;
            }
            else
            {
                bgColor = Color.FromArgb(245, 245, 245);
                titleBarColor = Color.FromArgb(220, 220, 220);
                textColor = Color.Black;
                borderColor = Color.FromArgb(180, 180, 180);
                this.BackColor = borderColor;
                containerPanel.BackColor = bgColor;
                flowPanel.BackColor = bgColor;
            }
            
            RefreshCards();
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
            catch { }
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

                RefreshCards();
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                string dir = Path.GetDirectoryName(settingsPath);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                using (FileStream fs = new FileStream(settingsPath, FileMode.Create))
                using (BinaryWriter writer = new BinaryWriter(fs))
                {
                    writer.Write(this.Location.X);
                    writer.Write(this.Location.Y);
                    writer.Write(isDarkTheme);
                }
            }
            catch { }
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(settingsPath))
                {
                    using (FileStream fs = new FileStream(settingsPath, FileMode.Open))
                    using (BinaryReader reader = new BinaryReader(fs))
                    {
                        int x = reader.ReadInt32();
                        int y = reader.ReadInt32();
                        this.Location = new Point(x, y);
                        
                        if (fs.Position < fs.Length)
                        {
                            isDarkTheme = reader.ReadBoolean();
                        }
                        return;
                    }
                }
            }
            catch { }
            
            this.StartPosition = FormStartPosition.CenterScreen;
            isDarkTheme = true;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveHistory();
            SaveSettings();
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            trayIcon.Visible = false;
            
            if (clipboardMonitor != null)
            {
                clipboardMonitor.Dispose();
            }
            
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
                if (clipboardMonitor != null)
                {
                    clipboardMonitor.Dispose();
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
