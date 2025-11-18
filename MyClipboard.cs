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
        
        public override string ToString()
        {
            if (!string.IsNullOrEmpty(Text))
            {
                string preview = Text.Length > 80 ? Text.Substring(0, 80) + "..." : Text;
                preview = preview.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
                return preview;
            }
            return Format;
        }
    }

    public class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private ListView listView;
        private ContextMenuStrip listContextMenu;
        private List<ClipboardItem> clipboardHistory = new List<ClipboardItem>();
        private string dataPath = @"C:\ProgramData\MyClipboard\history.dat";
        private string settingsPath = @"C:\ProgramData\MyClipboard\settings.dat";
        private ClipboardMonitor clipboardMonitor;
        private string lastClipboardText = null;
        private bool isDragging = false;
        private Point dragStartPoint;
        private ImageList imageList;
        private ImageList largeImageList;
        private bool isDarkTheme = true;
        private Color bgColor1, bgColor2, textColor, borderColor;

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
            this.Size = new Size(400, 600);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            
            // 窗體拖動事件
            this.MouseDown += Form_MouseDown;
            this.MouseMove += Form_MouseMove;
            this.MouseUp += Form_MouseUp;
            
            // 添加邊框效果
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.Padding = new Padding(1);

            // 內容面板
            Panel contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.BackColor = Color.FromArgb(30, 30, 30);
            contentPanel.MouseDown += Form_MouseDown;
            contentPanel.MouseMove += Form_MouseMove;
            contentPanel.MouseUp += Form_MouseUp;
            this.Controls.Add(contentPanel);

            // ImageList for thumbnails
            imageList = new ImageList();
            imageList.ImageSize = new Size(48, 48);
            imageList.ColorDepth = ColorDepth.Depth32Bit;
            
            largeImageList = new ImageList();
            largeImageList.ImageSize = new Size(48, 48);
            largeImageList.ColorDepth = ColorDepth.Depth32Bit;

            // ListView設置 - 使用OwnerDraw自繪
            listView = new ListView();
            listView.Dock = DockStyle.Fill;
            listView.View = View.Details;
            listView.FullRowSelect = true;
            listView.BackColor = Color.FromArgb(30, 30, 30);
            listView.ForeColor = Color.White;
            listView.BorderStyle = BorderStyle.None;
            listView.Font = new Font("Microsoft YaHei UI", 11F);
            listView.HeaderStyle = ColumnHeaderStyle.None;
            listView.Columns.Add("", 380);
            listView.OwnerDraw = true;
            listView.DrawItem += ListView_DrawItem;
            listView.Scrollable = false;
            listView.DoubleClick += ListView_DoubleClick;
            listView.MouseClick += ListView_MouseClick;
            listView.MouseDown += Form_MouseDown;
            listView.MouseMove += Form_MouseMove;
            listView.MouseUp += Form_MouseUp;
            contentPanel.Controls.Add(listView);

            // 右鍵選單
            listContextMenu = new ContextMenuStrip();
            listContextMenu.Items.Add("複製", null, CopyItem_Click);
            listContextMenu.Items.Add("刪除", null, DeleteItem_Click);
            listView.ContextMenuStrip = listContextMenu;

            // 托盤圖示
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Visible = true;
            trayIcon.Text = "MyClipboard";
            
            // 托盤右鍵選單
            ContextMenuStrip trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("淺色", null, (s, ev) => {
                isDarkTheme = false;
                ApplyTheme();
                SaveSettings();
            });
            trayMenu.Items.Add("深色", null, (s, ev) => {
                isDarkTheme = true;
                ApplyTheme();
                SaveSettings();
            });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("退出", null, (s, ev) => {
                Application.Exit();
            });
            trayIcon.ContextMenuStrip = trayMenu;

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
            // 啟動時隱藏窗口
            this.Hide();
        }

        private void MainForm_Deactivate(object sender, EventArgs e)
        {
            // 失去焦點時隱藏
            this.Hide();
        }

        private void ListView_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // 隴行換色
            Color itemBgColor = (e.ItemIndex % 2 == 0) ? bgColor1 : bgColor2;
            using (SolidBrush bgBrush = new SolidBrush(itemBgColor))
            {
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            }

            ClipboardItem item = e.Item.Tag as ClipboardItem;
            if (item == null) return;

            // 給製文字（左對齊）
            Rectangle textRect = new Rectangle(e.Bounds.X + 5, e.Bounds.Y, e.Bounds.Width - 60, e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, item.ToString(), listView.Font, textRect, 
                textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            // 給製圖片（右對齊）
            if (item.Format == "Image" && item.Data != null && e.Item.ImageIndex >= 0)
            {
                try
                {
                    Image thumbnail = imageList.Images[e.Item.ImageIndex];
                    int imgX = e.Bounds.Right - 53;
                    int imgY = e.Bounds.Y + (e.Bounds.Height - 48) / 2;
                    e.Graphics.DrawImage(thumbnail, imgX, imgY, 48, 48);
                }
                catch { }
            }

            // 選中時的邊框
            if (e.Item.Selected)
            {
                using (Pen pen = new Pen(Color.FromArgb(0, 120, 215), 2))
                {
                    Rectangle selRect = new Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
                    e.Graphics.DrawRectangle(pen, selRect);
                }
            }
        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                dragStartPoint = e.Location;
            }
        }

        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentScreenPos = Control.MousePosition;
                this.Location = new Point(currentScreenPos.X - dragStartPoint.X, currentScreenPos.Y - dragStartPoint.Y);
            }
        }

        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                isDragging = false;
                SaveWindowPosition();
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
            // 延遲一下確保剪貼板數據已經準備好
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

                // 優先處理文本
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

                // 避免重複
                if (clipboardHistory.Count > 0)
                {
                    var last = clipboardHistory[0];
                    if (last.Text == item.Text && last.Format == item.Format)
                        return;
                }

                clipboardHistory.Insert(0, item);
                
                // 在UI線程中更新界面
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => {
                        RefreshListView();
                        SaveHistory();
                    }));
                }
                else
                {
                    RefreshListView();
                    SaveHistory();
                }
            }
            catch
            {
                // 剪貼板訪問失敗時忽略
            }
        }

        private void RefreshListView()
        {
            listView.BeginUpdate();
            listView.Items.Clear();
            imageList.Images.Clear();
            
            int imageIndex = 0;
            foreach (var item in clipboardHistory)
            {
                ListViewItem lvi = new ListViewItem(item.ToString());
                lvi.Tag = item;
                
                // 為圖片添加縮略圖
                if (item.Format == "Image" && item.Data != null)
                {
                    try
                    {
                        using (MemoryStream ms = new MemoryStream(item.Data))
                        {
                            Image img = Image.FromStream(ms);
                            Image thumbnail = CreateThumbnail(img, 48, 48);
                            imageList.Images.Add(thumbnail);
                            lvi.ImageIndex = imageIndex;
                            imageIndex++;
                        }
                    }
                    catch { }
                }
                
                listView.Items.Add(lvi);
            }
            
            // 設置固定高度，無間距
            if (listView.Items.Count > 0)
            {
                listView.Items[0].EnsureVisible();
            }
            
            listView.EndUpdate();
            listView.Invalidate();
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
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                int x = (width - newWidth) / 2;
                int y = (height - newHeight) / 2;
                g.DrawImage(image, x, y, newWidth, newHeight);
            }
            return newImage;
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
                MessageBox.Show("貼上失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("複製失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void SaveWindowPosition()
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
        
        private void SaveSettings()
        {
            SaveWindowPosition();
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
            
            // 如果沒有保存的設置，默認居中和深色主題
            this.StartPosition = FormStartPosition.CenterScreen;
            isDarkTheme = true;
        }

        private void ApplyTheme()
        {
            if (isDarkTheme)
            {
                bgColor1 = Color.FromArgb(37, 37, 38);
                bgColor2 = Color.FromArgb(45, 45, 48);
                textColor = Color.White;
                borderColor = Color.FromArgb(63, 63, 70);
                this.BackColor = Color.FromArgb(45, 45, 48);
                listView.BackColor = Color.FromArgb(30, 30, 30);
                listView.ForeColor = Color.White;
            }
            else
            {
                bgColor1 = Color.FromArgb(245, 245, 245);
                bgColor2 = Color.FromArgb(255, 255, 255);
                textColor = Color.Black;
                borderColor = Color.FromArgb(200, 200, 200);
                this.BackColor = Color.FromArgb(230, 230, 230);
                listView.BackColor = Color.FromArgb(250, 250, 250);
                listView.ForeColor = Color.Black;
            }
            
            listView.Invalidate();
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
                if (imageList != null)
                {
                    imageList.Dispose();
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
