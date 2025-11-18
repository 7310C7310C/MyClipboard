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
                return Text;
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
        private bool isDarkTheme = true;
        private Color bgColor1, bgColor2, textColor, borderColor;
        private int scrollOffset = 0;
        private Panel listPanel;
        private VScrollBar vScrollBar;
        private const int MIN_ITEM_HEIGHT = 60;
        private const int MAX_TEXT_LINES = 4;
        private List<int> itemHeights = new List<int>();

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

            // 清單面板（自繪）
            listPanel = new Panel();
            listPanel.Dock = DockStyle.Fill;
            listPanel.BackColor = Color.FromArgb(30, 30, 30);
            listPanel.Font = new Font("Microsoft YaHei UI", 12F);
            listPanel.Paint += ListPanel_Paint;
            listPanel.MouseDown += ListPanel_MouseDown;
            listPanel.MouseMove += Form_MouseMove;
            listPanel.MouseUp += Form_MouseUp;
            listPanel.MouseClick += ListPanel_MouseClick;
            listPanel.MouseDoubleClick += ListPanel_MouseDoubleClick;
            listPanel.MouseWheel += ListPanel_MouseWheel;
            contentPanel.Controls.Add(listPanel);

            // 右鍵選單
            listContextMenu = new ContextMenuStrip();
            listContextMenu.Items.Add("複製", null, CopyItem_Click);
            listContextMenu.Items.Add("刪除", null, DeleteItem_Click);
            listContextMenu.Items.Add(new ToolStripSeparator());
            listContextMenu.Items.Add("清空", null, ClearAll_Click);

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

        private void ListPanel_Paint(object sender, PaintEventArgs e)
        {
            if (clipboardHistory.Count == 0)
                return;

            int y = -scrollOffset;
            int panelWidth = listPanel.ClientSize.Width;
            
            for (int i = 0; i < clipboardHistory.Count; i++)
            {
                if (i >= itemHeights.Count)
                    break;
                    
                int itemHeight = itemHeights[i];
                
                // 只繪製可見的項目
                if (y + itemHeight < 0)
                {
                    y += itemHeight;
                    continue;
                }
                if (y > listPanel.ClientSize.Height)
                    break;

                ClipboardItem item = clipboardHistory[i];
                Rectangle itemRect = new Rectangle(0, y, panelWidth, itemHeight);
                
                // 隴行換色
                Color itemBgColor = (i % 2 == 0) ? bgColor1 : bgColor2;
                using (SolidBrush bgBrush = new SolidBrush(itemBgColor))
                {
                    e.Graphics.FillRectangle(bgBrush, itemRect);
                }

                // 給製圖片（右對齊）
                int textWidth = panelWidth - 10;
                if (item.Format == "Image" && item.Data != null)
                {
                    try
                    {
                        int imageIndex = GetImageIndex(i);
                        if (imageIndex >= 0 && imageIndex < imageList.Images.Count)
                        {
                            Image thumbnail = imageList.Images[imageIndex];
                            int imgX = panelWidth - 53;
                            int imgY = y + (itemHeight - 48) / 2;
                            e.Graphics.DrawImage(thumbnail, imgX, imgY, 48, 48);
                            textWidth = panelWidth - 65;
                        }
                    }
                    catch { }
                }

                // 給製文字（左對齊）
                Rectangle textRect = new Rectangle(5, y + 5, textWidth, itemHeight - 10);
                TextRenderer.DrawText(e.Graphics, item.ToString(), listPanel.Font, textRect, 
                    textColor, TextFormatFlags.Left | TextFormatFlags.Top | TextFormatFlags.WordBreak);

                y += itemHeight;
            }
        }   listPanel.Invalidate();
        }

        private void ListPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                dragStartPoint = e.Location;
            }
        }

        private void ListPanel_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int itemIndex = GetItemIndexAtPoint(e.Location);
                if (itemIndex >= 0)
                {
                    listContextMenu.Show(listPanel, e.Location);
                }
                else
                {
                    // 空白处也显示菜单
                    listContextMenu.Show(listPanel, e.Location);
                }
            }
        }

        private void ListPanel_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                int itemIndex = GetItemIndexAtPoint(e.Location);
                if (itemIndex >= 0 && itemIndex < clipboardHistory.Count)
                {
                    PasteItem(clipboardHistory[itemIndex]);
                }
            }
        }

        private int GetItemIndexAtPoint(Point point)
        {
            int y = -scrollOffset;
            for (int i = 0; i < clipboardHistory.Count && i < itemHeights.Count; i++)
            {
                int itemHeight = itemHeights[i];
                if (point.Y >= y && point.Y < y + itemHeight)
                {
                    return i;
                }
                y += itemHeight;
            }
            return -1;
        }

        private int GetImageIndex(int itemIndex)
        {
            int imageIndex = 0;
            for (int i = 0; i < itemIndex && i < clipboardHistory.Count; i++)
            {
                if (clipboardHistory[i].Format == "Image" && clipboardHistory[i].Data != null)
                {
                    imageIndex++;
                }
            }
            return imageIndex;
        }

        private int CalculateItemHeight(ClipboardItem item)
        {
            if (item.Format == "Image")
            {
                return MIN_ITEM_HEIGHT;
            }

            if (string.IsNullOrEmpty(item.Text))
            {
                return MIN_ITEM_HEIGHT;
            }

            // 計算文字需要的行數
            string[] lines = item.Text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            int displayLines = Math.Min(lines.Length, MAX_TEXT_LINES);
            
            // 每行大約20px，再加上上下邊距
            int textHeight = displayLines * 22 + 16;
            return Math.Max(MIN_ITEM_HEIGHT, textHeight);
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

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // 使用明確的狀態切換
                bool wasVisible = this.Visible;
                if (wasVisible)
                {
                    this.Hide();
                }
                else
                {
                    this.Show();
                    this.Activate();
                    this.BringToFront();
                }
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
            // 重新計算所有項目高度
            itemHeights.Clear();
            imageList.Images.Clear();
            
            foreach (var item in clipboardHistory)
            {
                itemHeights.Add(CalculateItemHeight(item));
                
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
                        }
                    }
                    catch { }
                }
            }
            
            listPanel.Invalidate();
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

        private void CopyItem_Click(object sender, EventArgs e)
        {
            // 由於不再使用ListView，需要其他方式選擇項目
            // 暫時簡化，只處理最新項目
            if (clipboardHistory.Count > 0)
            {
                ClipboardItem item = clipboardHistory[0];
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
            if (clipboardHistory.Count > 0)
            {
                clipboardHistory.RemoveAt(0);
                RefreshListView();
                SaveHistory();
            }
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
                scrollOffset = 0;
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
                this.BackColor = Color.FromArgb(63, 63, 70);
                if (listPanel != null)
                {
                    listPanel.BackColor = Color.FromArgb(30, 30, 30);
                    listPanel.ForeColor = Color.White;
                }
            }
            else
            {
                bgColor1 = Color.FromArgb(245, 245, 245);
                bgColor2 = Color.FromArgb(255, 255, 255);
                textColor = Color.Black;
                borderColor = Color.FromArgb(180, 180, 180);
                this.BackColor = Color.FromArgb(180, 180, 180);
                if (listPanel != null)
                {
                    listPanel.BackColor = Color.FromArgb(250, 250, 250);
                    listPanel.ForeColor = Color.Black;
                }
            }
            
            if (listPanel != null)
            {
                listPanel.Invalidate();
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
