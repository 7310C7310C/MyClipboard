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
        public bool IsFavorite { get; set; }
        
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
        private Panel scrollBarPanel;
        private Button minimizeButton;
        private Button favoritesButton;
        private Button searchClearButton;
        private Panel searchPanel;
        private TextBox searchBox;
        private Point lastContextMenuPosition;
        private const int ITEM_HEIGHT = 60;
        private const int SEARCH_BOX_HEIGHT = 35;
        private bool firstRun = true;
        private bool scrollBarDragging = false;
        private int scrollBarDragStart = 0;
        private int scrollOffsetDragStart = 0;
        private bool showingFavorites = false;
        private string searchFilter = "";
        private Color favoriteBgColor1Light, favoriteBgColor2Light;
        private Color favoriteBgColor1Dark, favoriteBgColor2Dark;
        private int selectedIndex = -1;
        private Form imagePreviewForm = null;

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
            this.KeyPreview = true;
            
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
            contentPanel.KeyDown += MainForm_KeyDown;
            this.Controls.Add(contentPanel);

            // ImageList for thumbnails
            imageList = new ImageList();
            imageList.ImageSize = new Size(54, 54);
            imageList.ColorDepth = ColorDepth.Depth32Bit;

            // 清單面板（自繪）
            listPanel = new Panel();
            listPanel.Dock = DockStyle.None;
            listPanel.BackColor = Color.FromArgb(30, 30, 30);
            listPanel.Font = new Font("Consolas", 10F);
            listPanel.TabStop = true;
            // 啟用雙緩衝減少閃爍
            typeof(Panel).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, listPanel, new object[] { true });
            listPanel.Paint += ListPanel_Paint;
            listPanel.MouseDown += ListPanel_MouseDown;
            listPanel.MouseMove += Form_MouseMove;
            listPanel.MouseUp += Form_MouseUp;
            listPanel.MouseClick += ListPanel_MouseClick;
            listPanel.MouseDoubleClick += ListPanel_MouseDoubleClick;
            listPanel.MouseWheel += ListPanel_MouseWheel;
            listPanel.KeyDown += MainForm_KeyDown;
            contentPanel.Controls.Add(listPanel);
            
            // 调整listPanel大小以避开搜索框
            Action updateListPanelSize = () => {
                listPanel.Location = new Point(0, 0);
                listPanel.Size = new Size(contentPanel.ClientSize.Width, contentPanel.ClientSize.Height - SEARCH_BOX_HEIGHT);
                // 确保滚动条不超过listPanel的高度
                if (scrollBarPanel != null && scrollBarPanel.Height + scrollBarPanel.Top > listPanel.Height)
                {
                    int maxHeight = listPanel.Height - scrollBarPanel.Top;
                    if (maxHeight > 0)
                    {
                        scrollBarPanel.Height = Math.Min(scrollBarPanel.Height, maxHeight);
                    }
                }
            };
            contentPanel.SizeChanged += (s, ev) => updateListPanelSize();
            updateListPanelSize();
            
            // Material Design 滚动条
            scrollBarPanel = new Panel();
            scrollBarPanel.Width = 8;
            scrollBarPanel.BackColor = Color.FromArgb(120, 120, 120);
            scrollBarPanel.Visible = true;
            scrollBarPanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            scrollBarPanel.Cursor = Cursors.Hand;
            scrollBarPanel.MouseDown += ScrollBar_MouseDown;
            scrollBarPanel.MouseMove += ScrollBar_MouseMove;
            scrollBarPanel.MouseUp += ScrollBar_MouseUp;
            scrollBarPanel.KeyDown += MainForm_KeyDown;
            contentPanel.Controls.Add(scrollBarPanel);
            scrollBarPanel.BringToFront();
            
            // 滚动条不再需要计时器隐藏
            
            // 收藏按钮
            favoritesButton = new Button();
            favoritesButton.Text = "我的收藏";
            favoritesButton.Size = new Size(80, 30);
            favoritesButton.Location = new Point(this.ClientSize.Width - 125, 5);
            favoritesButton.FlatStyle = FlatStyle.Flat;
            favoritesButton.FlatAppearance.BorderSize = 0;
            favoritesButton.BackColor = Color.FromArgb(60, 60, 60);
            favoritesButton.ForeColor = Color.White;
            favoritesButton.Font = new Font("Arial", 9F);
            favoritesButton.Cursor = Cursors.Hand;
            favoritesButton.TabStop = false;
            favoritesButton.Click += FavoritesButton_Click;
            favoritesButton.GotFocus += (s, ev) => {
                // 按钮获得焦点时立即转移到 listPanel
                listPanel.Focus();
            };
            favoritesButton.KeyDown += (s, ev) => {
                MainForm_KeyDown(this, ev);
            };
            contentPanel.Controls.Add(favoritesButton);
            favoritesButton.BringToFront();
            
            // 最小化按钮
            minimizeButton = new Button();
            minimizeButton.Text = "—";
            minimizeButton.Size = new Size(30, 30);
            minimizeButton.Location = new Point(this.ClientSize.Width - 35, 5);
            minimizeButton.FlatStyle = FlatStyle.Flat;
            minimizeButton.FlatAppearance.BorderSize = 0;
            minimizeButton.BackColor = Color.FromArgb(60, 60, 60);
            minimizeButton.ForeColor = Color.White;
            minimizeButton.Font = new Font("Arial", 12F, FontStyle.Bold);
            minimizeButton.Cursor = Cursors.Hand;
            minimizeButton.TabStop = false;
            minimizeButton.Click += MinimizeButton_Click;
            minimizeButton.GotFocus += (s, ev) => {
                // 按钮获得焦点时立即转移到 listPanel
                listPanel.Focus();
            };
            minimizeButton.KeyDown += (s, ev) => {
                MainForm_KeyDown(this, ev);
            };
            contentPanel.Controls.Add(minimizeButton);
            minimizeButton.BringToFront();
            
            // 搜索框（底部）
            searchPanel = new Panel();
            searchPanel.Dock = DockStyle.Bottom;
            searchPanel.Height = SEARCH_BOX_HEIGHT;
            searchPanel.BackColor = Color.FromArgb(0, 90, 158);
            searchPanel.KeyDown += MainForm_KeyDown;
            contentPanel.Controls.Add(searchPanel);
            
            searchBox = new TextBox();
            searchBox.Dock = DockStyle.Fill;
            searchBox.Font = new Font("Consolas", 16F);
            searchBox.ForeColor = Color.FromArgb(180, 210, 240);
            searchBox.Text = "搜索……";
            searchBox.BackColor = Color.FromArgb(0, 90, 158);
            searchBox.BorderStyle = BorderStyle.None;
            searchBox.TextAlign = HorizontalAlignment.Left;
            searchBox.Multiline = true;
            searchBox.ReadOnly = true;
            searchBox.Padding = new Padding(10, 9, 0, 0);
            searchBox.TextChanged += SearchBox_TextChanged;
            searchBox.KeyDown += SearchBox_KeyDown;
            searchBox.Enter += (s, ev) => {
                if (searchBox.Text == "搜索……")
                {
                    searchBox.ReadOnly = false;
                    searchBox.Text = "";
                    searchBox.ForeColor = isDarkTheme ? Color.White : Color.Black;
                }
                selectedIndex = -1;
                listPanel.Invalidate();
            };
            searchBox.Leave += (s, ev) => {
                if (string.IsNullOrWhiteSpace(searchBox.Text))
                {
                    searchBox.Text = "搜索……";
                    searchBox.ForeColor = isDarkTheme ? Color.FromArgb(180, 210, 240) : Color.FromArgb(60, 60, 60);
                    searchBox.ReadOnly = true;
                    searchClearButton.Visible = false;
                }
            };
            searchPanel.Controls.Add(searchBox);
            
            // 搜索框清除按钮
            searchClearButton = new Button();
            searchClearButton.Text = "✕";
            searchClearButton.Width = 40;
            searchClearButton.Dock = DockStyle.Right;
            searchClearButton.FlatStyle = FlatStyle.Flat;
            searchClearButton.FlatAppearance.BorderSize = 0;
            searchClearButton.BackColor = Color.FromArgb(0, 90, 158);
            searchClearButton.ForeColor = Color.White;
            searchClearButton.Font = new Font("Arial", 11F);
            searchClearButton.TextAlign = ContentAlignment.MiddleCenter;
            searchClearButton.Cursor = Cursors.Hand;
            searchClearButton.TabStop = false;
            searchClearButton.Visible = false;
            searchClearButton.Click += (s, ev) => {
                searchBox.Text = "";
                searchBox.Focus();
            };
            searchPanel.Controls.Add(searchClearButton);
            searchClearButton.BringToFront();

            // 右鍵選單
            listContextMenu = new ContextMenuStrip();
            listContextMenu.Items.Add("收藏", null, ToggleFavorite_Click);
            listContextMenu.Items.Add(new ToolStripSeparator());
            listContextMenu.Items.Add("編輯", null, EditItem_Click);
            listContextMenu.Items.Add("複製", null, CopyItem_Click);
            listContextMenu.Items.Add("刪除", null, DeleteItem_Click);
            listContextMenu.Items.Add(new ToolStripSeparator());
            listContextMenu.Items.Add("清空", null, ClearAll_Click);
            listContextMenu.Items.Add(new ToolStripSeparator());

            // 在右键菜单中加入“切換主題”及二级菜单（放在“幫助”之前）
            ToolStripMenuItem ctxThemeMenu = new ToolStripMenuItem("切換主題");
            ctxThemeMenu.Name = "ctxThemeMenu";
            ToolStripMenuItem ctxLightTheme = new ToolStripMenuItem("淺色");
            ctxLightTheme.Name = "ctxLightTheme";
            ctxLightTheme.Click += (s, ev) => {
                isDarkTheme = false;
                ApplyTheme();
                SaveSettings();
                listPanel.Invalidate();
            };
            ToolStripMenuItem ctxDarkTheme = new ToolStripMenuItem("深色");
            ctxDarkTheme.Name = "ctxDarkTheme";
            ctxDarkTheme.Click += (s, ev) => {
                isDarkTheme = true;
                ApplyTheme();
                SaveSettings();
                listPanel.Invalidate();
            };
            ctxThemeMenu.DropDownItems.Add(ctxLightTheme);
            ctxThemeMenu.DropDownItems.Add(ctxDarkTheme);

            listContextMenu.Items.Add(ctxThemeMenu);

            listContextMenu.Items.Add("幫助", null, Help_Click);
            // 与托盘 About 相同的“關於”项
            listContextMenu.Items.Add("關於", null, About_Click);
            listContextMenu.Opening += ListContextMenu_Opening;

            // 托盤圖示
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Visible = true;
            trayIcon.Text = "MyClipboard";
            
            // 托盤右鍵選單
            ContextMenuStrip trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("顯示 / 隱藏（Ctrl + Alt + X）", null, (s, ev) => {
                ToggleWindow();
            });
            
            // 主題切換二級菜單
            ToolStripMenuItem themeMenuItem = new ToolStripMenuItem("切換主題");
            ToolStripMenuItem lightThemeItem = new ToolStripMenuItem("淺色");
            lightThemeItem.Click += (s, ev) => {
                isDarkTheme = false;
                ApplyTheme();
                SaveSettings();
            };
            ToolStripMenuItem darkThemeItem = new ToolStripMenuItem("深色");
            darkThemeItem.Click += (s, ev) => {
                isDarkTheme = true;
                ApplyTheme();
                SaveSettings();
            };
            themeMenuItem.DropDownItems.Add(lightThemeItem);
            themeMenuItem.DropDownItems.Add(darkThemeItem);
            
            // 設置選中狀態
            themeMenuItem.DropDownOpening += (s, ev) => {
                lightThemeItem.Checked = !isDarkTheme;
                darkThemeItem.Checked = isDarkTheme;
            };
            
            trayMenu.Items.Add(themeMenuItem);
            // 在托盘菜单中加入“幫助”，作用与界面右键的 Help_Click 相同
            trayMenu.Items.Add("幫助", null, Help_Click);
            // 添加“關於”菜单项（与切換主題同组）
            trayMenu.Items.Add("關於", null, About_Click);
            
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("退出", null, (s, ev) => {
                Application.Exit();
            });
            trayIcon.ContextMenuStrip = trayMenu;
            
            // 左键点击托盘图标也显示菜单
            trayIcon.MouseClick += (s, ev) => {
                if (ev.Button == MouseButtons.Left)
                {
                    // 使用反射调用显示菜单的内部方法
                    System.Reflection.MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    if (mi != null)
                    {
                        mi.Invoke(trayIcon, null);
                    }
                }
            };

                        // 尝试加载自定义图标
            try
            {
                string exePath = Application.ExecutablePath;
                string exeDir = Path.GetDirectoryName(exePath);
                string icoPath = Path.Combine(exeDir, "icon.ico");
                
                // 优先尝试加载 icon.ico 文件
                if (File.Exists(icoPath))
                {
                    trayIcon.Icon = new Icon(icoPath);
                    this.Icon = new Icon(icoPath);
                }
                else if (!exePath.StartsWith("\\\\"))
                {
                    // 只在非 UNC 路径时尝试提取图标
                    Icon extractedIcon = Icon.ExtractAssociatedIcon(exePath);
                    if (extractedIcon != null)
                    {
                        this.Icon = extractedIcon;
                        trayIcon.Icon = extractedIcon;
                    }
                }
                // UNC 路径且无 icon.ico 时，使用默认图标（已在初始化时设置）
            }
            catch
            {
                // 加载失败时使用默认图标（已在初始化时设置）
            }

            // 窗体事件
            this.Load += MainForm_Load;
            this.Deactivate += MainForm_Deactivate;
            this.KeyDown += MainForm_KeyDown;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // 啟動時顯示窗口
            this.Show();
            this.Activate();
            this.BringToFront();
            this.TopMost = true;
            
            // 让列表面板获得焦点，并设置默认选中第一项
            listPanel.Focus();
            if (clipboardHistory.Count > 0)
            {
                selectedIndex = 0;
                listPanel.Invalidate();
            }
            
            // 首次運行提示
            if (firstRun)
            {
                ShowFirstRunTip();
            }
        }

        private void MainForm_Deactivate(object sender, EventArgs e)
        {
            // 不自动隱藏，保持置顶
            // this.Hide();
        }
        
        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 如果搜索框有焦点，只处理 Ctrl+F 和 Escape
            if (searchBox.Focused)
            {
                // Ctrl+F 在搜索框中无需处理
                if (e.Control && e.KeyCode == Keys.F)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                return;
            }
            
            // Ctrl+F 聚焦到搜索框
            if (e.Control && e.KeyCode == Keys.F)
            {
                if (searchBox.Text == "搜索……")
                {
                    searchBox.ReadOnly = false;
                    searchBox.Text = "";
                    searchBox.ForeColor = isDarkTheme ? Color.White : Color.Black;
                }
                searchBox.Focus();
                searchBox.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            
            // 其他键盘操作（上下键、PageUp/Down等）
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            if (displayList.Count == 0)
                return;

            int oldIndex = selectedIndex;
            
            switch (e.KeyCode)
            {
                case Keys.Up:
                    if (selectedIndex > 0)
                        selectedIndex--;
                    else
                        selectedIndex = 0;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.Down:
                    if (selectedIndex < displayList.Count - 1)
                        selectedIndex++;
                    else if (selectedIndex < 0)
                        selectedIndex = 0;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.PageUp:
                    int pageSize = listPanel.ClientSize.Height / ITEM_HEIGHT;
                    selectedIndex = Math.Max(0, selectedIndex - pageSize);
                    if (selectedIndex < 0)
                        selectedIndex = 0;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.PageDown:
                    pageSize = listPanel.ClientSize.Height / ITEM_HEIGHT;
                    selectedIndex = Math.Min(displayList.Count - 1, selectedIndex + pageSize);
                    if (selectedIndex < 0)
                        selectedIndex = 0;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.Home:
                    selectedIndex = 0;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.End:
                    selectedIndex = displayList.Count - 1;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                    
                case Keys.Enter:
                    if (selectedIndex >= 0 && selectedIndex < displayList.Count)
                    {
                        PasteItem(displayList[selectedIndex]);
                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
            }
            
            // 如果选中项改变，确保其可见并重绘
            if (oldIndex != selectedIndex)
            {
                EnsureSelectedVisible();
                listPanel.Invalidate();
            }
        }
        
        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            if (searchBox.Text == "搜索……")
            {
                searchFilter = "";
                searchClearButton.Visible = false;
            }
            else
            {
                searchFilter = searchBox.Text;
                searchClearButton.Visible = !string.IsNullOrEmpty(searchBox.Text);
            }
            scrollOffset = 0;
            selectedIndex = -1;
            CloseImagePreview(this, EventArgs.Empty);
            listPanel.Invalidate();
        }
        
        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            // 在搜索框中按下 Escape 键清除搜索并返回列表
            if (e.KeyCode == Keys.Escape)
            {
                searchBox.Text = "搜索……";
                searchBox.ForeColor = isDarkTheme ? Color.FromArgb(180, 210, 240) : Color.FromArgb(60, 60, 60);
                searchBox.ReadOnly = true;
                searchFilter = "";
                searchClearButton.Visible = false;
                listPanel.Focus();
                e.Handled = true;
            }
        }
        
        private void ListPanel_Paint(object sender, PaintEventArgs e)
        {
            // 获取要显示的列表（支持搜索过滤）
            List<ClipboardItem> displayList = clipboardHistory;
            
            // 应用收藏过滤
            if (showingFavorites)
            {
                displayList = displayList.Where(item => item.IsFavorite).ToList();
            }
            
            // 应用搜索过滤
            if (!string.IsNullOrWhiteSpace(searchFilter))
            {
                displayList = displayList.Where(item => 
                    item.Text != null && item.Text.IndexOf(searchFilter, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();
            }
            
            if (displayList.Count == 0)
                return;

            int y = -scrollOffset;
            int panelWidth = listPanel.ClientSize.Width;
            int startIndex = Math.Max(0, scrollOffset / ITEM_HEIGHT);
            int visibleCount = (listPanel.ClientSize.Height / ITEM_HEIGHT) + 2;
            
            for (int i = startIndex; i < Math.Min(startIndex + visibleCount, displayList.Count); i++)
            {
                int itemY = i * ITEM_HEIGHT - scrollOffset;
                
                // 只繪製可見的項目
                if (itemY + ITEM_HEIGHT < 0 || itemY > listPanel.ClientSize.Height)
                    continue;

                ClipboardItem item = displayList[i];
                Rectangle itemRect = new Rectangle(0, itemY, panelWidth, ITEM_HEIGHT);
                
                // 收藏项使用黄色背景，普通项隔行换色
                Color itemBgColor;
                if (item.IsFavorite)
                {
                    itemBgColor = (i % 2 == 0) 
                        ? (isDarkTheme ? favoriteBgColor1Dark : favoriteBgColor1Light)
                        : (isDarkTheme ? favoriteBgColor2Dark : favoriteBgColor2Light);
                }
                else
                {
                    itemBgColor = (i % 2 == 0) ? bgColor1 : bgColor2;
                }
                
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
                        int imageIndex = GetImageIndexForItem(item);
                        if (imageIndex >= 0 && imageIndex < imageList.Images.Count)
                        {
                            Image thumbnail = imageList.Images[imageIndex];
                            int imgX = panelWidth - 58;
                            int imgY = itemY + (ITEM_HEIGHT - 54) / 2;
                            e.Graphics.DrawImage(thumbnail, imgX, imgY, 54, 54);
                            textWidth = panelWidth - 68;
                        }
                    }
                    catch { }
                }

                // 给製文字（左對齊）- 优化：限制显示长度，避免长文本卡顿
                string displayText = item.ToString();
                if (displayText.Length > 144)
                {
                    displayText = displayText.Substring(0, 144) + "...";
                }
                
                Rectangle textRect = new Rectangle(5, itemY + 5, textWidth, ITEM_HEIGHT - 10);
                
                // 如果是选中项，绘制高亮边框
                if (i == selectedIndex)
                {
                    using (Pen highlightPen = new Pen(Color.FromArgb(0, 120, 215), 2))
                    {
                        e.Graphics.DrawRectangle(highlightPen, new Rectangle(1, itemY + 1, panelWidth - 2, ITEM_HEIGHT - 2));
                    }
                    
                    // 如果是图片，在中央显示"預覽"按钮
                    if (item.Format == "Image" && item.Data != null)
                    {
                        string buttonText = "預覽";
                        Font buttonFont = new Font("微软雅黑", 10F, FontStyle.Bold);
                        SizeF textSize = e.Graphics.MeasureString(buttonText, buttonFont);
                        int buttonWidth = (int)textSize.Width + 20;
                        int buttonHeight = (int)textSize.Height + 10;
                        int buttonX = (panelWidth - buttonWidth) / 2;
                        int buttonY = itemY + (ITEM_HEIGHT - buttonHeight) / 2;
                        
                        Rectangle buttonRect = new Rectangle(buttonX, buttonY, buttonWidth, buttonHeight);
                        
                        // 绘制按钮背景和边框
                        using (SolidBrush buttonBrush = new SolidBrush(Color.FromArgb(0, 120, 215)))
                        using (Pen buttonPen = new Pen(Color.FromArgb(0, 90, 180), 2))
                        {
                            e.Graphics.FillRectangle(buttonBrush, buttonRect);
                            e.Graphics.DrawRectangle(buttonPen, buttonRect);
                        }
                        
                        // 绘制按钮文字
                        TextRenderer.DrawText(e.Graphics, buttonText, buttonFont, buttonRect,
                            Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                    }
                }
                
                TextRenderer.DrawText(e.Graphics, displayText, listPanel.Font, textRect, 
                    textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
            
            // 更新滚动条
            UpdateScrollBar();
        }

        private void ListPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // 检查是否点击了预览按钮区域
                int itemIndex = GetItemIndexAtPoint(e.Location);
                if (itemIndex >= 0 && itemIndex == selectedIndex)
                {
                    List<ClipboardItem> displayList = GetFilteredDisplayList();
                    if (itemIndex < displayList.Count)
                    {
                        ClipboardItem item = displayList[itemIndex];
                        if (item.Format == "Image" && item.Data != null)
                        {
                            // 计算按钮区域
                            int itemY = itemIndex * ITEM_HEIGHT - scrollOffset;
                            int panelWidth = listPanel.ClientSize.Width;
                            
                            using (Graphics g = listPanel.CreateGraphics())
                            {
                                Font buttonFont = new Font("微软雅黑", 10F, FontStyle.Bold);
                                SizeF textSize = g.MeasureString("預覽", buttonFont);
                                int buttonWidth = (int)textSize.Width + 20;
                                int buttonHeight = (int)textSize.Height + 10;
                                int buttonX = (panelWidth - buttonWidth) / 2;
                                int buttonY = itemY + (ITEM_HEIGHT - buttonHeight) / 2;
                                
                                Rectangle buttonRect = new Rectangle(buttonX, buttonY, buttonWidth, buttonHeight);
                                
                                // 如果点击在按钮区域内，不关闭预览并直接显示
                                if (buttonRect.Contains(e.Location))
                                {
                                    ShowImagePreview();
                                    return;
                                }
                            }
                        }
                    }
                }
                
                // 不是点击预览按钮，关闭预览
                CloseImagePreview(this, EventArgs.Empty);
                
                // 不立即设置 isDragging，等待 MouseMove 时再判断
                dragStartPoint = e.Location;
            }
            else
            {
                // 右键按下时关闭预览
                CloseImagePreview(this, EventArgs.Empty);
            }
        }

        private void ListPanel_MouseClick(object sender, MouseEventArgs e)
        {
            // 点击列表时移除搜索框焦点
            if (searchBox.Focused)
            {
                listPanel.Focus();
            }
            
            if (e.Button == MouseButtons.Left)
            {
                // 设置选中项
                int itemIndex = GetItemIndexAtPoint(e.Location);
                if (itemIndex >= 0)
                {
                    selectedIndex = itemIndex;
                    listPanel.Invalidate();
                    listPanel.Focus();
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                // 保存右键点击位置
                lastContextMenuPosition = e.Location;
                
                int itemIndex = GetItemIndexAtPoint(e.Location);
                if (itemIndex >= 0)
                {
                    selectedIndex = itemIndex;
                    listPanel.Invalidate();
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
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                
                if (itemIndex >= 0 && itemIndex < displayList.Count)
                {
                    PasteItem(displayList[itemIndex]);
                }
            }
        }

        private void ListPanel_MouseWheel(object sender, MouseEventArgs e)
        {
            // 滚动时关闭预览
            CloseImagePreview(this, EventArgs.Empty);
            
            int delta = e.Delta / 120;
            scrollOffset -= delta * 60;
            
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            
            int totalHeight = displayList.Count * ITEM_HEIGHT;
            int maxScroll = Math.Max(0, totalHeight - listPanel.ClientSize.Height);
            scrollOffset = Math.Max(0, Math.Min(scrollOffset, maxScroll));
            
            listPanel.Invalidate();
        }

        private void EnsureSelectedVisible()
        {
            if (selectedIndex < 0)
                return;
                
            int selectedY = selectedIndex * ITEM_HEIGHT;
            int viewportTop = scrollOffset;
            int viewportBottom = scrollOffset + listPanel.ClientSize.Height;
            
            // 如果选中项在视口上方
            if (selectedY < viewportTop)
            {
                scrollOffset = selectedY;
            }
            // 如果选中项在视口下方
            else if (selectedY + ITEM_HEIGHT > viewportBottom)
            {
                scrollOffset = selectedY + ITEM_HEIGHT - listPanel.ClientSize.Height;
            }
            
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            int totalHeight = displayList.Count * ITEM_HEIGHT;
            int maxScroll = Math.Max(0, totalHeight - listPanel.ClientSize.Height);
            scrollOffset = Math.Max(0, Math.Min(scrollOffset, maxScroll));
        }

        private void ScrollBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                scrollBarDragging = true;
                // 记录滚动条当前的Top位置（相对于listPanel）
                scrollBarDragStart = scrollBarPanel.Top;
                scrollOffsetDragStart = scrollOffset;
            }
        }

        private void ScrollBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (scrollBarDragging)
            {
                // 计算滚动条的新位置（鼠标相对于listPanel的Y坐标 - 鼠标在滚动条内的相对位置）
                Point mouseInPanel = listPanel.PointToClient(Control.MousePosition);
                int newScrollBarTop = mouseInPanel.Y - (scrollBarPanel.Height / 2);
                
                // 限制滚动条在有效范围内
                int maxScrollBarTop = listPanel.ClientSize.Height - scrollBarPanel.Height;
                newScrollBarTop = Math.Max(0, Math.Min(newScrollBarTop, maxScrollBarTop));
                
                // 根据滚动条位置计算scrollOffset
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                
                int totalHeight = displayList.Count * ITEM_HEIGHT;
                int maxScroll = Math.Max(0, totalHeight - listPanel.ClientSize.Height);
                
                if (maxScrollBarTop > 0)
                {
                    // 正确计算scrollOffset：当滚动条在底部时，scrollOffset应该等于maxScroll
                    scrollOffset = (newScrollBarTop * maxScroll) / maxScrollBarTop;
                    scrollOffset = Math.Max(0, Math.Min(scrollOffset, maxScroll));
                    listPanel.Invalidate();
                }
            }
        }

        private void ScrollBar_MouseUp(object sender, MouseEventArgs e)
        {
            scrollBarDragging = false;
        }

        private void UpdateScrollBar()
        {
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            
            if (displayList.Count == 0)
            {
                scrollBarPanel.Visible = false;
                return;
            }

            int totalHeight = displayList.Count * ITEM_HEIGHT;
            if (totalHeight <= listPanel.ClientSize.Height)
            {
                scrollBarPanel.Visible = false;
                return;
            }

            scrollBarPanel.Visible = true;
            // 設置最小高度為60px，確保大量記錄時也好操作
            int scrollBarHeight = Math.Max(60, (listPanel.ClientSize.Height * listPanel.ClientSize.Height) / totalHeight);
            
            // 计算滚动条位置
            int maxScroll = Math.Max(0, totalHeight - listPanel.ClientSize.Height);
            int maxScrollBarY = listPanel.ClientSize.Height - scrollBarHeight;
            int scrollBarY = 0;
            
            if (maxScroll > 0 && maxScrollBarY > 0)
            {
                scrollBarY = (scrollOffset * maxScrollBarY) / maxScroll;
            }
            
            // 确保滚动条不超出listPanel的范围
            scrollBarY = Math.Max(0, Math.Min(scrollBarY, maxScrollBarY));

            scrollBarPanel.Height = scrollBarHeight;
            scrollBarPanel.Top = scrollBarY;
            scrollBarPanel.Left = listPanel.Right - 10;
        }

        private void MinimizeButton_Click(object sender, EventArgs e)
        {
            CloseImagePreview(this, EventArgs.Empty);
            this.Hide();
        }
        
        private void FavoritesButton_Click(object sender, EventArgs e)
        {
            showingFavorites = !showingFavorites;
            favoritesButton.Text = showingFavorites ? "返回首頁" : "我的收藏";
            
            // 切换视图时清空搜索
            searchBox.Text = "搜索……";
            searchBox.ForeColor = isDarkTheme ? Color.FromArgb(180, 210, 240) : Color.FromArgb(60, 60, 60);
            searchFilter = "";
            
            scrollOffset = 0;
            selectedIndex = -1;
            CloseImagePreview(this, EventArgs.Empty);
            listPanel.Invalidate();
        }
        
        private void ToggleFavorite_Click(object sender, EventArgs e)
        {
            Point mousePos = listPanel.PointToClient(Control.MousePosition);
            int itemIndex = GetItemIndexAtPoint(mousePos);
            
            if (itemIndex >= 0)
            {
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                
                if (itemIndex < displayList.Count)
                {
                    displayList[itemIndex].IsFavorite = !displayList[itemIndex].IsFavorite;
                    listPanel.Invalidate();
                    SaveHistory();
                }
            }
        }
        
        private List<ClipboardItem> GetFilteredDisplayList()
        {
            List<ClipboardItem> displayList = clipboardHistory;
            
            // 应用收藏过滤
            if (showingFavorites)
            {
                displayList = displayList.Where(item => item.IsFavorite).ToList();
            }
            
            // 应用搜索过滤
            if (!string.IsNullOrWhiteSpace(searchFilter))
            {
                displayList = displayList.Where(item => 
                    item.Text != null && item.Text.IndexOf(searchFilter, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();
            }
            
            return displayList;
        }
        
        private void ListContextMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Point mousePos = listPanel.PointToClient(Control.MousePosition);
            int itemIndex = GetItemIndexAtPoint(mousePos);
            
            // 在收藏界面隐藏"清空"选项（根据文本查找以避免索引变动）
            var clearItem = listContextMenu.Items.Cast<ToolStripItem>().FirstOrDefault(it => it.Text == "清空");
            if (clearItem != null)
            {
                clearItem.Visible = !showingFavorites;
            }

            // 同步右键菜单中主题子菜单的选中状态
            var found = listContextMenu.Items.Find("ctxThemeMenu", false);
            if (found != null && found.Length > 0)
            {
                var themeMenu = found[0] as ToolStripMenuItem;
                if (themeMenu != null && themeMenu.DropDownItems.Count >= 2)
                {
                    var lightItem = themeMenu.DropDownItems[0] as ToolStripMenuItem;
                    var darkItem = themeMenu.DropDownItems[1] as ToolStripMenuItem;
                    if (lightItem != null) lightItem.Checked = !isDarkTheme;
                    if (darkItem != null) darkItem.Checked = isDarkTheme;
                }
            }
            
            if (itemIndex >= 0)
            {
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                
                if (itemIndex < displayList.Count)
                {
                    ClipboardItem item = displayList[itemIndex];
                    bool isFavorite = item.IsFavorite;
                    // 更新收藏菜单项的文字（通过文本查找，避免索引依赖）
                    var favItem = listContextMenu.Items.Cast<ToolStripItem>().FirstOrDefault(it => it.Text == "收藏" || it.Text == "取消收藏");
                    if (favItem != null)
                    {
                        favItem.Text = isFavorite ? "取消收藏" : "收藏";
                    }
                }
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
                // 优化：先设置不透明再显示，避免渐变效果
                this.Opacity = 1.0;
                this.Show();
                this.Activate();
                this.BringToFront();
                this.TopMost = true;
                // 只重绘，不重建缩略图
                listPanel.Invalidate();
                // 让列表面板获得焦点
                listPanel.Focus();
            }
        }

        private int GetItemIndexAtPoint(Point point)
        {
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            
            int itemIndex = (scrollOffset + point.Y) / ITEM_HEIGHT;
            if (itemIndex >= 0 && itemIndex < displayList.Count)
            {
                return itemIndex;
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
        
        private int GetImageIndexForItem(ClipboardItem targetItem)
        {
            int imageIndex = 0;
            foreach (var item in clipboardHistory)
            {
                if (item == targetItem)
                {
                    return imageIndex;
                }
                if (item.Format == "Image" && item.Data != null)
                {
                    imageIndex++;
                }
            }
            return -1;
        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                // 拖动时关闭预览
                CloseImagePreview(this, EventArgs.Empty);
                isDragging = true;
                dragStartPoint = e.Location;
            }
        }

        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && !isDragging)
            {
                // 检测是否真的在拖动（移动超过5像素才算拖动）
                int deltaX = Math.Abs(e.Location.X - dragStartPoint.X);
                int deltaY = Math.Abs(e.Location.Y - dragStartPoint.Y);
                if (deltaX > 5 || deltaY > 5)
                {
                    isDragging = true;
                }
            }
            
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
                // 限制窗口在屏幕工作区内
                Screen currentScreen = Screen.FromPoint(this.Location);
                Rectangle workingArea = currentScreen.WorkingArea;
                
                int x = this.Location.X;
                int y = this.Location.Y;
                
                // 确保窗口不超出屏幕边界
                if (x < workingArea.Left) x = workingArea.Left;
                if (y < workingArea.Top) y = workingArea.Top;
                if (x + this.Width > workingArea.Right) x = workingArea.Right - this.Width;
                if (y + this.Height > workingArea.Bottom) y = workingArea.Bottom - this.Height;
                
                this.Location = new Point(x, y);
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
                    // 优化：直接显示，减少不必要的调用
                    this.Opacity = 1.0;
                    this.Visible = true;
                    this.Activate();
                    this.BringToFront();
                    this.TopMost = true;
                    // 只重绘，不重建缩略图
                    listPanel.Invalidate();
                    // 让列表面板获得焦点
                    listPanel.Focus();
                }
            }
            base.WndProc(ref m);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 如果搜索框有焦点，不拦截键盘事件
            if (searchBox.Focused)
            {
                return base.ProcessCmdKey(ref msg, keyData);
            }

            List<ClipboardItem> displayList = GetFilteredDisplayList();
            if (displayList.Count == 0)
                return base.ProcessCmdKey(ref msg, keyData);

            // 支持 Ctrl+C 复制当前选中记录
            if ((keyData & Keys.Control) == Keys.Control && (keyData & Keys.KeyCode) == Keys.C)
            {
                if (selectedIndex >= 0 && selectedIndex < displayList.Count)
                {
                    var item = displayList[selectedIndex];
                    try
                    {
                        if (item.Format == "Text")
                        {
                            Clipboard.SetText(item.Text ?? "");
                        }
                        else if (item.Format == "RTF")
                        {
                            string rtf = Encoding.UTF8.GetString(item.Data ?? new byte[0]);
                            Clipboard.SetData(DataFormats.Rtf, rtf);
                        }
                        else if (item.Format == "HTML")
                        {
                            string html = Encoding.UTF8.GetString(item.Data ?? new byte[0]);
                            Clipboard.SetData(DataFormats.Html, html);
                        }
                        else if (item.Format == "Image" && item.Data != null)
                        {
                            using (MemoryStream ms = new MemoryStream(item.Data))
                            {
                                Image img = Image.FromStream(ms);
                                Clipboard.SetImage(img);
                            }
                        }
                    }
                    catch { }
                }
                return true;
            }

            int oldIndex = selectedIndex;
            bool handled = false;

            switch (keyData)
            {
                case Keys.Up:
                    if (selectedIndex > 0)
                        selectedIndex--;
                    else if (selectedIndex < 0)
                        selectedIndex = 0;
                    handled = true;
                    break;

                case Keys.Down:
                    if (selectedIndex < displayList.Count - 1)
                        selectedIndex++;
                    else if (selectedIndex < 0)
                        selectedIndex = 0;
                    handled = true;
                    break;

                case Keys.PageUp:
                    int pageSize = listPanel.ClientSize.Height / ITEM_HEIGHT;
                    selectedIndex = Math.Max(0, selectedIndex - pageSize);
                    if (selectedIndex < 0)
                        selectedIndex = 0;
                    handled = true;
                    break;

                case Keys.PageDown:
                    pageSize = listPanel.ClientSize.Height / ITEM_HEIGHT;
                    selectedIndex = Math.Min(displayList.Count - 1, selectedIndex + pageSize);
                    if (selectedIndex < 0)
                        selectedIndex = 0;
                    handled = true;
                    break;

                case Keys.Home:
                    selectedIndex = 0;
                    handled = true;
                    break;

                case Keys.End:
                    selectedIndex = displayList.Count - 1;
                    handled = true;
                    break;

                case Keys.Enter:
                    if (selectedIndex >= 0 && selectedIndex < displayList.Count)
                    {
                        PasteItem(displayList[selectedIndex]);
                        handled = true;
                    }
                    break;
                    
                case Keys.Escape:
                    this.Hide();
                    handled = true;
                    break;
            }

            if (handled && oldIndex != selectedIndex)
            {
                EnsureSelectedVisible();
                listPanel.Invalidate();
                // 选中项改变时关闭预览
                CloseImagePreview(this, EventArgs.Empty);
            }

            return handled || base.ProcessCmdKey(ref msg, keyData);
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
                    this.Invoke((MethodInvoker)delegate {
                        // 只在添加图片时重建缩略图
                        if (item.Format == "Image")
                        {
                            RebuildImageList();
                        }
                        RefreshListView();
                        SaveHistory();
                    });
                }
                else
                {
                    // 只在添加图片时重建缩略图
                    if (item.Format == "Image")
                    {
                        RebuildImageList();
                    }
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
            // 只重绘列表，不重新加载图片
            listPanel.Invalidate();
        }
        
        private void RebuildImageList()
        {
            // 重新加载图片缩略图
            imageList.Images.Clear();
            
            foreach (var item in clipboardHistory)
            {
                // 為圖片添加縮略圖
                if (item.Format == "Image" && item.Data != null)
                {
                    try
                    {
                        using (MemoryStream ms = new MemoryStream(item.Data))
                        {
                            Image img = Image.FromStream(ms);
                            Image thumbnail = CreateThumbnail(img, 54, 54);
                            imageList.Images.Add(thumbnail);
                        }
                    }
                    catch { }
                }
            }
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

                // 粘贴后隱藏界面
                this.Hide();

                // 等待更長時間確保貼上完成，避免被清空
                Thread.Sleep(300);

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

        private void EditItem_Click(object sender, EventArgs e)
        {
            // 使用保存的右键点击位置
            int itemIndex = GetItemIndexAtPoint(lastContextMenuPosition);
            
            if (itemIndex >= 0)
            {
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                if (itemIndex < displayList.Count)
                {
                    ClipboardItem item = displayList[itemIndex];
                    
                    if (item.Format == "Text")
                    {
                        // 文本类型：创建编辑对话框
                        Form editForm = new Form();
                        editForm.Text = "編輯內容";
                        editForm.Size = new Size(500, 400);
                        editForm.StartPosition = FormStartPosition.CenterParent;
                        editForm.FormBorderStyle = FormBorderStyle.Sizable;
                        editForm.MinimizeBox = false;
                        editForm.MaximizeBox = true;
                        editForm.TopMost = true;
                        
                        // 使用主程序的图标
                        if (this.Icon != null)
                        {
                            editForm.Icon = this.Icon;
                        }
                        
                        TextBox editBox = new TextBox();
                        editBox.Multiline = true;
                        editBox.ScrollBars = ScrollBars.Both;
                        editBox.Dock = DockStyle.Fill;
                        editBox.Font = new Font("Consolas", 10F);
                        editBox.Text = item.Text;
                        editBox.Padding = new Padding(5);
                        editBox.SelectionStart = 0;
                        editBox.SelectionLength = 0;
                        editForm.Controls.Add(editBox);
                        
                        Panel buttonPanel = new Panel();
                        buttonPanel.Height = 50;
                        buttonPanel.Dock = DockStyle.Bottom;
                        editForm.Controls.Add(buttonPanel);
                        
                        Button saveButton = new Button();
                        saveButton.Text = "保存";
                        saveButton.Size = new Size(80, 30);
                        saveButton.Location = new Point(160, 10);
                        saveButton.Click += (s, ev) => {
                            item.Text = editBox.Text;
                            item.Time = DateTime.Now;
                            SaveHistory();
                            listPanel.Invalidate();
                            editForm.Close();
                        };
                        buttonPanel.Controls.Add(saveButton);
                        
                        Button cancelButton = new Button();
                        cancelButton.Text = "取消";
                        cancelButton.Size = new Size(80, 30);
                        cancelButton.Location = new Point(260, 10);
                        cancelButton.Click += (s, ev) => {
                            editForm.Close();
                        };
                        buttonPanel.Controls.Add(cancelButton);
                        
                        editForm.ShowDialog(this);
                    }
                    else if (item.Format == "Image" && item.Data != null)
                    {
                        // 图片类型：保存为临时文件并用外部程序打开
                        try
                        {
                            string tempPath = Path.Combine(Path.GetTempPath(), "MyClipboard_" + DateTime.Now.Ticks + ".png");
                            File.WriteAllBytes(tempPath, item.Data);
                            
                            // 使用 Paint 打开图片
                            System.Diagnostics.Process.Start("mspaint.exe", "\""+tempPath+"\"");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("打開編輯器失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        
        private void CopyItem_Click(object sender, EventArgs e)
        {
            // 使用保存的右键点击位置
            int itemIndex = GetItemIndexAtPoint(lastContextMenuPosition);
            
            if (itemIndex >= 0)
            {
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                if (itemIndex < displayList.Count)
                {
                    ClipboardItem item = displayList[itemIndex];
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
        }

        private void DeleteItem_Click(object sender, EventArgs e)
        {
            // 使用保存的右键点击位置
            int itemIndex = GetItemIndexAtPoint(lastContextMenuPosition);
            
            if (itemIndex >= 0)
            {
                List<ClipboardItem> displayList = GetFilteredDisplayList();
                if (itemIndex < displayList.Count)
                {
                    ClipboardItem itemToDelete = displayList[itemIndex];
                    clipboardHistory.Remove(itemToDelete);
                    RefreshListView();
                    SaveHistory();
                }
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
                    writer.Write(firstRun);
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
                        
                        if (fs.Position < fs.Length)
                        {
                            firstRun = reader.ReadBoolean();
                        }
                        return;
                    }
                }
            }
            catch { }
            
            // 如果沒有保存的設置，默認居中和深色主題
            this.StartPosition = FormStartPosition.CenterScreen;
            isDarkTheme = true;
            firstRun = true;
        }

        private void ShowFirstRunTip()
        {
            if (!firstRun)
                return;

            // 创建自定义对话框
            Form tipForm = new Form();
            tipForm.Text = "使用提示";
            tipForm.Size = new Size(400, 250);
            tipForm.StartPosition = FormStartPosition.CenterScreen;
            tipForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            tipForm.MaximizeBox = false;
            tipForm.MinimizeBox = false;
            tipForm.TopMost = true;

            Label messageLabel = new Label();
            messageLabel.Text = "歡迎使用 MyClipboard！\n\n" +
                "快捷鍵：Ctrl + Alt + X 顯示 / 隱藏界面；\n\n" +
                "雙擊記錄可直接粘貼。";
            messageLabel.AutoSize = false;
            messageLabel.Size = new Size(360, 120);
            messageLabel.Location = new Point(20, 20);
            messageLabel.Font = new Font("Consolas", 10F);
            tipForm.Controls.Add(messageLabel);

            CheckBox dontShowCheckbox = new CheckBox();
            dontShowCheckbox.Text = "不再顯示此提示";
            dontShowCheckbox.Location = new Point(20, 150);
            dontShowCheckbox.AutoSize = true;
            tipForm.Controls.Add(dontShowCheckbox);

            Button okButton = new Button();
            okButton.Text = "確定";
            okButton.Size = new Size(80, 30);
            okButton.Location = new Point(160, 180);
            okButton.Click += (s, ev) => {
                if (dontShowCheckbox.Checked)
                {
                    firstRun = false;
                    SaveSettings();
                }
                tipForm.Close();
            };
            tipForm.Controls.Add(okButton);

            tipForm.ShowDialog();
        }

        private void About_Click(object sender, EventArgs e)
        {
            // 临时取消主窗口置顶，避免遮挡弹窗
            bool wasTopMost = this.TopMost;
            this.TopMost = false;
            
            MessageBox.Show("聯繫微信：676400126", "關於", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            // 恢复置顶状态
            this.TopMost = wasTopMost;
        }

        private void Help_Click(object sender, EventArgs e)
        {
            // 临时取消主窗口置顶，避免遮挡弹窗
            bool wasTopMost = this.TopMost;
            this.TopMost = false;
            
            // 创建帮助对话框
            Form helpForm = new Form();
            helpForm.Text = "使用提示";
            helpForm.Size = new Size(400, 220);
            helpForm.StartPosition = FormStartPosition.CenterScreen;
            helpForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            helpForm.MaximizeBox = false;
            helpForm.MinimizeBox = false;
            helpForm.TopMost = true;
            
            // 使用主程序的图标
            if (this.Icon != null)
            {
                helpForm.Icon = this.Icon;
            }

            Label messageLabel = new Label();
            messageLabel.Text = "歡迎使用 MyClipboard！\n\n" +
                "快捷鍵：Ctrl + Alt + X 顯示 / 隱藏界面；\n\n" +
                "雙擊記錄可直接粘貼。";
            messageLabel.AutoSize = false;
            messageLabel.Size = new Size(360, 120);
            messageLabel.Location = new Point(20, 20);
            messageLabel.Font = new Font("Consolas", 10F);
            helpForm.Controls.Add(messageLabel);

            Button okButton = new Button();
            okButton.Text = "確定";
            okButton.Size = new Size(80, 30);
            okButton.Location = new Point(160, 150);
            okButton.Click += (s, ev) => {
                helpForm.Close();
            };
            helpForm.Controls.Add(okButton);

            helpForm.ShowDialog(this);
            
            // 恢复置顶状态
            this.TopMost = wasTopMost;
        }

        private void ApplyTheme()
        {
            if (isDarkTheme)
            {
                bgColor1 = Color.FromArgb(30, 30, 32);
                bgColor2 = Color.FromArgb(50, 50, 52);
                textColor = Color.White;
                borderColor = Color.FromArgb(63, 63, 70);
                this.BackColor = Color.FromArgb(63, 63, 70);
                
                // 深色主题收藏颜色（较暗的黄色）
                favoriteBgColor1Dark = Color.FromArgb(70, 70, 30);
                favoriteBgColor2Dark = Color.FromArgb(90, 90, 40);
                
                if (listPanel != null)
                {
                    listPanel.BackColor = Color.FromArgb(30, 30, 30);
                    listPanel.ForeColor = Color.White;
                }
                if (minimizeButton != null)
                {
                    minimizeButton.BackColor = Color.FromArgb(60, 60, 60);
                    minimizeButton.ForeColor = Color.White;
                }
                if (favoritesButton != null)
                {
                    favoritesButton.BackColor = Color.FromArgb(60, 60, 60);
                    favoritesButton.ForeColor = Color.White;
                }
                if (searchBox != null && searchBox.Text != "搜索……")
                {
                    searchBox.BackColor = Color.FromArgb(0, 90, 158);
                    searchBox.ForeColor = Color.White;
                }
                else if (searchBox != null)
                {
                    searchBox.BackColor = Color.FromArgb(0, 90, 158);
                    searchBox.ForeColor = Color.FromArgb(180, 210, 240);
                }
                if (searchPanel != null)
                {
                    searchPanel.BackColor = Color.FromArgb(0, 90, 158);
                }
                if (searchClearButton != null)
                {
                    searchClearButton.BackColor = Color.FromArgb(0, 90, 158);
                    searchClearButton.ForeColor = Color.White;
                }
            }
            else
            {
                bgColor1 = Color.FromArgb(240, 240, 240);
                bgColor2 = Color.FromArgb(255, 255, 255);
                textColor = Color.Black;
                borderColor = Color.FromArgb(180, 180, 180);
                this.BackColor = Color.FromArgb(180, 180, 180);
                
                // 浅色主题收藏颜色（浅黄色）
                favoriteBgColor1Light = Color.FromArgb(255, 255, 220);
                favoriteBgColor2Light = Color.FromArgb(255, 255, 200);
                
                if (listPanel != null)
                {
                    listPanel.BackColor = Color.FromArgb(250, 250, 250);
                    listPanel.ForeColor = Color.Black;
                }
                if (minimizeButton != null)
                {
                    minimizeButton.BackColor = Color.FromArgb(220, 220, 220);
                    minimizeButton.ForeColor = Color.Black;
                }
                if (favoritesButton != null)
                {
                    favoritesButton.BackColor = Color.FromArgb(220, 220, 220);
                    favoritesButton.ForeColor = Color.Black;
                }
                if (searchBox != null && searchBox.Text != "搜索……")
                {
                    searchBox.BackColor = Color.FromArgb(180, 210, 255);
                    searchBox.ForeColor = Color.Black;
                }
                else if (searchBox != null)
                {
                    searchBox.BackColor = Color.FromArgb(180, 210, 255);
                    searchBox.ForeColor = Color.FromArgb(80, 80, 80);
                }
                if (searchPanel != null)
                {
                    searchPanel.BackColor = Color.FromArgb(180, 210, 255);
                }
                if (searchClearButton != null)
                {
                    searchClearButton.BackColor = Color.FromArgb(180, 210, 255);
                    searchClearButton.ForeColor = Color.FromArgb(60, 60, 60);
                }
            }
            
            if (listPanel != null)
            {
                listPanel.Invalidate();
            }
        }

        private void ShowImagePreview()
        {
            // 关闭已有的预览窗口
            if (imagePreviewForm != null && !imagePreviewForm.IsDisposed)
            {
                imagePreviewForm.Close();
                imagePreviewForm.Dispose();
                imagePreviewForm = null;
            }

            // 检查选中项是否为图片
            List<ClipboardItem> displayList = GetFilteredDisplayList();
            if (selectedIndex < 0 || selectedIndex >= displayList.Count)
                return;

            ClipboardItem selectedItem = displayList[selectedIndex];
            if (selectedItem.Format != "Image" || selectedItem.Data == null)
                return;

            try
            {
                // 加载图片
                Image previewImage;
                using (MemoryStream ms = new MemoryStream(selectedItem.Data))
                {
                    previewImage = Image.FromStream(ms);
                }

                // 创建预览窗口
                imagePreviewForm = new Form();
                imagePreviewForm.FormBorderStyle = FormBorderStyle.None;
                imagePreviewForm.BackColor = Color.Black;
                imagePreviewForm.StartPosition = FormStartPosition.Manual;
                imagePreviewForm.TopMost = true;
                imagePreviewForm.ShowInTaskbar = false;

                // 计算预览窗口大小（最大400x400，保持宽高比）
                int maxSize = 400;
                double scale = Math.Min((double)maxSize / previewImage.Width, (double)maxSize / previewImage.Height);
                if (scale > 1) scale = 1; // 不放大小图
                int previewWidth = (int)(previewImage.Width * scale);
                int previewHeight = (int)(previewImage.Height * scale);
                imagePreviewForm.Size = new Size(previewWidth + 4, previewHeight + 4); // 加4px边框

                // 计算预览窗口位置（尽量在主窗口右侧，不遮挡且不超出屏幕）
                Screen currentScreen = Screen.FromControl(this);
                Rectangle workingArea = currentScreen.WorkingArea;
                
                int previewX = this.Right + 10;
                int previewY = this.Top;

                // 如果右侧空间不足，放在左侧
                if (previewX + imagePreviewForm.Width > workingArea.Right)
                {
                    previewX = this.Left - imagePreviewForm.Width - 10;
                }
                
                // 如果左侧也不足，放在主窗口上方或下方
                if (previewX < workingArea.Left)
                {
                    previewX = this.Left;
                    previewY = this.Bottom + 10;
                    
                    // 如果下方也不足，放在上方
                    if (previewY + imagePreviewForm.Height > workingArea.Bottom)
                    {
                        previewY = this.Top - imagePreviewForm.Height - 10;
                    }
                }

                // 确保不超出屏幕边界
                if (previewX < workingArea.Left) previewX = workingArea.Left;
                if (previewY < workingArea.Top) previewY = workingArea.Top;
                if (previewX + imagePreviewForm.Width > workingArea.Right) previewX = workingArea.Right - imagePreviewForm.Width;
                if (previewY + imagePreviewForm.Height > workingArea.Bottom) previewY = workingArea.Bottom - imagePreviewForm.Height;

                imagePreviewForm.Location = new Point(previewX, previewY);

                // 添加图片显示（带2px边框）
                PictureBox pictureBox = new PictureBox();
                pictureBox.Image = previewImage;
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox.Location = new Point(2, 2);
                pictureBox.Size = new Size(previewWidth, previewHeight);
                pictureBox.BackColor = Color.Black;
                imagePreviewForm.Controls.Add(pictureBox);

                // 绘制边框（与选中项相同的蓝色）
                imagePreviewForm.Paint += (s, ev) => {
                    using (Pen borderPen = new Pen(Color.FromArgb(0, 120, 215), 2))
                    {
                        ev.Graphics.DrawRectangle(borderPen, 0, 0, imagePreviewForm.Width - 1, imagePreviewForm.Height - 1);
                    }
                };

                // 点击预览窗口关闭
                imagePreviewForm.Click += (s, ev) => {
                    imagePreviewForm.Close();
                };
                pictureBox.Click += (s, ev) => {
                    imagePreviewForm.Close();
                };

                // 主窗口失去焦点或隐藏时关闭预览
                this.Deactivate += CloseImagePreview;
                this.VisibleChanged += CloseImagePreview;

                imagePreviewForm.Show();
            }
            catch
            {
                // 图片加载失败时忽略
                if (imagePreviewForm != null)
                {
                    imagePreviewForm.Dispose();
                    imagePreviewForm = null;
                }
            }
        }

        private void CloseImagePreview(object sender, EventArgs e)
        {
            if (imagePreviewForm != null && !imagePreviewForm.IsDisposed)
            {
                imagePreviewForm.Close();
                imagePreviewForm.Dispose();
                imagePreviewForm = null;
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
                        writer.Write(item.IsFavorite);
                        
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
                        
                        // 尝试读取收藏状态（兼容旧版本）
                        if (fs.Position < fs.Length)
                        {
                            try
                            {
                                item.IsFavorite = reader.ReadBoolean();
                            }
                            catch
                            {
                                item.IsFavorite = false;
                            }
                        }
                        
                        int dataLength = reader.ReadInt32();
                        if (dataLength > 0)
                        {
                            item.Data = reader.ReadBytes(dataLength);
                        }
                        
                        clipboardHistory.Add(item);
                    }
                }

                RebuildImageList();
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
            
            if (imagePreviewForm != null && !imagePreviewForm.IsDisposed)
            {
                imagePreviewForm.Close();
                imagePreviewForm.Dispose();
            }
            
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
