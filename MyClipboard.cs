using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace MyClipboardApp
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ClipboardForm());
        }
    }

    public sealed class ClipboardForm : Form
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 0xA001;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const int KEYEVENTF_KEYUP = 0x0002;

        private readonly NotifyIcon _trayIcon;
        private readonly BufferedListView _recordsView;
        private readonly ContextMenuStrip _recordMenu;
        private readonly ContextMenuStrip _trayMenu;
        private readonly List<ClipboardEntry> _entries = new List<ClipboardEntry>();
        private readonly string _storageDirectory;
        private readonly string _storageFilePath;
        private bool _suppressNextCapture;

        public ClipboardForm()
        {
            Text = "MyClipboard";
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Width = 320;
            Height = 540;
            TopMost = true;
            BackColor = Color.WhiteSmoke;
            Padding = new Padding(8);

            _recordsView = new BufferedListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.None,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                BackColor = Color.White,
                ForeColor = Color.Black,
                Activation = ItemActivation.Standard
            };
            _recordsView.Columns.Add("内容", 200);
            _recordsView.Columns.Add("格式", 60);
            _recordsView.Columns.Add("时间", 120);
            _recordsView.ItemActivate += (s, e) => PasteSelectedEntry();
            _recordsView.MouseDown += RecordsViewOnMouseDown;
            _recordsView.Resize += (s, e) => AutoSizeColumns();

            Controls.Add(_recordsView);

            _recordMenu = new ContextMenuStrip();
            var copyMenuItem = _recordMenu.Items.Add("复制到剪贴板");
            copyMenuItem.Click += (s, e) => CopySelectedEntry();
            var deleteMenuItem = _recordMenu.Items.Add("删除该记录");
            deleteMenuItem.Click += (s, e) => DeleteSelectedEntry();
            _recordsView.ContextMenuStrip = _recordMenu;

            _trayMenu = new ContextMenuStrip();
            var toggleItem = _trayMenu.Items.Add("显示/隐藏");
            toggleItem.Click += (s, e) => ToggleWindow();
            _trayMenu.Items.Add(new ToolStripSeparator());
            var exitItem = _trayMenu.Items.Add("退出");
            exitItem.Click += (s, e) => Application.Exit();

            _trayIcon = new NotifyIcon
            {
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                Visible = true,
                Text = "MyClipboard",
                ContextMenuStrip = _trayMenu
            };
            _trayIcon.MouseUp += TrayIconOnMouseUp;

            _storageDirectory = ResolveStorageDirectory();
            Directory.CreateDirectory(_storageDirectory);
            _storageFilePath = Path.Combine(_storageDirectory, "records.bin");

            LoadEntries();
            RefreshListView();

            Shown += (s, e) => Hide();
            FormClosing += ClipboardForm_FormClosing;
        }

        private void ClipboardForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                HideForm();
            }
            else
            {
                SaveEntries();
                _trayIcon.Visible = false;
            }
        }

        private void TrayIconOnMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ToggleWindow();
            }
        }

        private static string ResolveStorageDirectory()
        {
            const string windowsPath = @"C:\\ProgramData\\Myclipboard";
            var platform = Environment.OSVersion.Platform;
            if (platform == PlatformID.Win32NT || platform == PlatformID.Win32Windows)
            {
                return windowsPath;
            }

            var fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MyClipboard");
            return fallback;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            AddClipboardFormatListener(Handle);
            RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL | MOD_ALT, (uint)Keys.X);
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            RemoveClipboardFormatListener(Handle);
            UnregisterHotKey(Handle, HOTKEY_ID);
            base.OnHandleDestroyed(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
            {
                OnClipboardChanged();
            }
            else if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                ToggleWindow();
            }

            base.WndProc(ref m);
        }

        private void ToggleWindow()
        {
            if (Visible)
            {
                HideForm();
            }
            else
            {
                ShowFormNearCursor();
            }
        }

        private void HideForm()
        {
            Hide();
        }

        private void ShowFormNearCursor()
        {
            var cursor = Cursor.Position;
            var screen = Screen.FromPoint(cursor);
            int x = cursor.X - Width / 2;
            int y = cursor.Y - Height - 12;

            x = Math.Max(screen.WorkingArea.Left, Math.Min(x, screen.WorkingArea.Right - Width));
            y = Math.Max(screen.WorkingArea.Top, Math.Min(y, screen.WorkingArea.Bottom - Height));

            Location = new Point(x, y);
            Show();
            BringToFront();
            Activate();
        }

        private void RecordsViewOnMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var item = _recordsView.GetItemAt(e.X, e.Y);
            if (item != null)
            {
                item.Selected = true;
            }
        }

        private ClipboardEntry GetSelectedEntry()
        {
            if (_recordsView.SelectedItems.Count == 0)
            {
                return null;
            }

            return _recordsView.SelectedItems[0].Tag as ClipboardEntry;
        }

        private void PasteSelectedEntry()
        {
            var entry = GetSelectedEntry();
            if (entry == null)
            {
                return;
            }

            HideForm();
            var timer = new Timer { Interval = 120 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                ApplyEntryToClipboard(entry, true);
            };
            timer.Start();
        }

        private void CopySelectedEntry()
        {
            var entry = GetSelectedEntry();
            if (entry == null)
            {
                return;
            }

            ApplyEntryToClipboard(entry, false);
        }

        private void DeleteSelectedEntry()
        {
            var entry = GetSelectedEntry();
            if (entry == null)
            {
                return;
            }

            _entries.Remove(entry);
            RefreshListView();
            SaveEntries();
        }

        private void ApplyEntryToClipboard(ClipboardEntry entry, bool simulatePaste)
        {
            if (entry == null || entry.Formats == null || entry.Formats.Count == 0)
            {
                return;
            }

            var dataObject = new DataObject();
            foreach (var payload in entry.Formats)
            {
                var data = payload.ToObject();
                if (data != null)
                {
                    dataObject.SetData(payload.Format, data);
                }
            }

            if (dataObject.GetFormats().Length == 0)
            {
                return;
            }

            ExecuteWithRetry(() =>
            {
                _suppressNextCapture = true;
                Clipboard.SetDataObject(dataObject, true);
            });

            if (simulatePaste)
            {
                var pasteTimer = new Timer { Interval = 60 };
                pasteTimer.Tick += (s, e) =>
                {
                    pasteTimer.Stop();
                    pasteTimer.Dispose();
                    SendPasteKeystroke();
                };
                pasteTimer.Start();
            }
        }

        private void SendPasteKeystroke()
        {
            keybd_event((byte)Keys.ControlKey, 0, 0, 0);
            keybd_event((byte)Keys.V, 0, 0, 0);
            keybd_event((byte)Keys.V, 0, KEYEVENTF_KEYUP, 0);
            keybd_event((byte)Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0);
        }

        private void OnClipboardChanged()
        {
            if (_suppressNextCapture)
            {
                _suppressNextCapture = false;
                return;
            }

            var dataObject = TryGetClipboardData();
            if (dataObject == null)
            {
                return;
            }

            var formats = dataObject.GetFormats(false);
            if (formats == null || formats.Length == 0)
            {
                return;
            }

            var payloads = new List<ClipboardFormatPayload>();
            foreach (var format in formats.Distinct())
            {
                try
                {
                    var data = dataObject.GetData(format, false) ?? dataObject.GetData(format, true);
                    var payload = ClipboardFormatPayload.FromData(format, data);
                    if (payload != null)
                    {
                        payloads.Add(payload);
                    }
                }
                catch
                {
                    // Ignore formats that cannot be captured.
                }
            }

            if (payloads.Count == 0)
            {
                return;
            }

            var entry = new ClipboardEntry
            {
                Id = Guid.NewGuid(),
                Timestamp = DateTime.Now,
                Formats = payloads,
                Preview = ClipboardEntry.BuildPreview(payloads)
            };

            if (_entries.Count > 0 && ClipboardEntry.AreEquivalent(_entries[0], entry))
            {
                return;
            }

            _entries.Insert(0, entry);
            RefreshListView();
            SaveEntries();
        }

        private IDataObject TryGetClipboardData()
        {
            for (var i = 0; i < 4; i++)
            {
                try
                {
                    return Clipboard.GetDataObject();
                }
                catch (ExternalException)
                {
                    Thread.Sleep(40);
                }
            }

            return null;
        }

        private void ExecuteWithRetry(Action action)
        {
            for (var i = 0; i < 4; i++)
            {
                try
                {
                    action();
                    return;
                }
                catch (ExternalException)
                {
                    Thread.Sleep(40);
                }
            }
        }

        private void AutoSizeColumns()
        {
            if (_recordsView.Columns.Count < 3)
            {
                return;
            }

            int width = _recordsView.ClientSize.Width;
            _recordsView.Columns[0].Width = (int)(width * 0.55);
            _recordsView.Columns[1].Width = (int)(width * 0.15);
            _recordsView.Columns[2].Width = width - _recordsView.Columns[0].Width - _recordsView.Columns[1].Width;
        }

        private void LoadEntries()
        {
            if (!File.Exists(_storageFilePath))
            {
                return;
            }

            try
            {
                using (var fs = File.OpenRead(_storageFilePath))
                {
                    var formatter = new BinaryFormatter();
                    if (formatter.Deserialize(fs) is List<ClipboardEntry> stored)
                    {
                        _entries.Clear();
                        _entries.AddRange(stored.OrderByDescending(e => e.Timestamp));
                    }
                }
            }
            catch
            {
                _entries.Clear();
            }
        }

        private void SaveEntries()
        {
            try
            {
                using (var fs = File.Open(_storageFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(fs, _entries);
                }
            }
            catch
            {
                // Ignore persistence errors to avoid crashing the app.
            }
        }

        private void RefreshListView()
        {
            _recordsView.BeginUpdate();
            _recordsView.Items.Clear();

            foreach (var entry in _entries)
            {
                var item = new ListViewItem(entry.Preview ?? "");
                item.SubItems.Add(entry.Formats?.Count.ToString() ?? "0");
                item.SubItems.Add(entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"));
                item.Tag = entry;
                _recordsView.Items.Add(item);
            }

            _recordsView.EndUpdate();
            AutoSizeColumns();
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }

    [Serializable]
    public sealed class ClipboardEntry
    {
        private static readonly byte[] EmptyBytes = new byte[0];

        public Guid Id { get; set; }
        public DateTime Timestamp { get; set; }
        public string Preview { get; set; }
        public List<ClipboardFormatPayload> Formats { get; set; }

        public ClipboardEntry()
        {
            Formats = new List<ClipboardFormatPayload>();
        }

        public static string BuildPreview(IEnumerable<ClipboardFormatPayload> payloads)
        {
            if (payloads == null)
            {
                return string.Empty;
            }

            var textPayload = payloads.FirstOrDefault(p => p.PayloadType == ClipboardPayloadType.Text);
            if (textPayload != null)
            {
                var text = textPayload.Data != null ? Encoding.Unicode.GetString(textPayload.Data) : string.Empty;
                text = NormalizePreview(text);
                if (!string.IsNullOrEmpty(text))
                {
                    return text;
                }
            }

            var first = payloads.FirstOrDefault();
            return first != null ? first.Format : "Clipboard Data";
        }

        private static string NormalizePreview(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var flattened = text.Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (flattened.Length > 120)
            {
                flattened = flattened.Substring(0, 120) + "…";
            }

            return flattened;
        }

        public static bool AreEquivalent(ClipboardEntry a, ClipboardEntry b)
        {
            if (a == null || b == null)
            {
                return false;
            }

            if (a.Formats.Count != b.Formats.Count)
            {
                return false;
            }

            for (int i = 0; i < a.Formats.Count; i++)
            {
                var left = a.Formats[i];
                var right = b.Formats[i];
                if (!string.Equals(left.Format, right.Format, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (left.PayloadType != right.PayloadType)
                {
                    return false;
                }

                var leftBytes = left.Data ?? EmptyBytes;
                var rightBytes = right.Data ?? EmptyBytes;
                if (!leftBytes.SequenceEqual(rightBytes))
                {
                    return false;
                }
            }

            return true;
        }
    }

    public enum ClipboardPayloadType
    {
        Text,
        Binary,
        Image,
        Serialized
    }

    [Serializable]
    public sealed class ClipboardFormatPayload
    {
        public string Format { get; set; }
        public ClipboardPayloadType PayloadType { get; set; }
        public byte[] Data { get; set; }

        public static ClipboardFormatPayload FromData(string format, object data)
        {
            if (string.IsNullOrWhiteSpace(format) || data == null)
            {
                return null;
            }

            if (data is string str)
            {
                return new ClipboardFormatPayload
                {
                    Format = format,
                    PayloadType = ClipboardPayloadType.Text,
                    Data = Encoding.Unicode.GetBytes(str)
                };
            }

            if (data is string[] lines)
            {
                var joined = string.Join(Environment.NewLine, lines);
                return new ClipboardFormatPayload
                {
                    Format = format,
                    PayloadType = ClipboardPayloadType.Text,
                    Data = Encoding.Unicode.GetBytes(joined)
                };
            }

            if (data is Bitmap bitmap)
            {
                using (var ms = new MemoryStream())
                {
                    bitmap.Save(ms, ImageFormat.Png);
                    return new ClipboardFormatPayload
                    {
                        Format = format,
                        PayloadType = ClipboardPayloadType.Image,
                        Data = ms.ToArray()
                    };
                }
            }

            if (data is Image image)
            {
                using (var ms = new MemoryStream())
                {
                    image.Save(ms, ImageFormat.Png);
                    return new ClipboardFormatPayload
                    {
                        Format = format,
                        PayloadType = ClipboardPayloadType.Image,
                        Data = ms.ToArray()
                    };
                }
            }

            if (data is MemoryStream memoryStream)
            {
                return new ClipboardFormatPayload
                {
                    Format = format,
                    PayloadType = ClipboardPayloadType.Binary,
                    Data = memoryStream.ToArray()
                };
            }

            if (data is byte[] bytes)
            {
                return new ClipboardFormatPayload
                {
                    Format = format,
                    PayloadType = ClipboardPayloadType.Binary,
                    Data = bytes.ToArray()
                };
            }

            try
            {
                using (var ms = new MemoryStream())
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(ms, data);
                    return new ClipboardFormatPayload
                    {
                        Format = format,
                        PayloadType = ClipboardPayloadType.Serialized,
                        Data = ms.ToArray()
                    };
                }
            }
            catch
            {
                return null;
            }
        }

        public object ToObject()
        {
            if (Data == null || Data.Length == 0)
            {
                return null;
            }

            switch (PayloadType)
            {
                case ClipboardPayloadType.Text:
                    return Encoding.Unicode.GetString(Data);
                case ClipboardPayloadType.Image:
                    using (var ms = new MemoryStream(Data))
                    using (var image = Image.FromStream(ms))
                    {
                        return (Image)image.Clone();
                    }
                case ClipboardPayloadType.Binary:
                    return new MemoryStream(Data, writable: false);
                case ClipboardPayloadType.Serialized:
                    try
                    {
                        using (var ms = new MemoryStream(Data))
                        {
                            var formatter = new BinaryFormatter();
                            return formatter.Deserialize(ms);
                        }
                    }
                    catch
                    {
                        return null;
                    }
                default:
                    return null;
            }
        }
    }

    internal sealed class BufferedListView : ListView
    {
        public BufferedListView()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            UpdateStyles();
        }
    }
}
