using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;

namespace desktopcapt
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

        [DllImport("user32.dll")]
        private static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetShellWindow();

        // NEW: DPI awareness imports
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private const int SPI_SETDESKWALLPAPER = 0x0014;
        private const int SPIF_UPDATEINIFILE = 0x01;
        private const int SPIF_SENDCHANGE = 0x02;
        private const uint PW_RENDERFULLCONTENT = 0x00000002;
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const uint GW_HWNDNEXT = 2;

        // NEW: DPI constants
        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;

        private PictureBox pictureBox;
        private Button btnCapture;
        private Button btnSave;
        private Bitmap currentScreenshot;
        private CheckBox chkTransparent;
        private string originalWallpaper;
        private string tempBlackPath;
        private string tempWhitePath;

        public Form1()
        {
            // NEW: Set DPI awareness before anything else
            SetDPIAwareness();

            InitializeComponent();
            SetupUI();
            CreateTempWallpapers();
        }

        // NEW: Method to set DPI awareness
        private void SetDPIAwareness()
        {
            try
            {
                // Try Windows 8.1+ method first
                SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
            }
            catch
            {
                try
                {
                    // Fall back to Windows Vista+ method
                    SetProcessDPIAware();
                }
                catch { }
            }
        }

        // NEW: Get actual physical screen dimensions (not scaled)
        private void GetActualScreenDimensions(out int width, out int height)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            width = GetDeviceCaps(hdc, LOGPIXELSX);
            height = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(IntPtr.Zero, hdc);

            // Get the actual resolution
            width = Screen.PrimaryScreen.Bounds.Width;
            height = Screen.PrimaryScreen.Bounds.Height;

            // Apply DPI scaling factor
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = g.DpiX / 96f; // 96 is the default DPI
                float dpiY = g.DpiY / 96f;

                width = (int)(width * dpiX);
                height = (int)(height * dpiY);
            }
        }

        private void SetupUI()
        {
            this.Text = "DesktopCapt";
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Capture button
            btnCapture = new Button
            {
                Text = "Capture Desktop Icons",
                Location = new Point(20, 20),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnCapture.Click += BtnCapture_Click;

            // Save button
            btnSave = new Button
            {
                Text = "Save image...",
                Location = new Point(230, 20),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Enabled = false
            };
            btnSave.Click += BtnSave_Click;

            // Transparent background checkbox
            chkTransparent = new CheckBox
            {
                Text = "Transparent Background (slower)",
                Location = new Point(440, 25),
                Size = new Size(220, 30),
                Font = new Font("Segoe UI", 9),
                Checked = true
            };

            // Picture box for preview
            pictureBox = new PictureBox
            {
                Location = new Point(20, 70),
                Size = new Size(840, 580),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(255, 64, 64, 64)
            };

            this.Controls.Add(btnCapture);
            this.Controls.Add(btnSave);
            this.Controls.Add(chkTransparent);
            this.Controls.Add(pictureBox);
        }

        private void CreateTempWallpapers()
        {
            try
            {
                string tempFolder = Path.GetTempPath();
                tempBlackPath = Path.Combine(tempFolder, "temp_black_wallpaper.bmp");
                tempWhitePath = Path.Combine(tempFolder, "temp_white_wallpaper.bmp");

                // UPDATED: Get actual DPI-aware screen resolution
                GetActualScreenDimensions(out int screenWidth, out int screenHeight);

                // Create solid black wallpaper
                using (Bitmap blackBmp = new Bitmap(screenWidth, screenHeight))
                using (Graphics g = Graphics.FromImage(blackBmp))
                {
                    g.Clear(Color.Black);
                    blackBmp.Save(tempBlackPath, ImageFormat.Bmp);
                }

                // Create solid white wallpaper
                using (Bitmap whiteBmp = new Bitmap(screenWidth, screenHeight))
                using (Graphics g = Graphics.FromImage(whiteBmp))
                {
                    g.Clear(Color.White);
                    whiteBmp.Save(tempWhitePath, ImageFormat.Bmp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating temp wallpapers: {ex.Message}", "Error");
            }
        }

        private string GetCurrentWallpaper()
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", false);
                return key?.GetValue("Wallpaper")?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private void SetWallpaper(string path)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);

            // Force a full desktop refresh to make sure wallpaper updates behind icons
            System.Threading.Thread.Sleep(300);

            IntPtr progman = FindWindow("Progman", null);
            if (progman != IntPtr.Zero)
            {
                InvalidateRect(progman, IntPtr.Zero, true);
                UpdateWindow(progman);
            }

            // Also refresh all WorkerW windows
            IntPtr tempWorkerW = IntPtr.Zero;
            do
            {
                tempWorkerW = FindWindowEx(IntPtr.Zero, tempWorkerW, "WorkerW", null);
                if (tempWorkerW != IntPtr.Zero)
                {
                    InvalidateRect(tempWorkerW, IntPtr.Zero, true);
                    UpdateWindow(tempWorkerW);
                }
            } while (tempWorkerW != IntPtr.Zero);

            System.Threading.Thread.Sleep(200); // Give Windows time to redraw
        }

        private IntPtr FindDesktopIconWindow()
        {
            IntPtr progman = FindWindow("Progman", null);
            IntPtr workerW = IntPtr.Zero;

            IntPtr tempWorkerW = IntPtr.Zero;
            do
            {
                tempWorkerW = FindWindowEx(IntPtr.Zero, tempWorkerW, "WorkerW", null);
                if (tempWorkerW != IntPtr.Zero)
                {
                    IntPtr shellView = FindWindowEx(tempWorkerW, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (shellView != IntPtr.Zero)
                    {
                        workerW = tempWorkerW;
                        break;
                    }
                }
            } while (tempWorkerW != IntPtr.Zero);

            return workerW != IntPtr.Zero ? workerW : progman;
        }

        private Bitmap CaptureProgman()
        {
            // UPDATED: Get DPI-aware screen dimensions
            GetActualScreenDimensions(out int screenWidth, out int screenHeight);

            Bitmap screenshot = new Bitmap(screenWidth, screenHeight, PixelFormat.Format32bppArgb);

            using (Graphics g = Graphics.FromImage(screenshot))
            {
                // Capture directly from screen with proper DPI handling
                g.CopyFromScreen(0, 0, 0, 0, new Size(screenWidth, screenHeight), CopyPixelOperation.SourceCopy);
            }

            return screenshot;
        }

        private List<IntPtr> HideAllWindows()
        {
            List<IntPtr> hiddenWindows = new List<IntPtr>();
            IntPtr shellWindow = GetShellWindow();

            // Hide taskbar first
            IntPtr taskbar = FindWindow("Shell_TrayWnd", null);
            if (taskbar != IntPtr.Zero)
            {
                ShowWindow(taskbar, SW_HIDE);
                hiddenWindows.Add(taskbar);
            }

            // Hide start button specifically (Windows 7)
            IntPtr startButton = FindWindowEx(taskbar, IntPtr.Zero, "Button", null);
            if (startButton != IntPtr.Zero)
            {
                ShowWindow(startButton, SW_HIDE);
                hiddenWindows.Add(startButton);
            }

            // Hide secondary taskbars on multi-monitor setups
            IntPtr secondaryTaskbar = IntPtr.Zero;
            do
            {
                secondaryTaskbar = FindWindowEx(IntPtr.Zero, secondaryTaskbar, "Shell_SecondaryTrayWnd", null);
                if (secondaryTaskbar != IntPtr.Zero)
                {
                    ShowWindow(secondaryTaskbar, SW_HIDE);
                    hiddenWindows.Add(secondaryTaskbar);
                }
            } while (secondaryTaskbar != IntPtr.Zero);

            // Use EnumWindows to catch ALL visible windows
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd) && hWnd != shellWindow && hWnd != this.Handle)
                {
                    // Skip desktop windows
                    if (hWnd != FindWindow("Progman", null))
                    {
                        IntPtr tempWorkerW = IntPtr.Zero;
                        bool isWorkerW = false;
                        do
                        {
                            tempWorkerW = FindWindowEx(IntPtr.Zero, tempWorkerW, "WorkerW", null);
                            if (tempWorkerW == hWnd)
                            {
                                isWorkerW = true;
                                break;
                            }
                        } while (tempWorkerW != IntPtr.Zero);

                        if (!isWorkerW)
                        {
                            ShowWindow(hWnd, SW_HIDE);
                            hiddenWindows.Add(hWnd);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            System.Threading.Thread.Sleep(1000); // Give everything time to hide
            return hiddenWindows;
        }

        private void ShowAllWindows(List<IntPtr> windows)
        {
            foreach (IntPtr hWnd in windows)
            {
                ShowWindow(hWnd, SW_SHOW);
            }
        }

        private static byte ToByte(int i)
        {
            return (byte)(i > 255 ? 255 : (i < 0 ? 0 : i));
        }

        private Bitmap ExtractIcons(Bitmap whiteCapture, Bitmap blackCapture)
        {
            int width = blackCapture.Width;
            int height = blackCapture.Height;
            Bitmap result = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            // Lock bits for faster processing
            BitmapData whiteData = whiteCapture.LockBits(new Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData blackData = blackCapture.LockBits(new Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData resultData = result.LockBits(new Rectangle(0, 0, width, height),
                ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* whitePtr = (byte*)whiteData.Scan0;
                byte* blackPtr = (byte*)blackData.Scan0;
                byte* resultPtr = (byte*)resultData.Scan0;

                int bytes = Math.Abs(blackData.Stride) * height;

                for (int i = 0; i < bytes; i += 4)
                {
                    byte whiteB = whitePtr[i];
                    byte whiteG = whitePtr[i + 1];
                    byte whiteR = whitePtr[i + 2];

                    byte blackB = blackPtr[i];
                    byte blackG = blackPtr[i + 1];
                    byte blackR = blackPtr[i + 2];

                    // Calculate alpha using Cvolton's algorithm
                    int alpha = ToByte((blackR - whiteR + 255 + blackG - whiteG + 255 + blackB - whiteB + 255) / 3);

                    if (alpha > 0)
                    {
                        // Black background optimization algorithm
                        resultPtr[i + 2] = ToByte(255 * blackR / alpha); // Red
                        resultPtr[i + 1] = ToByte(255 * blackG / alpha); // Green
                        resultPtr[i] = ToByte(255 * blackB / alpha);     // Blue
                        resultPtr[i + 3] = (byte)alpha;                   // Alpha
                    }
                    else
                    {
                        resultPtr[i] = 0;
                        resultPtr[i + 1] = 0;
                        resultPtr[i + 2] = 0;
                        resultPtr[i + 3] = 0; // Fully transparent
                    }
                }
            }

            whiteCapture.UnlockBits(whiteData);
            blackCapture.UnlockBits(blackData);
            result.UnlockBits(resultData);

            return result;
        }

        private void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                if (chkTransparent.Checked)
                {
                    // Save original wallpaper
                    originalWallpaper = GetCurrentWallpaper();

                    this.Cursor = Cursors.WaitCursor;
                    btnCapture.Enabled = false;
                    btnCapture.Text = "Hiding windows...";
                    Application.DoEvents();

                    // Hide all windows for clean capture
                    this.WindowState = FormWindowState.Minimized;
                    Application.DoEvents();
                    List<IntPtr> hiddenWindows = HideAllWindows();

                    btnCapture.Text = "Capturing black...";
                    Application.DoEvents();

                    // Capture with black background
                    SetWallpaper(tempBlackPath);
                    Bitmap blackCapture = CaptureProgman();

                    btnCapture.Text = "Capturing white...";
                    Application.DoEvents();

                    // Capture with white background
                    SetWallpaper(tempWhitePath);
                    Bitmap whiteCapture = CaptureProgman();

                    btnCapture.Text = "Restoring wallpaper...";
                    Application.DoEvents();

                    // Restore original wallpaper
                    if (!string.IsNullOrEmpty(originalWallpaper))
                    {
                        SetWallpaper(originalWallpaper);
                    }

                    // Show all windows again
                    ShowAllWindows(hiddenWindows);
                    this.WindowState = FormWindowState.Normal;
                    Application.DoEvents();

                    if (blackCapture == null || whiteCapture == null)
                    {
                        MessageBox.Show("Failed to capture desktop!", "Error");
                        return;
                    }

                    btnCapture.Text = "Processing transparency...";
                    Application.DoEvents();

                    // Extract icons with transparency
                    if (currentScreenshot != null)
                        currentScreenshot.Dispose();

                    currentScreenshot = ExtractIcons(whiteCapture, blackCapture);

                    blackCapture.Dispose();
                    whiteCapture.Dispose();

                    pictureBox.Image = currentScreenshot;
                    btnSave.Enabled = true;

                    MessageBox.Show($"Desktop icons captured with transparency!\nSize: {currentScreenshot.Width}x{currentScreenshot.Height}px",
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Simple capture without transparency - hide windows for clean capture
                    this.Cursor = Cursors.WaitCursor;
                    this.WindowState = FormWindowState.Minimized;
                    Application.DoEvents();

                    List<IntPtr> hiddenWindows = HideAllWindows();

                    if (currentScreenshot != null)
                        currentScreenshot.Dispose();

                    currentScreenshot = CaptureProgman();

                    ShowAllWindows(hiddenWindows);
                    this.WindowState = FormWindowState.Normal;
                    Application.DoEvents();

                    if (currentScreenshot == null)
                    {
                        MessageBox.Show("Could not capture desktop!", "Error");
                        return;
                    }

                    pictureBox.Image = currentScreenshot;
                    btnSave.Enabled = true;

                    MessageBox.Show($"Desktop icons captured!\nSize: {currentScreenshot.Width}x{currentScreenshot.Height}px",
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error capturing screenshot: {ex.Message}", "Error");

                // Try to restore wallpaper on error
                if (!string.IsNullOrEmpty(originalWallpaper))
                {
                    try { SetWallpaper(originalWallpaper); } catch { }
                }
            }
            finally
            {
                this.Cursor = Cursors.Default;
                btnCapture.Enabled = true;
                btnCapture.Text = "Capture Desktop Icons";
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (currentScreenshot == null)
            {
                MessageBox.Show("No screenshot to save!", "Error");
                return;
            }

            try
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Title = "Save Desktop Screenshot";
                    saveDialog.Filter = "PNG Image|*.png|JPEG Image|*.jpg|Bitmap Image|*.bmp|All Files|*.*";
                    saveDialog.FilterIndex = 1;
                    saveDialog.DefaultExt = "png";
                    saveDialog.FileName = $"Desktop_Icons_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.png";
                    saveDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    if (saveDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(saveDialog.FileName))
                    {
                        // Determine format based on file extension
                        ImageFormat format = ImageFormat.Png;
                        string ext = Path.GetExtension(saveDialog.FileName).ToLower();

                        switch (ext)
                        {
                            case ".jpg":
                            case ".jpeg":
                                format = ImageFormat.Jpeg;
                                break;
                            case ".bmp":
                                format = ImageFormat.Bmp;
                                break;
                            default:
                                format = ImageFormat.Png;
                                break;
                        }

                        currentScreenshot.Save(saveDialog.FileName, format);

                        MessageBox.Show($"Screenshot saved!\n{Path.GetFileName(saveDialog.FileName)}",
                            "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving screenshot: {ex.Message}", "Error");
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (File.Exists(tempBlackPath)) File.Delete(tempBlackPath);
                if (File.Exists(tempWhitePath)) File.Delete(tempWhitePath);
            }
            catch { }

            base.OnFormClosing(e);
        }
    }
}