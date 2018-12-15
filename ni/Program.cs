using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Globalization;
using Microsoft.Win32;
using System.Text;
using System.Windows.Forms.VisualStyles;
 
namespace ni
{
    static class Program
    {


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {                

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new appcontext());
            }
            catch (Exception ex)
            {
                Debug.WriteLine(System.Reflection.MethodBase.GetCurrentMethod().Name + " " + ex.Message);
            }
        }



    }


    #region appcontextCLS
    class appcontext : ApplicationContext
    {
        #region NumLock Always On
        //https://bytes.com/topic/c-sharp/answers/464602-possible-turn-off-caps_lock
        //https://msdn.microsoft.com/pt-pt/library/windows/desktop/ms646310(v=vs.85).aspx
        //https://www.codeproject.com/Articles/26317/Read-and-Update-CAPS-or-NUM-Lock-Status-from-your
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")]
        internal static extern short GetKeyState(int keyCode);

        const int KEYEVENTF_EXTENDEDKEY = 0x1;
        const int KEYEVENTF_KEYUP = 0x2;
        const int CAPSLOCK = 0x14;
        const int NUMLOCK = 0x90;
        const int SCROLLLOCK = 0x91;
        #endregion

        #region Variables
        System.Windows.Forms.Timer t = null;
        ToolStripMenuItem openAfterCaptureToolStripMenuItem = null;
        ToolStripMenuItem keepNumLockAlwaysOnToolStripMenuItem = null;
        ToolStripMenuItem openFolderToolStripMenuItem = null;
        ToolStripMenuItem defineFolderToolStripMenuItem = null;
        ToolStripMenuItem defineSecondsBetweenCaptureToolStripMenuItem = null;
        ToolStripMenuItem openPictureToolStripMenuItem = null;
        ToolStripMenuItem CaptureEvery_x_SecToolStripMenuItem = null;
        ToolStripMenuItem CaptureToolStripMenuItem = null;
        NotifyIcon ni = null;

        /// <summary>
        /// The name of the last picture taken
        /// </summary>
        string PictureName = "";
        UInt32 counter = 0;
        #endregion

        public appcontext()
        {
            try
            {

                if (InstanceIsAlreadyRunning() != null)
                {
                    MessageBox.Show("Capture Screen already running.\nCheck your notification area.");
                    Environment.Exit(0);
                }

                ThreadPool.QueueUserWorkItem(SwitchToClassicTheme, true);

                //Init Timer every 1 sec
                t = new System.Windows.Forms.Timer();
                t.Interval = 1000;
                t.Tick += T_Tick;
                t.Enabled = true;


                //Init NotifyIcon
                ni = new NotifyIcon();
                ni.Icon = Properties.Resources.Red_camera;
                ni.ContextMenuStrip = new ContextMenuStrip();
                ni.MouseClick += (s, e) =>
                {
                    try
                    {
                        if (e.Button == MouseButtons.Left)
                        {
                            if (!CaptureEvery_x_SecToolStripMenuItem.Checked)
                            {
                                Debug.WriteLine(DateTime.Now.ToString() + " Capture from Icon");
                                CaptureScreen();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("ni.MouseClick: " + ex.Message);
                    }
                };


                //Context Menu displays each item as a button like a menu, not like a strip
                //ni.ContextMenu = new ContextMenu();
                //ni.ContextMenu.MenuItems.Add("1");
                //ni.ContextMenu.MenuItems.Add("2");
                //ni.ContextMenu.MenuItems.Add("3");
                //ni.ContextMenu.MenuItems.Add("Exit", (s, e) => Application.Exit());

                #region Context Menus


                ni.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Rev " + Assembly.GetEntryAssembly().GetName().Version, null));
                //Separator
                ni.ContextMenuStrip.Items.Add("-");

                //Define Folder
                defineFolderToolStripMenuItem = new ToolStripMenuItem("Define Folder", null,
                    (s, e) =>
                    {
                        try
                        {
                            CheckTargetFolderExists(openFolderToolStripMenuItem, ni);
                            FolderSelectDialog dialog = new FolderSelectDialog
                            {
                                InitialDirectory = Properties.Settings.Default.TargetFolder,
                                Title = "Select a folder to save your screen captures"
                            };

                            if (dialog.Show())
                            {
                                Properties.Settings.Default.TargetFolder = dialog.FileName;
                                Properties.Settings.Default.Save();
                                CheckTargetFolderExists(openFolderToolStripMenuItem, ni);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("defineFolderToolStripMenuItem: " + ex.Message);
                        }
                    });
                ni.ContextMenuStrip.Items.Add(defineFolderToolStripMenuItem);


                




                //Open Folder
                openFolderToolStripMenuItem = new ToolStripMenuItem("Will be updated below", //see CheckTargetFolderExists
                    null,
                    (s, e) =>
                    {
                        try
                        {
                            System.Diagnostics.Process.Start("explorer.exe", Properties.Settings.Default.TargetFolder);
                        }
                        catch (Exception ex)
                        {

                            Debug.WriteLine("openFolderToolStripMenuItem: " + ex.Message);
                        }
                    });

                ni.ContextMenuStrip.Items.Add(openFolderToolStripMenuItem);


                //Separator
                ni.ContextMenuStrip.Items.Add("-");
                // Keep NumLock On.
                keepNumLockAlwaysOnToolStripMenuItem = new ToolStripMenuItem("Keep NumLock Always On"); //Triggers no event
                keepNumLockAlwaysOnToolStripMenuItem.CheckOnClick = true;
                ni.ContextMenuStrip.Items.Add(keepNumLockAlwaysOnToolStripMenuItem);
                keepNumLockAlwaysOnToolStripMenuItem.Checked = Properties.Settings.Default.KeepNumLockAlwaysOn; //Read from settings

                //Separator
                ni.ContextMenuStrip.Items.Add("-");

                // Open After Capture.
                openAfterCaptureToolStripMenuItem = new ToolStripMenuItem("Open After Capture"); //Triggers no event
                openAfterCaptureToolStripMenuItem.CheckOnClick = true;
                ni.ContextMenuStrip.Items.Add(openAfterCaptureToolStripMenuItem);
                openAfterCaptureToolStripMenuItem.Checked = true;
                //Open Last Picture
                openPictureToolStripMenuItem = new ToolStripMenuItem("Open Last Picture", null,
                    (s, e) =>
                    {
                        try
                        {
                            if (!PictureName.Equals(""))
                            {
                                Process.Start(PictureName);
                            }
                        }
                        catch (Exception ex)
                        {

                            Debug.WriteLine("openPictureToolStripMenuItem: " + ex.Message);
                        }
                    });
                openPictureToolStripMenuItem.Enabled = false; //Only available when PictureName != ""
                ni.ContextMenuStrip.Items.Add(openPictureToolStripMenuItem);
                
                //Separator
                ni.ContextMenuStrip.Items.Add("-");
                //Capture Every x sec Check
                CaptureEvery_x_SecToolStripMenuItem = new ToolStripMenuItem("Capture Every "+ Properties.Settings.Default.SecondsBetweenCapture.ToString()+ " Sec", null,
                    (s, e) =>
                    {
                        try
                        {
                            if (CaptureEvery_x_SecToolStripMenuItem.Checked)
                            {
                                CaptureToolStripMenuItem.Enabled = false;
                            }
                            else
                            {
                                CaptureToolStripMenuItem.Enabled = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("CaptureEvery_x_SecToolStripMenuItem: " + ex.Message);
                        }
                    });
                CaptureEvery_x_SecToolStripMenuItem.CheckOnClick = true;
                ni.ContextMenuStrip.Items.Add(CaptureEvery_x_SecToolStripMenuItem); //Triggers no event

                //Define Seconds Between Capture
                defineSecondsBetweenCaptureToolStripMenuItem = new ToolStripMenuItem("Define Seconds Between Capture", null,
                    (s, e) =>
                    {
                        try
                        {
                            frmSecondsBetweenCapture f = 
                                new frmSecondsBetweenCapture(Properties.Settings.Default.SecondsBetweenCapture);
                            
                            if (f.ShowDialog() == DialogResult.OK)
                            {
                                //Convert to uint
                                if (uint.TryParse(f.tbSecondsBetweenCapture.Text,
                                    out uint result))
                                {
                                    Properties.Settings.Default.SecondsBetweenCapture = result;
                                    Properties.Settings.Default.Save();

                                    //Change text
                                    CaptureEvery_x_SecToolStripMenuItem.Text =
                                        "Capture Every " + Properties.Settings.Default.SecondsBetweenCapture.ToString() + " Sec";
                                }
                            }
                            f.Close();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("defineSecondsBetweenCaptureToolStripMenuItem: " + ex.Message);
                        }
                    });
                ni.ContextMenuStrip.Items.Add(defineSecondsBetweenCaptureToolStripMenuItem);
                
                //Separator
                ni.ContextMenuStrip.Items.Add("-");

                //Capture
                CaptureToolStripMenuItem = new ToolStripMenuItem("Capture", null,
                    (s, e) =>
                    {
                        try
                        {
                            if (!CaptureEvery_x_SecToolStripMenuItem.Checked)
                            {
                                Debug.WriteLine(DateTime.Now.ToString() + " Capture from menu");
                                CaptureScreen();
                            }
                        }
                        catch (Exception ex)
                        {

                            Debug.WriteLine("CaptureToolStripMenuItem: " + ex.Message);
                        }
                    });
                ni.ContextMenuStrip.Items.Add(CaptureToolStripMenuItem);

                //Separator
                ni.ContextMenuStrip.Items.Add("-");

                //Exit                      
                ni.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Exit", Properties.Resources.Exit,
                        (s, e) =>
                        {
                            Debug.WriteLine(DateTime.Now.ToString() + " " + ((ToolStripMenuItem)s).Text);
                            t.Enabled = false;
                            ni.Visible = false;
                            Environment.Exit(0);
                        }));

                #endregion

                CheckTargetFolderExists(openFolderToolStripMenuItem, ni);

                ni.Visible = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("appcontext(): " + ex.Message);
            }
        }//appcontext

        #region Functions

        /// <summary>
        /// This thread will set Classic theme and vanish
        /// </summary>
        /// <param name="CheckRegistry"></param>
        void SwitchToClassicTheme(Object CheckRegistry)
        {
            //if any theme is active, switch to classic theme 
            //string ThemeIsActive = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive", "1");

            //if (ThemeIsActive.Equals("1"))
            //{
            //   bool b = Theming.SwitchToClassicTheme();
            //}

            //Change display background to a dark blue
            WallpaperColorChanger.SetColor(Color.FromArgb(0, 75, 100));
        }

        #region Set BackGround Color
        /// <summary>
        /// See https://stackoverflow.com/questions/7309943/c-set-desktop-wallpaper-to-a-solid-color
        /// </summary>
        public static class WallpaperColorChanger
        {

            public static void SetColor(Color color)
            {

                // Remove the current wallpaper
                NativeMethods.SystemParametersInfo(
                    NativeMethods.SPI_SETDESKWALLPAPER,
                    0,
                    "",
                    NativeMethods.SPIF_UPDATEINIFILE | NativeMethods.SPIF_SENDWININICHANGE);

                // Set the new desktop solid color for the current session
                int[] elements = { NativeMethods.COLOR_DESKTOP };
                int[] colors = { System.Drawing.ColorTranslator.ToWin32(color) };
                NativeMethods.SetSysColors(elements.Length, elements, colors);

                // Save value in registry so that it will persist
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Colors", true);
                key.SetValue(@"Background", string.Format("{0} {1} {2}", color.R, color.G, color.B));
            }

            private static class NativeMethods
            {
                public const int COLOR_DESKTOP = 1;
                public const int SPI_SETDESKWALLPAPER = 20;
                public const int SPIF_UPDATEINIFILE = 0x01;
                public const int SPIF_SENDWININICHANGE = 0x02;

                [DllImport("user32.dll")]
                public static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);

                [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
                public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
            }
        }
        #endregion

        /// <summary>
        /// If Properties.Settings.Default.TargetFolder exists, do nothing. 
        /// If not set Properties.Settings.Default.TargetFolder to MyPictures Folder. 
        /// 
        /// set openFolderToolStripMenuItem Text to the Properties.Settings.Default.TargetFolder
        /// </summary>
        void CheckTargetFolderExists(ToolStripMenuItem openFolderToolStripMenuItem, NotifyIcon ni)
        {
            try
            {
                if (!Directory.Exists(Properties.Settings.Default.TargetFolder))
                {
                    Properties.Settings.Default.TargetFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                    Properties.Settings.Default.Save();
                }

                if (openFolderToolStripMenuItem != null)
                {
                    openFolderToolStripMenuItem.Text = "Open Folder (" + Properties.Settings.Default.TargetFolder + ")";
                }

                if (ni != null)
                {
                    ni.Text = "Capture Screen to Folder " + Properties.Settings.Default.TargetFolder;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(System.Reflection.MethodBase.GetCurrentMethod().Name + " " + ex.Message);
            }

        }

        string GetFullPathPictureName()
        {
            CheckTargetFolderExists(openFolderToolStripMenuItem, ni);
            string t = Properties.Settings.Default.TargetFolder;
            DateTime d = DateTime.Now;
            if (!t.EndsWith(@"\"))
            {
                t += @"\";
            }
            return t +
                d.Hour.ToString("00") + "h" +
                d.Minute.ToString("00") + "m" +
                d.Second.ToString("00") + "s " +
                d.Year.ToString("0000") + "-" +
                d.Month.ToString("00") + "-" +
                d.Day.ToString("00") + ".jpg";

        }

        void CaptureScreen()
        {

            //Change icon to green
            ni.Icon = Properties.Resources.Green_camera;

            try
            {

                Image image = new ScreenCapture().CaptureDesktop();

                PictureName = GetFullPathPictureName();

                //Enable OpenPicture because PictureName is not empty anymore
                openPictureToolStripMenuItem.Enabled = true;

                image.Save(PictureName, ImageFormat.Jpeg);

                if (openAfterCaptureToolStripMenuItem.Checked)
                {
                    Process.Start(PictureName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            //Change icon to red
            ni.Icon = Properties.Resources.Red_camera;

        }

        /// <summary>
        /// Turn on NumLock
        /// </summary>
        void ManageKeyboardKeys()
        {
            try
            {
                if (GetKeyState((int)Keys.NumLock) == 0)
                {
                    const int KEYEVENTF_EXTENDEDKEY = 0x1;
                    const int KEYEVENTF_KEYUP = 0x2;
                    keybd_event(0x90, 0x45, KEYEVENTF_EXTENDEDKEY, (UIntPtr)0);
                    keybd_event(0x90, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, (UIntPtr)0);
                }


                //If CapsLock is On, camera turns to yellow
                if (GetKeyState((int)Keys.CapsLock) == 0)
                {
                    ni.Icon = Properties.Resources.Red_camera;
                }
                else
                {
                    ni.Icon = Properties.Resources.camera_yellow;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(System.Reflection.MethodBase.GetCurrentMethod().Name + " " + ex.Message);
            }
        }

        /// <summary>
        /// Every 1 sec
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void T_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Properties.Settings.Default.KeepNumLockAlwaysOn != keepNumLockAlwaysOnToolStripMenuItem.Checked)
                {
                    Properties.Settings.Default.KeepNumLockAlwaysOn = keepNumLockAlwaysOnToolStripMenuItem.Checked;
                    Properties.Settings.Default.Save();
                    Debug.WriteLine("SAVED " + Properties.Settings.Default.KeepNumLockAlwaysOn.ToString());
                }
                 
                if (keepNumLockAlwaysOnToolStripMenuItem.Checked)
                {
                    Debug.WriteLine(Properties.Settings.Default.KeepNumLockAlwaysOn);
                    ManageKeyboardKeys();
                }
                

                if (CaptureEvery_x_SecToolStripMenuItem.Checked)
                {
                    if (counter <= 0)
                    {
                        CaptureScreen();
                        counter = Properties.Settings.Default.SecondsBetweenCapture-1; //perform wait according to property sec
                    }
                    else
                    {
                        counter--;
                    }
                }
                else
                {
                    counter = 0;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(System.Reflection.MethodBase.GetCurrentMethod().Name + " " + ex.Message);
            }
        }

        /// <summary>
        /// make sure only one process is running
        /// </summary>
        /// <returns></returns>
        private Process InstanceIsAlreadyRunning()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            //Loop through the running processes in with the same name 
            foreach (Process process in processes)
            {
                //Ignore the current process 
                if (process.Id != current.Id)
                {
                    //Make sure that the process is running from the exe file. 
                    if (Assembly.GetExecutingAssembly().Location.
                         Replace("/", "\\") == current.MainModule.FileName)
                    {
                        //Return the other process instance.  
                        return process;
                    }
                }
            }
            //No other instance was found, return null.  
            return null;
        }
        #endregion Functions
    }

    #endregion appcontextCLS

    #region FolderSelectDialogCLS
    /// <summary>
    /// https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp
    /// ability to copy and paste from a textbox at the bottom and the navigation pane on the left with favorites and common locations
    /// </summary>
    public class FolderSelectDialog
    {
        //[DllImport("dwmapi.dll")]
        //public static extern IntPtr DwmIsCompositionEnabled(out bool pfEnabled);

        private string _initialDirectory;
        private string _title;
        private string _fileName = "";
        /// <summary>
        /// 
        /// </summary>
        public string InitialDirectory
        {
            get { return string.IsNullOrEmpty(_initialDirectory) ? Environment.CurrentDirectory : _initialDirectory; }
            set { _initialDirectory = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Title
        {
            get { return _title ?? "Select a folder"; }
            set { _title = value; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FileName
        {
            get
            {
                return _fileName;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool Show()
        {
            return Show(IntPtr.Zero);
        }

        /// <param name="hWndOwner">Handle of the control or window to be the parent of the file dialog</param>
        /// <returns>true if the user clicks OK</returns>
        public bool Show(IntPtr hWndOwner)
        {
            var result = Environment.OSVersion.Version.Major >= 6
                ? VistaDialog.Show(hWndOwner, InitialDirectory, Title)
                : ShowXpDialog(hWndOwner, InitialDirectory, Title);
            _fileName = result.FileName;
            return result.Result;
        }

        private struct ShowDialogResult
        {
            public bool Result { get; set; }
            public string FileName { get; set; }
        }

        private static ShowDialogResult ShowXpDialog(IntPtr ownerHandle, string initialDirectory, string title)
        {
            var folderBrowserDialog = new FolderBrowserDialog
            {
                Description = title,
                SelectedPath = initialDirectory,
                ShowNewFolderButton = false
            };
            var dialogResult = new ShowDialogResult();
            if (folderBrowserDialog.ShowDialog(new WindowWrapper(ownerHandle)) == DialogResult.OK)
            {
                dialogResult.Result = true;
                dialogResult.FileName = folderBrowserDialog.SelectedPath;
            }
            return dialogResult;
        }

        private static class VistaDialog
        {
            private const string c_foldersFilter = "Folders|\n";

            private const BindingFlags c_flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            private readonly static Assembly s_windowsFormsAssembly = typeof(FileDialog).Assembly;
            private readonly static Type s_iFileDialogType = s_windowsFormsAssembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog");
            private readonly static MethodInfo s_createVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("CreateVistaDialog", c_flags);
            private readonly static MethodInfo s_onBeforeVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("OnBeforeVistaDialog", c_flags);
            private readonly static MethodInfo s_getOptionsMethodInfo = typeof(FileDialog).GetMethod("GetOptions", c_flags);
            private readonly static MethodInfo s_setOptionsMethodInfo = s_iFileDialogType.GetMethod("SetOptions", c_flags);
            private readonly static uint s_fosPickFoldersBitFlag = (uint)s_windowsFormsAssembly
                .GetType("System.Windows.Forms.FileDialogNative+FOS")
                .GetField("FOS_PICKFOLDERS")
                .GetValue(null);
            private readonly static ConstructorInfo s_vistaDialogEventsConstructorInfo = s_windowsFormsAssembly
                .GetType("System.Windows.Forms.FileDialog+VistaDialogEvents")
                .GetConstructor(c_flags, null, new[] { typeof(FileDialog) }, null);
            private readonly static MethodInfo s_adviseMethodInfo = s_iFileDialogType.GetMethod("Advise");
            private readonly static MethodInfo s_unAdviseMethodInfo = s_iFileDialogType.GetMethod("Unadvise");
            private readonly static MethodInfo s_showMethodInfo = s_iFileDialogType.GetMethod("Show");

            public static ShowDialogResult Show(IntPtr ownerHandle, string initialDirectory, string title)
            {
                var openFileDialog = new OpenFileDialog
                {

                    AddExtension = false,
                    CheckFileExists = false,
                    DereferenceLinks = true,
                    Filter = c_foldersFilter,
                    InitialDirectory = initialDirectory,
                    Multiselect = false,
                    Title = title
                };
                var iFileDialog = s_createVistaDialogMethodInfo.Invoke(openFileDialog, new object[] { });
                s_onBeforeVistaDialogMethodInfo.Invoke(openFileDialog, new[] { iFileDialog });
                s_setOptionsMethodInfo.Invoke(iFileDialog, new object[] { (uint)s_getOptionsMethodInfo.Invoke(openFileDialog, new object[] { }) | s_fosPickFoldersBitFlag });
                var adviseParametersWithOutputConnectionToken = new[] { s_vistaDialogEventsConstructorInfo.Invoke(new object[] { openFileDialog }), 0U };
                s_adviseMethodInfo.Invoke(iFileDialog, adviseParametersWithOutputConnectionToken);

                try
                {
                    int retVal = (int)s_showMethodInfo.Invoke(iFileDialog, new object[] { ownerHandle });
                    return new ShowDialogResult
                    {

                        Result = retVal == 0,
                        FileName = openFileDialog.FileName
                    };
                }
                finally
                {
                    s_unAdviseMethodInfo.Invoke(iFileDialog, new[] { adviseParametersWithOutputConnectionToken[1] });
                }
            }
        }

        // Wrap an IWin32Window around an IntPtr
        private class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _handle;
            public WindowWrapper(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }
    }
    #endregion

    #region CaptureScreenCLS
    /// <summary>
    /// Class used to capture screen and save it to a file
    /// </summary>
    class ScreenCapture
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetDesktopWindow();

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

        public Image CaptureDesktop()
        {
            return CaptureWindow(GetDesktopWindow());
        }

        public Bitmap CaptureActiveWindow()
        {
            return CaptureWindow(GetForegroundWindow());
        }

        public Bitmap CaptureWindow(IntPtr handle)
        {
            var rect = new Rect();
            GetWindowRect(handle, ref rect);
            var bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
            var result = new Bitmap(bounds.Width, bounds.Height);

            using (var graphics = Graphics.FromImage(result))
            {
                graphics.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
            }

            return result;
        }
    }
    #endregion

    #region ThemeCLS
    static class Theming
    {
        /// Handles to Win 32 API
        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string sClassName, string sAppName);
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        /// Windows Constants
        private const uint WM_CLOSE = 0x10;

        private static String StartProcessAndWait(string filename, string arguments, int seconds, ref Boolean bExited)
        {
            String msg = String.Empty;
            Process p = new Process();
            p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
            p.StartInfo.FileName = filename;
            p.StartInfo.Arguments = arguments;
            p.Start();

            bExited = false;
            int counter = 0;
            // give it "seconds" seconds to run
            while (!bExited && counter < seconds)
            {
                bExited = p.HasExited;
                counter++;
                System.Threading.Thread.Sleep(1000);
            }//while
            if (counter == seconds)
            {
                msg = "Program did not close in expected time.";
            }//if

            return msg;
        }

        public static Boolean SwitchTheme(string themePath)
        {
            try
            {
                //String themePath = System.Environment.GetFolderPath(Environment.SpecialFolder.Windows) + @"\Resources\Ease of Access Themes\basic.theme";
                // Set the theme
                Boolean bExited = false;
                // essentially runs the command line:  rundll32.exe %SystemRoot%\system32\shell32.dll,Control_RunDLL %SystemRoot%\system32\desk.cpl desk,@Themes /Action:OpenTheme /file:"%WINDIR%\Resources\Ease of Access Themes\classic.theme"
                String ThemeOutput = StartProcessAndWait("rundll32.exe", System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\shell32.dll,Control_RunDLL " +
                    System.Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\desk.cpl desk,@Themes /Action:OpenTheme /file:\"" + themePath + "\"", 30, ref bExited);

                //Console.WriteLine(ThemeOutput);

                // Wait for the theme to be set
                System.Threading.Thread.Sleep(1000);

                // Close the Theme UI Window
                IntPtr hWndTheming = FindWindow("CabinetWClass", null);
                SendMessage(hWndTheming, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occured while setting the theme: " + ex.Message);

                return false;
            }
            return true;
        }

        public static Boolean SwitchToClassicTheme()
        {
            return SwitchTheme(System.Environment.GetFolderPath(Environment.SpecialFolder.Windows) + @"\Resources\Ease of Access Themes\classic.theme");
        }

        //public Boolean SwitchToAeroTheme()
        //{
        //    return SwitchTheme(System.Environment.GetFolderPath(Environment.SpecialFolder.Windows) + @"\Resources\Themes\aero.theme");
        //}

        //public string GetTheme()
        //{
        //    string RegistryKey = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes";
        //    string theme;
        //    theme = (string)Registry.GetValue(RegistryKey, "CurrentTheme", string.Empty);
        //    theme = theme.Split('\\').Last().Split('.').First().ToString();
        //    return theme;
        //}

        // end of object Theming
    }
    #endregion

}
