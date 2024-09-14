using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;

namespace Keyman
{
    public partial class Form1 : Form
    {
        private static Mutex _mutex;
        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOPMOST = 0x0008;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOMOVE = 0x0002;
        private const char CAPS = '\u0014'; // Capital key code
        private const char ESCAPE = '\u001b'; // Excape key code
        private const char ENTER = '\r'; // Enter key code

        private static LowLevelKeyboardProc _keyboardProc = KeyboardHookCallback;
        private static LowLevelMouseProc _mouseProc = MouseHookCallback;
        private static IntPtr _keyboardHookID = IntPtr.Zero;
        private static IntPtr _mouseHookID = IntPtr.Zero;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private static Random _random = new Random();
        private static bool _isExecuting = false;
        private static Queue<char> _queue = new Queue<char>(10);
        private static List<char> _keys = new List<char>(10);

        private static List<KeyValuePair<List<char>, Action>> _directives = new List<KeyValuePair<List<char>, Action>>();
        private static System.Timers.Timer _timer;

        public Form1()
        {
            EnsureSingletonOtherwiseShutDown();
            InitializeComponent();
            //hide the application
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Hide();
            //register directives
            //Action rapidInput = () => ReadAndTypeTextFromFile("keyman.txt");
            Action topmost = () => ToggleTopMostForCurrentWindow();
            Action exit = () => ExitApplication();
            Action goldbricking = () => KeepActiveWithProtection();
            //_directives.Add(new KeyValuePair<List<char>, Action>(new List<char>() { CAPS, CAPS }, rapidInput));
            _directives.Add(new KeyValuePair<List<char>, Action>(new List<char>() { 'T', 'O', 'P', 'M', 'O', 'S', 'T' }, topmost));
            _directives.Add(new KeyValuePair<List<char>, Action>(new List<char>() { ESCAPE, ESCAPE, ESCAPE, ESCAPE, ESCAPE }, exit));
            _directives.Add(new KeyValuePair<List<char>, Action>(new List<char>() { 'N', 'W', 'A', 'M', 'T', 'F' }, goldbricking));
            //set hook for the application
            _keyboardHookID = SetHookForKeyBoard(_keyboardProc);
            Application.Run(this);
            UnhookWindowsHookEx(_keyboardHookID);
        }

        #region SetHooks
        private static IntPtr SetHookForKeyBoard(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr SetHookForMouse(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    char key = (char)vkCode;
                    //record user's recent keystroke
                    _queue.Enqueue(key);
                    //keep track the last 10 keystrokes
                    if (_queue.Count > 10)
                        _queue.Dequeue();
                    //convert to list, so it can efficiently do comparison
                    _keys = new List<char>(_queue);
                    _keys.Reverse();

                    foreach (var dir in _directives)
                    {
                        //if it matches any directive and executed, then stop looking for other directives
                        if (CheckAndExecute(dir))
                            break;
                    }
                }
                return CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Keyman exited! exception occurred: " + ex.ToString());
                if (_mutex != null)
                    _mutex.ReleaseMutex();
                Application.Exit();
                throw ex;
            }
        }

        private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)// //this condition is only for left mouse click
                {
                    if (_timer != null && _timer.Enabled)
                    {
                        LockWorkStation();
                    }
                }
                return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Keyman exited! exception occurred: " + ex.ToString());
                if (_mutex != null)
                    _mutex.ReleaseMutex();
                Application.Exit();
                throw ex;
            }
        }
        #endregion

        #region Core detect logic & other important functions
        public void EnsureSingletonOtherwiseShutDown()
        {
            bool createNew;
            var mutex = new Mutex(true, "KeymanSingleton", out createNew);
            if (!createNew)
            {
                MessageBox.Show("You already have a keyman who is silently saving your life!");
                throw new Exception();
            }
            _mutex = mutex;
        }

        private static bool CheckAndExecute(KeyValuePair<List<char>, Action> directive)
        {
            if (checkIsInGoldbricking())
                return true;

            bool res = false;
            var i = directive.Key.Count - 1;
            foreach (char c in _keys)
            {
                //if the recent keystrokes don't match trigger, then stop checking
                if (c != directive.Key[i])
                    break;
                //if iterates to 0, that means full match. stop looping
                if (i == 0)
                {
                    res = true;
                    _queue.Clear();
                    break;
                }
                //check all trigger charactors
                if (i > 0)
                    i--;
            }
            if (res)
            {
                //wait for a while in case it interrupts the input processing
                Thread.Sleep(300);
                directive.Value();
            }
            return res;
        }
        #endregion

        #region Goldbricking
        private static bool checkIsInGoldbricking()
        {
            if (_timer != null && _timer.Enabled)
            {
                //prevent the lighthouse function to lock windows
                if (_keys.Count > 0 && _keys[0] == CAPS)
                    return true;

                if (_keys.Count > 0 && _keys[0] == 'Q')
                    ClearTimer();
                else
                    LockWorkStation();

                return true;
            }
            return false;
        }

        private static void KeepActiveWithProtection()
        {
            if (_isExecuting) return;

            try
            {
                _isExecuting = true;
                //pop up a window to get the goldbricking minutes
                int minutes = 0;
                using (Form2 form2 = new Form2())
                {
                    var result = form2.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        string input = form2.InputValue.Trim();
                        input = string.IsNullOrEmpty(input) ? "30" : input; 
                        minutes = int.Parse(input);
                    }
                    else
                        return;
                }
                //keep active
                //minute counter
                int minCounter = 0;
                int triggerSecond = 5;
                //move mouse every 5 seconds
                _timer = new System.Timers.Timer(triggerSecond * 1000);
                _timer.Elapsed += OnTimedEvent;
                _timer.AutoReset = true;
                _timer.Start();
                //set hook for mouse event, in case malicious behave during leave
                _mouseHookID = SetHookForMouse(_mouseProc);

                void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
                {
                    //minus 1, because the first OnTimedEvent will be invoked after 60 seconds
                    if (minCounter < (minutes - 1) * (60 / triggerSecond))
                    {
                        //SetCursorPos(minCounter * 10, minCounter * 10);
                        SendKeys.SendWait("{CAPSLOCK}");
                        minCounter++;
                    }
                    else
                    {
                        _timer.Stop();
                        _timer.Dispose();
                        //after all, lock screen
                        LockWorkStation();
                    }
                }
            }
            finally
            {
                _isExecuting = false;
            }
        }

        private static void ClearTimer()
        {
            if (_isExecuting) return;

            try
            {
                _isExecuting = true;

                if (_timer != null && _timer.Enabled)
                {
                    _timer.Close();
                    _timer.Dispose();
                    MessageBox.Show("The timer has been closed", "Focus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            finally
            {
                _isExecuting = false;
            }
        }
        #endregion

        #region TopMost
        private static void ToggleTopMostForCurrentWindow()
        {
            IntPtr handle = GetForegroundWindow();
            int exStyle = GetWindowLong(handle, GWL_EXSTYLE);
            bool isTopMost = (exStyle & WS_EX_TOPMOST) == WS_EX_TOPMOST;

            if (isTopMost)
                SetWindowPos(handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            else
                SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        }
        #endregion

        #region Exit
        private static void ExitApplication()
        {
            DialogResult result = MessageBox.Show("Do you want to exit keyman?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (_mutex != null)
                    _mutex.ReleaseMutex();
                Application.Exit();
            }
        }
        #endregion

        #region RapidInput
        private static void ReadAndTypeTextFromFile(string filePath)
        {
            if (_isExecuting) return;

            try
            {
                _isExecuting = true;

                using (StreamReader reader = new StreamReader(filePath))
                {
                    string content = reader.ReadToEnd();

                    string[] parts = content.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    string firstPart = parts.Length > 0 ? parts[0] : string.Empty;
                    string secondPart = parts.Length > 1 ? parts[1] : string.Empty;

                    foreach (char c in firstPart)
                    {
                        SendKeyAndSleepAWhilie(c);
                    }

                    if (!string.IsNullOrEmpty(secondPart))
                    {
                        SendKeys.SendWait("{TAB}");
                        string decodedString = DecodeBase64(secondPart);
                        foreach (char dc in decodedString)
                        {
                            SendKeyAndSleepAWhilie(dc);
                        }
                        SendKeys.SendWait("{ENTER}");
                    }
                }
            }
            finally
            {
                _isExecuting = false;
            }
        }
        #endregion

        #region Common utils
        private static void SendKeyAndSleepAWhilie(char c)
        {
            SendKeys.SendWait(c.ToString());
            Thread.Sleep(_random.Next(50, 200));
        }

        private static string DecodeBase64(string base64String)
        {
            byte[] bytes = Convert.FromBase64String(base64String);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }
        #endregion

        #region Win32 API Declarations

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool LockWorkStation();

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        #endregion
    }
}