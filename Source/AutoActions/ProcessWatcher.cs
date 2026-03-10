using AutoActions.Threading;
using AutoActions.UWP;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace AutoActions
{
    /// <summary>
    /// Optimized ProcessWatcher:
    /// 1. WinEvent hook for foreground focus tracking — event-driven, zero idle CPU cost
    /// 2. Process.GetProcessesByName() for non-UWP apps instead of scanning all processes
    /// 3. Proper Dispose() on all Process objects to prevent handle/GC pressure
    /// 4. Foreground PID cached once per tick, not re-queried per process
    /// </summary>
    public class ProcessWatcher : IManagedThread
    {
        #region P/Invoke — WinEvent hook + message loop

        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const uint WINEVENT_OUTOFCONTEXT   = 0x0000;
        private const uint WM_QUIT                 = 0x0012;

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax,
            IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
            uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll")]
        private static extern bool PostThreadMessage(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint min, uint max);

        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint   message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint   time;
            public int    ptX;
            public int    ptY;
        }

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        #endregion

        // --- State ---
        private volatile int  _foregroundPid        = -1;
        private volatile uint _hookThreadNativeId    = 0;
        private IntPtr        _winEventHookHandle    = IntPtr.Zero;
        private Thread        _hookThread;
        private WinEventDelegate _winEventDelegate;   // must be kept alive (GC)

        private Thread _watchProcessThread;
        private volatile bool _stopRequested = false;
        private bool _isRunning = false;

        private readonly object _applicationsLock = new object();
        private readonly object _accessLock       = new object();

        private Dictionary<ApplicationItem, ApplicationState> _applications
            = new Dictionary<ApplicationItem, ApplicationState>();

        private Dictionary<ApplicationItem, HashSet<int>> _applicationPids
            = new Dictionary<ApplicationItem, HashSet<int>>();

        public bool IsRunning             { get => _isRunning; private set => _isRunning = value; }
        public bool ManagedThreadIsActive => IsRunning;

        public bool OneProcessIsRunning   { get; private set; } = false;
        public bool OneProcessIsFocused   { get; private set; } = false;

        public event EventHandler<string>                      NewLog;
        public event EventHandler<ApplicationChangedEventArgs> ApplicationChanged;

        public ProcessWatcher() { }

        private void CallNewLog(string msg) => NewLog?.Invoke(this, msg);

        // ------------------------------------------------------------------ //
        //  Public API                                                         //
        // ------------------------------------------------------------------ //

        public void AddProcess(ApplicationItem application)
        {
            lock (_applicationsLock)
            {
                if (!_applications.ContainsKey(application))
                {
                    _applications[application]    = ApplicationState.None;
                    _applicationPids[application] = new HashSet<int>();
                    CallNewLog($"Application added to process watcher: {application}");
                }
            }
        }

        public void RemoveProcess(ApplicationItem application)
        {
            lock (_applicationsLock)
            {
                if (_applications.ContainsKey(application))
                {
                    _applications.Remove(application);
                    _applicationPids.Remove(application);
                    CallNewLog($"Application removed from process watcher: {application}");
                }
            }
        }

        public IReadOnlyDictionary<ApplicationItem, ApplicationState> Applications
        {
            get
            {
                lock (_applicationsLock)
                    return new ReadOnlyDictionary<ApplicationItem, ApplicationState>(
                        _applications.ToDictionary(e => e.Key, e => e.Value));
            }
        }

        public void StartManagedThread()
        {
            if (_stopRequested || IsRunning) return;
            lock (_accessLock)
            {
                CallNewLog("Starting process watcher...");
                _stopRequested = false;
                _isRunning     = true;

                try { _foregroundPid = WinAPIFunctions.GetWindowProcessId(WinAPIFunctions.GetforegroundWindow()); }
                catch { }

                StartWinEventHook();

                _watchProcessThread = new Thread(WatchProcessLoop)
                {
                    IsBackground = true,
                    Name         = "AutoActions_ProcessWatcher"
                };
                _watchProcessThread.Start();
                CallNewLog("Process watcher started");
            }
        }

        public void StopManagedThread()
        {
            if (_stopRequested || !IsRunning) return;
            lock (_accessLock)
            {
                CallNewLog("Stopping process watcher...");
                _stopRequested = true;
                StopWinEventHook();
                _watchProcessThread?.Join(5000);
                _stopRequested      = false;
                _isRunning          = false;
                _watchProcessThread = null;
                CallNewLog("Process watcher stopped.");
            }
        }

        // ------------------------------------------------------------------ //
        //  WinEvent Hook — focus detection, completely event-driven           //
        // ------------------------------------------------------------------ //

        private void StartWinEventHook()
        {
            _winEventDelegate = OnWinEventForeground;

            _hookThread = new Thread(() =>
            {
                _hookThreadNativeId = GetCurrentThreadId();

                _winEventHookHandle = SetWinEventHook(
                    EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
                    IntPtr.Zero, _winEventDelegate,
                    0, 0, WINEVENT_OUTOFCONTEXT);

                if (_winEventHookHandle == IntPtr.Zero)
                {
                    CallNewLog("WinEventHook failed — focus changes will be caught by poll loop only.");
                    return;
                }

                MSG msg;
                while (GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
                {
                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }

                UnhookWinEvent(_winEventHookHandle);
                _winEventHookHandle = IntPtr.Zero;
            })
            {
                IsBackground = true,
                Name         = "AutoActions_WinEventHook"
            };
            _hookThread.SetApartmentState(ApartmentState.STA);
            _hookThread.Start();
        }

        private void StopWinEventHook()
        {
            if (_hookThreadNativeId != 0)
            {
                PostThreadMessage(_hookThreadNativeId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                _hookThreadNativeId = 0;
            }
            _hookThread?.Join(2000);
            _hookThread = null;
        }

        private void OnWinEventForeground(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            try
            {
                int pid = WinAPIFunctions.GetWindowProcessId(hwnd);
                _foregroundPid = pid;

                ThreadPool.QueueUserWorkItem(_ =>
                {
                    lock (_applicationsLock)
                        ApplyFocusChange(pid);
                });
            }
            catch { }
        }

        private void ApplyFocusChange(int newForegroundPid)
        {
            foreach (var application in _applications.Keys.ToList())
            {
                ApplicationState oldState = _applications[application];
                if (oldState == ApplicationState.None) continue;

                bool isFocused = _applicationPids.TryGetValue(application, out var pids)
                    && pids.Contains(newForegroundPid);

                if (isFocused && oldState != ApplicationState.Focused)
                {
                    _applications[application] = ApplicationState.Focused;
                    CallApplicationChanged(application, ApplicationChangedType.GotFocus);
                }
                else if (!isFocused && oldState == ApplicationState.Focused)
                {
                    _applications[application] = ApplicationState.Running;
                    CallApplicationChanged(application, ApplicationChangedType.LostFocus);
                }
            }
        }

        // ------------------------------------------------------------------ //
        //  Polling loop — only responsible for start/stop detection           //
        // ------------------------------------------------------------------ //

        private void WatchProcessLoop()
        {
            while (!_stopRequested)
            {
                lock (_applicationsLock)
                    UpdateApplications();
                Thread.Sleep(Globals.GlobalRefreshInterval);
            }
        }

        private void UpdateApplications()
        {
            var applications = _applications.Keys.ToList();
            if (applications.Count == 0) return;

            int currentForegroundPid = _foregroundPid;

            bool hasUwp = applications.Any(a => a.IsUWP);
            Process[] allProcesses = hasUwp ? Process.GetProcesses() : null;

            try
            {
                foreach (ApplicationItem application in applications)
                {
                    ApplicationState oldState = _applications[application];
                    var matchedPids = new HashSet<int>();

                    if (application.IsUWP && allProcesses != null)
                    {
                        foreach (var proc in allProcesses)
                        {
                            try
                            {
                                string procName = proc.ProcessName == "WWAHost"
                                    ? WWAHostHandler.GetProcessName(proc.Id)
                                    : proc.ProcessName;

                                bool matches =
                                    application.ApplicationName.ToUpperInvariant().Equals(procName.ToUpperInvariant())
                                    || (!string.IsNullOrEmpty(application.UWPIdentity)
                                        && procName.Contains(application.UWPIdentity));

                                if (matches) matchedPids.Add(proc.Id);
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // GetProcessesByName: only fetches processes with that name,
                        // much cheaper than enumerating all ~200 system processes
                        Process[] matched = Process.GetProcessesByName(application.ApplicationName);
                        foreach (var p in matched)
                        {
                            matchedPids.Add(p.Id);
                            p.Dispose();
                        }
                    }

                    _applicationPids[application] = matchedPids;

                    ApplicationState newState = ApplicationState.None;
                    bool callNewRunning = false;
                    bool callGotFocus   = false;
                    bool callLostFocus  = false;
                    bool callClosed     = false;

                    if (matchedPids.Count > 0)
                    {
                        bool isFocused = matchedPids.Contains(currentForegroundPid);
                        newState = isFocused ? ApplicationState.Focused : ApplicationState.Running;

                        if (oldState == ApplicationState.None)
                            callNewRunning = true;

                        if (isFocused && oldState != ApplicationState.Focused)
                            callGotFocus = true;
                        else if (!isFocused && oldState == ApplicationState.Focused)
                            callLostFocus = true;
                    }
                    else if (oldState != ApplicationState.None)
                    {
                        callClosed = true;
                    }

                    _applications[application] = newState;

                    if (callNewRunning) CallApplicationChanged(application, ApplicationChangedType.Started);
                    if (callGotFocus)   CallApplicationChanged(application, ApplicationChangedType.GotFocus);
                    if (callLostFocus)  CallApplicationChanged(application, ApplicationChangedType.LostFocus);
                    if (callClosed)     CallApplicationChanged(application, ApplicationChangedType.Closed);
                }
            }
            finally
            {
                if (allProcesses != null)
                    foreach (var p in allProcesses)
                        p.Dispose();
            }
        }

        private void CallApplicationChanged(ApplicationItem application, ApplicationChangedType changedType)
            => ApplicationChanged?.Invoke(this, new ApplicationChangedEventArgs(application, changedType));
    }
}
