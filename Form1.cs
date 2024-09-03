using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;

using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using WUApiLib;
using EventLogger.Properties;
using System.Drawing;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Data.Common;

namespace EventLogger
{
    public partial class Form1 : Form
    {

        #region Set Threshold
        //Set Threshold in Seconds
        const int SetTimerTime = 5;
        const int SetReChkEvtTime = 3600;
        const int WifiSec = 300000;

        const int SetCPUTime = 30;
        const int SetMemTime = 10;
        const int SetDiskTime = 36000;
        public string typeChk = "";

        static decimal setCpu = 90;
        static decimal setMem = 90;
        static decimal setDisk = 70;
        static decimal setHDD = 5;
        public  string notify = "";
        public string notifytype = "";

        static decimal WUDy = 900000;
        static decimal WIFISig = 100;

        //Set Thread Threshold in MS
        const int SetResTrackingTime = 5000;
        const int SetEVTTrackingTime = 60000;
        const int SetSNTrackingTime = 360000;
        const int WUTime = 360000;

        #endregion




        DispatcherTimer Timer99 = new DispatcherTimer();
        static PerformanceCounter myCounter = new PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total");
        static PerformanceCounter myCounter1 = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        static List<data> MemData = new List<data>();
        static List<data> CpuData = new List<data>();
        static List<data> DiskData = new List<data>();

        static Dictionary<string, DateTime> AppTrack = new Dictionary<string, DateTime>();
        static Dictionary<string ,DateTime> EvtTrack= new Dictionary<string,DateTime>();
        static Dictionary<string, DateTime> AVTrack = new Dictionary<string, DateTime>();
        static Dictionary<string, DateTime> WUTrack = new Dictionary<string, DateTime>();
        static Dictionary<string, DateTime> MSTTrack = new Dictionary<string, DateTime>();
        static decimal WifiStr = 0;
        static decimal WUDays = 0;

        static string aStr = "";
        static List<Dictionary<int, DateTime>> keyValuePairs = new List<Dictionary<int, DateTime>>();

        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width, Screen.PrimaryScreen.WorkingArea.Height - this.Height);

            Thread ThreadObject1 = new Thread(GetEvt);
            ThreadObject1.Start();
            Thread ThreadObject2 = new Thread(GetRes);
            ThreadObject2.Start();
            Thread ThreadObject3 = new Thread(GetSNT);
            ThreadObject3.Start();
            Thread ThreadObject4 = new Thread(getAntivirusStatus);
            ThreadObject4.Start();
            Thread ThreadObject5 = new Thread(WindowsUpdate);
            ThreadObject5.Start();
        Thread ThreadObject6 = new Thread(WifiSignal);
            ThreadObject6.Start();



            Timer99.Tick += Timer99_Tick;
            Timer99.Interval = new TimeSpan(0, 0, 0, SetTimerTime, 0);
            Timer99.IsEnabled = true;

            
        }

        #region formevents
        /// <summary>
        /// Set event time to restrict the duplicate LLM call
        /// </summary>
        /// <param name="Evt"></param>
        /// <returns></returns>
        public void SetEvtTime(string Evt)
        {
            DateTime EvtTime = DateTime.Now;
            if (EvtTrack.ContainsKey(Evt))
            {
                EvtTime = EvtTrack[Evt];
                TimeSpan span = EvtTime.Subtract(DateTime.Now);
                if (span.TotalSeconds > SetReChkEvtTime)
                {
                    EvtTrack.Remove(Evt);
                    EvtTrack.Add(Evt, DateTime.Now);
                }
            }
            else
            {
                EvtTrack.Add(Evt, DateTime.Now);
            }
        }
        /// <summary>
        /// Get Event Time
        /// </summary>
        /// <param name="Evt"></param>
        /// <returns></returns>
        public bool GetEvtTime(string Evt)
        {
            bool TrigLLM = false;
            DateTime EvtTime = DateTime.Now;
            if (EvtTrack.ContainsKey(Evt))
            {
                EvtTime = EvtTrack[Evt];
                TimeSpan span = EvtTime.Subtract(DateTime.Now);
                if (span.TotalSeconds > SetReChkEvtTime)
                {
                    TrigLLM = true;
                }
                else
                {
                    TrigLLM = false;
                }
            }
            else
            {
                TrigLLM = true;
            }
            return TrigLLM;
        }

        /// <summary>
        /// Timer event to track the system events to trigger LLM call
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Timer99_Tick(object sender, EventArgs e)
        {
            var drive = new DriveInfo("c");

            long freeSpaceInBytes = drive.AvailableFreeSpace;
            long TotalSpaceInBytes = drive.TotalSize;
            long percentage = (freeSpaceInBytes * 100) / TotalSpaceInBytes;
            try
            {
                if (percentage < setHDD && typeChk == "HDD")
                {
                    if (GetEvtTime("HDD"))
                    {
                        textBox1.Text += "Available Disk space in C: " + percentage.ToString() + Environment.NewLine;
                        long harddiskper = 99 - percentage;
                        SetEvtTime("HDD");
                        notify = "Critical hard disk error found, c: running low";
                        notifytype = "warn";
                        await callLLMAsync("Hard disk utilization is " + harddiskper.ToString() + " percentage", true);
                    }
                }
                if (CpuData.Count > 0 && typeChk == "CPU")
                {
                    if (GetEvtTime("CPU"))
                    {
                        textBox1.Text += "CPU Usage : " + CpuData[0].doc.ToShortTimeString() + ":" + CpuData[0].value.ToString() + " # " + CpuData[CpuData.Count - 1].doc.ToShortTimeString() + ":" + CpuData[CpuData.Count - 1].value.ToString() + Environment.NewLine;
                        TimeSpan span = CpuData[CpuData.Count - 1].doc.Subtract(CpuData[0].doc);
                        if (span.TotalSeconds > SetCPUTime)
                        {
                            SetEvtTime("CPU");
                            notify = "High CPU utilization. Scanning the process causing the issue.";
                            notifytype = "err";
                            await callLLMAsync("Processes running with high cpu utilization", true);
                        }
                    }
                }
                if (MemData.Count > 0)
                {
                    if (GetEvtTime("MEM"))
                    {
                        textBox1.Text += "Memory Usage : " + MemData[0].doc.ToShortTimeString() + ":" + MemData[0].value.ToString() + " # " + MemData[MemData.Count - 1].doc.ToShortTimeString() + ":" + MemData[MemData.Count - 1].value.ToString() + Environment.NewLine;
                        TimeSpan span = MemData[MemData.Count - 1].doc.Subtract(MemData[0].doc);
                        if (span.TotalSeconds > SetMemTime)
                        {
                            SetEvtTime("MEM");
                            await callLLMAsync("Memory usage greater than 90 percentage for more than 30 second", true);
                        }
                    }
                }
                if (DiskData.Count > 0)
                {
                    if (GetEvtTime("DISK"))
                    {
                        textBox1.Text += "Disk Usage : " + DiskData[0].doc.ToShortTimeString() + ":" + DiskData[0].value.ToString() + " # " + DiskData[DiskData.Count - 1].doc.ToShortTimeString() + ":" + DiskData[DiskData.Count - 1].value.ToString() + Environment.NewLine;
                        TimeSpan span = DiskData[DiskData.Count - 1].doc.Subtract(DiskData[0].doc);
                        if (span.TotalSeconds > SetDiskTime)
                        {
                            SetEvtTime("DISK");
                            await callLLMAsync("Disk usage greater than 90 percentage for more than 30 second", true);
                        }
                    }
                }
                if (AppTrack.Count > 0)
                {
                    //Change code to track hourly basis
                    TimeSpan span = DateTime.Now.Subtract(AppTrack.Values.FirstOrDefault());
                    if (span.TotalSeconds > SetDiskTime)
                    {
                        await callLLMAsync(AppTrack.Keys.FirstOrDefault(), true);
                    }
                    AppTrack.Clear();
                }
                if (AVTrack.Count > 0)
                {
                    foreach (var a in AVTrack)
                    {
                        TimeSpan span = DateTime.Now.Subtract(a.Value);
                        if (span.TotalSeconds > SetDiskTime)
                        {
                            await callLLMAsync(a.Key, true);
                        }
                    }
                    AVTrack.Clear();
                }

                if (WUDays > WUDy && typeChk == "WU")
                {

                    if (WUTrack.ContainsKey("WUServiceTime"))
                    {
                        TimeSpan span = DateTime.Now.Subtract(WUTrack["WUServiceTime"]);
                        if (span.TotalSeconds < SetDiskTime)
                        {
                            return;
                        }
                        else
                        {
                            WUTrack.Clear();
                        }
                    }
                    WUTrack.Add("WUServiceTime", DateTime.Now);
                    notify = "Compliance violation, pending windows update.";
                    notifytype = "warn";
                    await callLLMAsync("windows update pending since last 10 days", true);

                }


                if (WifiStr < 50 && WifiStr > 0 && typeChk == "WIFI")
                {
                    notify = "WIFI signal is low, may affect performance.";
                    notifytype = "warn";
                    await callLLMAsync("WIFI strenght is Weak " + WifiStr.ToString() + "%", false);
                }
                if (IsAdobeCrashed() && typeChk == "APP CRASH")
                {
                    if (MSTTrack.ContainsKey("MSTServiceTime"))
                    {
                        TimeSpan span = DateTime.Now.Subtract(MSTTrack["MSTServiceTime"]);
                        if (span.TotalSeconds < SetDiskTime)
                        {
                            return;
                        }
                        else
                        {
                            MSTTrack.Clear();
                        }
                    }
                    MSTTrack.Add("MSTServiceTime", DateTime.Now);
                    notify = "Application crash tracked- Adobe reader";
                    notifytype = "err";
                    await callLLMAsync("Application found in crash state - Adobe reader", false);
                }
            }
            catch (Exception ex) { }
            flowLayoutPanel1.ScrollControlIntoView(label2);
        }

        public bool IsAdobeCrashed()
        {

            if (File.Exists("C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.dll"))
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        public bool IsTeamsAvailable()
        {
            bool isTeam = true;
            bool isOpen = IsProcessOpen("teams-1");
            bool isOpen1 = IsProcessOpen("msteams-1");
            bool isOpen2 = IsProcessOpen("ms-teams-1");
           // bool aIsRunning = ProgramIsRunning(@"C:\\Users\\RA20441721\\AppData\\Local\\Microsoft\\Teams\\Update.exe");
            if (!isOpen1 || !isOpen || !isOpen2)
            {
                ExecuteCommand("Start-Process -File $env:LOCALAPPDATA\\Microsoft\\Teams\\Update.exe -ArgumentList '--processStart \"Teams.exe\"'");
                Thread.Sleep(10000);
                isOpen = IsProcessOpen("teams");
                isOpen1 = IsProcessOpen("msteams");
                isOpen2 = IsProcessOpen("ms-teams");
                if (!isOpen1 || !isOpen || !isOpen2)
                {
                    //MessageBox.Show("teams is not running");
                    isTeam = false;
                }
            }
            return isTeam;
        }

        public bool IsProcessOpen(string name)
        {
            //string aProc = "";
            foreach (Process clsProcess in Process.GetProcesses())
            {
                //  aProc += " , " + clsProcess.ProcessName;
                if (clsProcess.ProcessName.Contains(name))
                {
                    return true;
                }
            }
            //aProc = aProc + "--";
            return false;
        }

        /// <summary>
        /// Get WIFI Strength
        /// </summary>
        /// <returns></returns>
        static void WifiSignal()
        {
            Thread.Sleep(WifiSec);
            decimal SignalStr = 100;
            try
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "netsh.exe";
                p.StartInfo.Arguments = "wlan show interfaces";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.Start();
                
                string s = p.StandardOutput.ReadToEnd();
                string s1 = s.Substring(s.IndexOf("SSID"));
                s1 = s1.Substring(s1.IndexOf(":"));
                s1 = s1.Substring(2, s1.IndexOf("\n")).Trim();

                string s2 = s.Substring(s.IndexOf("Signal"));
                s2 = s2.Substring(s2.IndexOf(":"));
                s2 = s2.Substring(2, s2.IndexOf("\n")).Trim();

                s2 = s2.Replace("%", "");
                SignalStr = Convert.ToDecimal(s2);

                p.WaitForExit();
                p.Close();
            }
            catch
            {
                SignalStr= 0;
            }
            WifiStr = SignalStr;         
           
        }

        /// <summary>
        /// Form Load 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button3.Enabled = false;
            button2.Enabled = false;
            this.Hide();
            textBox1.Text = "";
        }
        /// <summary>
        /// Code to restrict closing the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }
        #endregion

        
        #region action events
        /// <summary>
        /// restart system
        /// </summary>
        public void RestartPC() {
            ExecuteCommand("Restart-Computer");
        }
        /// <summary>
        /// disk cleanup
        /// </summary>
        public void diskCleanup()
        {

            string aCmd = @"write-Host ""Disk Cleanup Started"" -ForegroundColor Red  
$Path = 'C' + ':\$Recycle.Bin'
Clear-RecycleBin -Force
Get - ChildItem $Path - Force - Recurse - ErrorAction SilentlyContinue |
Remove - Item - Recurse - Exclude *.ini - ErrorAction SilentlyContinue


$Path1 = 'C' + ':\Windows\Temp'
Get - ChildItem $Path1 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path2 = 'C' + ':\Windows\Prefetch'
Get - ChildItem $Path2 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path3 = 'C' + ':\Users\*\AppData\Local\Temp'
Get - ChildItem $Path4 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue

cleanmgr / sagerun:1 | out-Null
Sleep 5
write-Host ""Disk Cleanup Successfully done"" -ForegroundColor Green  
Sleep 10";
            ExecuteCommand(aCmd);
        }
        /// <summary>
        /// temp file clean
        /// </summary>
        public void tempClean()
        {

            string aCmd = @"
$Path1 = 'C' + ':\Windows\Temp'
Get - ChildItem $Path1 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path2 = 'C' + ':\Windows\Prefetch'
Get - ChildItem $Path2 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path3 = 'C' + ':\Users\*\AppData\Local\Temp'
Get - ChildItem $Path4 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue";
            ExecuteCommand(aCmd);
        }
        /// <summary>
        /// updateWindows
        /// </summary>
        public void updateWindows()
        {
            string aCmd = @"Get-WindowsUpdate -AcceptAll -Install -AutoReboot";
            ExecuteCommand(aCmd);
        }

        public void diskCheck()
        {
            string aCmd = @"Repair-Volume -DriveLetter D -Scan -Verbose
            Read-Host ""Issue Resolved"" ";
            ExecuteCommand(aCmd);
        }

        /// <summary>
        /// cache clean
        /// </summary>
        public void cacheClean()
        {

            string aCmd = @"#------------------------------------------------------------------#
            #- Clear-GlobalWindowsCache                                        #
            #------------------------------------------------------------------#
Function Clear-GlobalWindowsCache {
    Remove-CacheFiles 'C:\Windows\Temp' 
    Remove-CacheFiles ""C:\`$Recycle.Bin""
    Clear-RecycleBin -Force
    Remove-CacheFiles ""C:\Windows\Prefetch""
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 255
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351
}

#------------------------------------------------------------------#
#- Clear-UserCacheFiles                                            #
#------------------------------------------------------------------#
Function Clear-UserCacheFiles {
    # Stop-BrowserSessions
    ForEach($localUser in (Get-ChildItem 'C:\users').Name)
    {
        Clear-ChromeCache $localUser
        Clear-EdgeCacheFiles $localUser
        Clear-FirefoxCacheFiles $localUser
        Clear-WindowsUserCacheFiles $localUser
        Clear-TeamsCacheFiles $localUser
    }
}

#------------------------------------------------------------------#
#- Clear-WindowsUserCacheFiles                                     #
#------------------------------------------------------------------#
Function Clear-WindowsUserCacheFiles {
    param([string]$user=$env:USERNAME)
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Temp""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\WER""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCookies""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatUaCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IEDownloadHistory""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\Temporary Internet Files""    
}

#Region HelperFunctions

#------------------------------------------------------------------#
#- Stop-BrowserSessions                                            #
#------------------------------------------------------------------#
Function Stop-BrowserSessions {
   $activeBrowsers = Get-Process Firefox*,Chrome*,Waterfox*,Edge*
   ForEach($browserProcess in $activeBrowsers)
   {
       try 
       {
           $browserProcess.CloseMainWindow() | Out-Null 
       } catch { }
   }
}

#------------------------------------------------------------------#
#- Get-StorageSize                                                 #
#------------------------------------------------------------------#
Function Get-StorageSize {
    Get-WmiObject Win32_LogicalDisk | 
    Where-Object { $_.DriveType -eq ""3"" } | 
    Select-Object SystemName, 
        @{ Name = ""Drive"" ; Expression = { ( $_.DeviceID ) } },
        @{ Name = ""Size (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Size / 1gb)}},
        @{ Name = ""FreeSpace (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Freespace / 1gb ) } },
        @{ Name = ""PercentFree"" ; Expression = {""{0:P1}"" -f ( $_.FreeSpace / $_.Size ) } } |
    Format-Table -AutoSize | Out-String
}

#------------------------------------------------------------------#
#- Remove-CacheFiles                                               #
#------------------------------------------------------------------#
Function Remove-CacheFiles {
    param([Parameter(Mandatory=$true)][string]$path)    
    BEGIN 
    {
        $originalVerbosePreference = $VerbosePreference
        $VerbosePreference = 'Continue'  
    }
    PROCESS 
    {
        if((Test-Path $path))
        {
            if([System.IO.Directory]::Exists($path))
            {
                try 
                {
                    if($path[-1] -eq '\')
                    {
                        [int]$pathSubString = $path.ToCharArray().Count - 1
                        $sanitizedPath = $path.SubString(0, $pathSubString)
                        Remove-Item -Path ""$sanitizedPath\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
                    }
                    else 
                    {
                        Remove-Item -Path ""$path\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose              
                    } 
                } catch { }
            }
            else 
            {
                try 
                {
                    Remove-Item -Path $path -Force -ErrorAction SilentlyContinue -Verbose
                } catch { }
            }
        }    
    }
    END 
    {
        $VerbosePreference = $originalVerbosePreference
    }
}

#Endregion HelperFunctions

#Region Browsers

#Region ChromiumBrowsers

#------------------------------------------------------------------#
#- Clear-ChromeCache                                               #
#------------------------------------------------------------------#
Function Clear-ChromeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Google\Chrome\User Data\Default""))
    {
        $chromeAppData = ""C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default"" 
        $possibleCachePaths = @('Cache','Cache2\entries\','Cookies','History','Top Sites','VisitedLinks','Web Data','Media Cache','Cookies-Journal','ChromeDWriteFontCache')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$chromeAppData\$cachePath""
        }      
    } 
}

#------------------------------------------------------------------#
#- Clear-EdgeCache                                                 #
#------------------------------------------------------------------#
Function Clear-EdgeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""))
    {
        $EdgeAppData = ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""
        $possibleCachePaths = @('Cache','Cache2\entries','Cookies','History','Top Sites','Visited Links','Web Data','Media History','Cookies-Journal')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$EdgeAppData$cachePath""
        }
        }
}

#Endregion ChromiumBrowsers

#Region FirefoxBrowsers

#------------------------------------------------------------------#
#- Clear-FirefoxCacheFiles                                         #
#------------------------------------------------------------------#
Function Clear-FirefoxCacheFiles {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $firefoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName 
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$firefoxAppDataPath\$cachePath""
        }
    } 
}

#------------------------------------------------------------------#
#- Clear-WaterfoxCacheFiles                                        #
#------------------------------------------------------------------#
Function Clear-WaterfoxCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Waterfox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $waterfoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Waterfox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$waterfoxAppDataPath\$cachePath""
        }
    }   
}

#Endregion FirefoxBrowsers

#Endregion Browsers

#Region CommunicationPlatforms

#------------------------------------------------------------------#
#- Clear-TeamsCacheFiles                                           #
#------------------------------------------------------------------#
Function Clear-TeamsCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Roaming\Microsoft\Teams""))
    {
        $possibleCachePaths = @('cache','blob_storage','databases','gpucache','Indexeddb','Local Storage','application cache\cache')
        $teamsAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Roaming\Microsoft\Teams"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$teamsAppDataPath\$cachePath""
        }
    }   
}

#Endregion CommunicationPlatforms

#------------------------------------------------------------------#
#- MAIN                                                            #
#------------------------------------------------------------------#

$StartTime = (Get-Date)

Get-StorageSize

Clear-UserCacheFiles
Clear-GlobalWindowsCache

Get-StorageSize

$EndTime = (Get-Date)
Write-Verbose ""Elapsed Time: $(($StartTime - $EndTime).totalseconds) seconds""";
            ExecuteCommand(aCmd);
        }

        /// <summary>
        /// browser data clean
        /// </summary>
    public void browserClean()
        {

            string aCmd = @"#------------------------------------------------------------------#
#- Clear-GlobalWindowsCache                                        #
#------------------------------------------------------------------#
Function Clear-GlobalWindowsCache {
    Remove-CacheFiles 'C:\Windows\Temp' 
    Remove-CacheFiles ""C:\`$Recycle.Bin""
    Remove-CacheFiles ""C:\Windows\Prefetch""
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 255
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351
}

#------------------------------------------------------------------#
#- Clear-UserCacheFiles                                            #
#------------------------------------------------------------------#
Function Clear-UserCacheFiles {
    # Stop-BrowserSessions
    ForEach($localUser in (Get-ChildItem 'C:\users').Name)
    {
        Clear-ChromeCache $localUser
        Clear-EdgeCacheFiles $localUser
        Clear-FirefoxCacheFiles $localUser
        #Clear-WindowsUserCacheFiles $localUser
        #Clear-TeamsCacheFiles $localUser
    }
}

#------------------------------------------------------------------#
#- Clear-WindowsUserCacheFiles                                     #
#------------------------------------------------------------------#
Function Clear-WindowsUserCacheFiles {
    param([string]$user=$env:USERNAME)
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Temp""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\WER""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCookies""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatUaCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IEDownloadHistory""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\Temporary Internet Files""    
}

#Region HelperFunctions

#------------------------------------------------------------------#
#- Stop-BrowserSessions                                            #
#------------------------------------------------------------------#
Function Stop-BrowserSessions {
   $activeBrowsers = Get-Process Firefox*,Chrome*,Waterfox*,Edge*
   ForEach($browserProcess in $activeBrowsers)
   {
       try 
       {
           $browserProcess.CloseMainWindow() | Out-Null 
       } catch { }
   }
}

#------------------------------------------------------------------#
#- Get-StorageSize                                                 #
#------------------------------------------------------------------#
Function Get-StorageSize {
    Get-WmiObject Win32_LogicalDisk | 
    Where-Object { $_.DriveType -eq ""3"" } | 
    Select-Object SystemName, 
        @{ Name = ""Drive"" ; Expression = { ( $_.DeviceID ) } },
        @{ Name = ""Size (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Size / 1gb)}},
        @{ Name = ""FreeSpace (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Freespace / 1gb ) } },
        @{ Name = ""PercentFree"" ; Expression = {""{0:P1}"" -f ( $_.FreeSpace / $_.Size ) } } |
    Format-Table -AutoSize | Out-String
}

#------------------------------------------------------------------#
#- Remove-CacheFiles                                               #
#------------------------------------------------------------------#
Function Remove-CacheFiles {
    param([Parameter(Mandatory=$true)][string]$path)    
    BEGIN 
    {
        $originalVerbosePreference = $VerbosePreference
        $VerbosePreference = 'Continue'  
    }
    PROCESS 
    {
        if((Test-Path $path))
        {
            if([System.IO.Directory]::Exists($path))
            {
                try 
                {
                    if($path[-1] -eq '\')
                    {
                        [int]$pathSubString = $path.ToCharArray().Count - 1
                        $sanitizedPath = $path.SubString(0, $pathSubString)
                        Remove-Item -Path ""$sanitizedPath\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
                    }
                    else 
                    {
                        Remove-Item -Path ""$path\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose              
                    } 
                } catch { }
            }
            else 
            {
                try 
                {
                    Remove-Item -Path $path -Force -ErrorAction SilentlyContinue -Verbose
                } catch { }
            }
        }    
    }
    END 
    {
        $VerbosePreference = $originalVerbosePreference
    }
}

#Endregion HelperFunctions

#Region Browsers

#Region ChromiumBrowsers

#------------------------------------------------------------------#
#- Clear-ChromeCache                                               #
#------------------------------------------------------------------#
Function Clear-ChromeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Google\Chrome\User Data\Default""))
    {
        $chromeAppData = ""C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default"" 
        $possibleCachePaths = @('Cache','Cache2\entries\','Cookies','History','Top Sites','VisitedLinks','Web Data','Media Cache','Cookies-Journal','ChromeDWriteFontCache')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$chromeAppData\$cachePath""
        }      
    } 
}

#------------------------------------------------------------------#
#- Clear-EdgeCache                                                 #
#------------------------------------------------------------------#
Function Clear-EdgeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""))
    {
        $EdgeAppData = ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""
        $possibleCachePaths = @('Cache','Cache2\entries','Cookies','History','Top Sites','Visited Links','Web Data','Media History','Cookies-Journal')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$EdgeAppData$cachePath""
        }
        }
}

#Endregion ChromiumBrowsers

#Region FirefoxBrowsers

#------------------------------------------------------------------#
#- Clear-FirefoxCacheFiles                                         #
#------------------------------------------------------------------#
Function Clear-FirefoxCacheFiles {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $firefoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName 
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$firefoxAppDataPath\$cachePath""
        }
    } 
}

#------------------------------------------------------------------#
#- Clear-WaterfoxCacheFiles                                        #
#------------------------------------------------------------------#
Function Clear-WaterfoxCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Waterfox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $waterfoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Waterfox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$waterfoxAppDataPath\$cachePath""
        }
    }   
}

#Endregion FirefoxBrowsers

#Endregion Browsers

#Region CommunicationPlatforms

#------------------------------------------------------------------#
#- Clear-TeamsCacheFiles                                           #
#------------------------------------------------------------------#
Function Clear-TeamsCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Roaming\Microsoft\Teams""))
    {
        $possibleCachePaths = @('cache','blob_storage','databases','gpucache','Indexeddb','Local Storage','application cache\cache')
        $teamsAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Roaming\Microsoft\Teams"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$teamsAppDataPath\$cachePath""
        }
    }   
}

#Endregion CommunicationPlatforms

#- MAIN                                                            #


Clear-UserCacheFiles
";
            ExecuteCommand(aCmd);
        }

        #endregion


        #region functions

        /// <summary>
        /// Get System Resource
        /// </summary>
        static void GetRes()
        {
            Thread.Sleep(SetResTrackingTime);
            decimal diskUsage = Convert.ToDecimal(myCounter.NextValue());
            decimal a1 = Convert.ToDecimal(myCounter1.NextValue());
            if (a1 > setCpu)
            {
                CpuData.Add(new data { value = a1, doc = DateTime.Now });
            }
            else
            {
                CpuData.Clear();
            }
            decimal aMem = GetMemPer();
            if (aMem > setMem)
            {
                MemData.Add(new data { value = aMem, doc = DateTime.Now });
            }
            else
            {
                MemData.Clear();
            }
            if (diskUsage > setDisk)
            {
                DiskData.Add(new data { value = diskUsage, doc = DateTime.Now });
            }
            else
            {
                DiskData.Clear();
            }

            GetRes();
        }
        /// <summary>
        /// Get Memory utilization percentage
        /// </summary>
        /// <returns></returns>
        static decimal GetMemPer()
        {
            var wmiObject = new ManagementObjectSearcher("select * from Win32_OperatingSystem");

            var memoryValues = wmiObject.Get().Cast<ManagementObject>().Select(mo => new
            {
                FreePhysicalMemory = Double.Parse(mo["FreePhysicalMemory"].ToString()),
                TotalVisibleMemorySize = Double.Parse(mo["TotalVisibleMemorySize"].ToString())
            }).FirstOrDefault();

            if (memoryValues != null)
            {
                var percent = ((memoryValues.TotalVisibleMemorySize - memoryValues.FreePhysicalMemory) / memoryValues.TotalVisibleMemorySize) * 100;
                return Convert.ToDecimal(percent);
            }
            else { return 0; }
        }

        /// <summary>
        /// Get System Events
        /// </summary>
        static void GetEvt()
        {
            Thread.Sleep(SetEVTTrackingTime);
            EventLog myLog = new EventLog();
            myLog.Log = "Application";

            aStr = myLog.Entries.Count.ToString();
            int a = 0;

            var entries = myLog.Entries.Cast<EventLogEntry>()
                         .Where(x => x.TimeGenerated <= DateTime.Now.AddMinutes(-1) && x.Category== "Application Crashing Events")
                         .OrderBy(x => x.TimeGenerated)
                         .Select(x => new
                         {
                             x.EventID,
                             x.InstanceId,
                             x.TimeGenerated,
                             x.MachineName,
                             x.Site,
                             x.Category,
                             x.Source,
                             x.Message
                         }).ToList();

            foreach (var entry in entries)
            {
                Dictionary<int, DateTime> dic = new Dictionary<int, DateTime>();
                dic.Add(entry.EventID, entry.TimeGenerated);
                bool isExist = keyValuePairs.Contains(dic);

                if (!isExist)
                {
                    keyValuePairs.Add(dic);
                    AppTrack.Add(entry.Message , entry.TimeGenerated);
                    a++;
                    aStr += "********************************************************************" + Environment.NewLine;
                    aStr += "Event-Instance:" + entry.EventID.ToString() + "-" + entry.InstanceId.ToString() + Environment.NewLine;
                    aStr += "Datetime:" + entry.TimeGenerated.ToShortDateString() + "-" + entry.TimeGenerated.ToShortTimeString() + Environment.NewLine;
                    aStr += "Category:" + entry.Category + Environment.NewLine;
                    aStr += "Machine:" + entry.MachineName + Environment.NewLine;
                    aStr += "Message:" + entry.Message + Environment.NewLine;
                    aStr += "Source:" + entry.Source + Environment.NewLine;
                }
            }

        }

        /// <summary>
        /// Get Service Now Ticket raised by the local machine
        /// </summary>
        static async void GetSNT()
        {
            
            try
            {
                string aMachine = Environment.MachineName;
                //aMachine = "s-intelema-001";
                //aMachine = "IRULMATHI";

                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Get, "https://wiprodemo4.service-now.com/api/now/table/incident?sysparm_query=cmdb_ci.name%3D" + aMachine + "&sysparm_limit=1");
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=DB2F201541916D323DBAE065786E6AFB; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=4F84FBA81BFD8E106777631ABC4BCB97; glide_user=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9; glide_user_route=glide.acd895053673453505bf5bcc2c835d68; glide_user_session=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9");
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

                var data = (JObject)JsonConvert.DeserializeObject(json);
                string data1 = data["result"][0]["short_description"].Value<string>();
                string sysid = data["result"][0]["sys_id"].Value<string>();
                //MessageBox.Show(data1);
                await new Form1().callLLMAsync(data1,false, sysid);
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);
            }
            Thread.Sleep(SetSNTrackingTime);
            GetSNT();
        }
        /// <summary>
        /// Create New Ticket if no resolution available in smartops for a non incident.
        /// </summary>
        /// <param name="shortDesc"></param>
        /// <param name="comment"></param>
        /// <returns></returns>
        static async Task<string> CreateTicket(string shortDesc, string comment)
        {
            string aTktNo = "";
            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Post, "https://wiprodemo4.service-now.com/api/now/v1/table/incident");
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=23A46115D0D9668BA264BC2FAAD48FB6; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=BA110E111BB582506777631ABC4BCB63; glide_user_route=glide.acd895053673453505bf5bcc2c835d68");
                var content = new StringContent("{\r\n\"short_description\":\"" + shortDesc + "\",\r\n\"comments\":\"" + comment + "\"\r\n}", null, "application/json");
                request.Content = content;
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

                var data = (JObject)JsonConvert.DeserializeObject(json);
                aTktNo = data["result"]["number"].Value<string>();
                string aSysId=data["result"]["sys_id"].Value<string>();
                aTktNo = aSysId;
                return aTktNo;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return "";
            }
        }
        /// <summary>
        /// Update ticket if there is no resolution to an existing ticket
        /// </summary>
        /// <param name="sysid"></param>
        /// <param name="note"></param>
        static async void UpdateTicket(String sysid, string note)
        {

            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Put, "https://wiprodemo4.service-now.com/api/now/table/incident/" + sysid);
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=DB2F201541916D323DBAE065786E6AFB; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=4F84FBA81BFD8E106777631ABC4BCB97; glide_user=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9; glide_user_route=glide.acd895053673453505bf5bcc2c835d68; glide_user_session=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9");
                var content = new StringContent("{\r\n\"work_notes\":\"" + note.Replace("\r\n", "") + "\"\r\n}", null, "application/json");
                request.Content = content;
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }



        /// <summary>
        /// execute script with file path
        /// </summary>
        /// <param name="pathToScript"></param>
        /// <returns></returns>
        public string ExecuteScript(string pathToScript)
        {
            var scriptArguments = "-ExecutionPolicy Bypass -File \"" + pathToScript + "\"";
            var processStartInfo = new ProcessStartInfo("powershell.exe", scriptArguments);
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.RedirectStandardError = true;

            var process = new Process();
            process.StartInfo = processStartInfo;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            return output; // I am invoked using ProcessStartInfoClass!
        }
        /// <summary>
        /// execute powershell command
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public string ExecuteCommand(string command)
        {
            var processStartInfo = new ProcessStartInfo();
            processStartInfo.FileName = "powershell.exe";
            processStartInfo.Arguments = $"-Command \"{command}\"";
            processStartInfo.UseShellExecute = false;
            processStartInfo.RedirectStandardOutput = true;

            var process = new Process();
            process.StartInfo = processStartInfo;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            return output;
        }


        /// <summary>
        /// Nexthink Remote Action
        /// </summary>
        /// <param name="remoteactionId"></param>
        public async void CallNQL(string remoteactionId) {
            try {
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Post, "https://uxmlive.api.meta.nexthink.cloud/api/v1/token");
                request.Headers.Add("Authorization", "Basic N2EzZDEzNGUyMDEwNDY3OThkYTQzMjEwZmE4YmIwM2U6TlZINjJZYnNEUW5RZWUyd2daU2lzOElCYUszNmV5NkNlZzNaMU9vaEhaQXFyaHFsMmhEMmZDamFSRDM1cl90Rw==");
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string strAuth = await response.Content.ReadAsStringAsync();
                var AuthTok = (JObject)JsonConvert.DeserializeObject(strAuth);
                string token = AuthTok["access_token"].Value<string>();
                //getting token
                client = new HttpClient();
                request = new HttpRequestMessage(HttpMethod.Post, "https://uxmlive.api.meta.nexthink.cloud/api/v1/act/execute");
                request.Headers.Add("Authorization", "Bearer " + token);
                var content = new StringContent("{\r\n  \"remoteActionId\": \"" + remoteactionId + "\",\r\n  \"devices\": [\r\n    \"30adce7c-657f-48d3-83e4-301e429475ce\"\r\n  ],\r\n  \"expiresInMinutes\": 60\r\n}", null, "application/json");
                request.Content = content;
                response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string resJson = await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex) { }
            }





        /// <summary>
        /// Call LLM API
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public async Task callLLMAsync(string query, bool RaiseTkt,string sysId="") {

            textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[LLM Query]" + query + Environment.NewLine;

            string aTktNo = sysId;//SysId
            if (RaiseTkt)
            {
                aTktNo = await CreateTicket(query, "created by agent");
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[SN Ticket]<"+ aTktNo +">" + query + Environment.NewLine;
            }
            var options = new RestClientOptions("http://127.0.0.1:8080")
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest("/apiAgent?query=" + query, Method.Post);
            request.AddHeader("Accept", "application/json");
            RestResponse response = await client.ExecuteAsync(request);

            var data= (JObject)JsonConvert.DeserializeObject(response.Content);


            string action = "";
            //Abhijeet Adding second label
            string action2 = "";
            string aNote = "";
            string Knotes = "";
            try
            {
                action = data["response"]["labels"][0].ToString();
            }
            catch (Exception ex) { }
            //Abhijeet Adding second label
            try
            {
                action2 = data["response"]["labels"][1].ToString();
            }
            catch (Exception ex) { }
            try
            {
                aNote = data["response"]["sequence"].ToString();
            }
            catch (Exception ex) { }
            try
            {
                Knotes = data["response"].ToString();
            }
            catch (Exception ex) { }


            if (aTktNo!="")
            {
                //update ticket
                UpdateTicket(aTktNo, aNote);
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[SN Ticket Updated]<" + aTktNo + ">"  + Environment.NewLine;
            }
            if (Knotes == "Level2")
            {
                //update ticket
                aNote = "We have analyzed this issue and identified this needs to be reviewed by Level 2 as I am not authorized to take action outside knowledge policy";
                UpdateTicket(aTktNo, aNote);
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[SN Ticket Updated]<" + aTktNo + ">" + aNote + Environment.NewLine;
            }
            if (action == "Clear temp files") {
                tempClean();
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "clear hard disk space for high disk utilization")
            {
                notify = "Disk cleanup initiated.";
                notifytype = "info";
                diskCleanup();
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "Clear Browser settings")
            {
                browserClean();
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "Restart")
            {
                DialogResult res = MessageBox.Show("System Needs to restart to run smoothly", "Restart PC", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.Yes)
                {
                    RestartPC();
                }                
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "delete Cache Cookies")
            {
                cacheClean();
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "Move closer to the router for better Wifi signal strength")
            {
                notify = "Wifi Signal is Weak, try getting closer to the router";
                notifytype = "info";
                MessageBox.Show("Low WIFI signal strength, try getting closer to the router.");
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "No Internet")
            {
                MessageBox.Show("Check your available internet options");
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action]" + action + Environment.NewLine;
            }
            if (action == "Run windows update")
            {
                notify = "Make sure to update windows for optimal computing. ";
                notifytype = "info";
                DialogResult res= MessageBox.Show("Allow to Start Windows Update (Note: PC May Restart)", "Initiate Update!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.Yes)
                {
                    notify = "Windows update will run soon, make sure to save your work as PC may restart. ";
                    notifytype = "info";
                    updateWindows();
                    //api to call update on NQL pending
                }
                else {
                    notify = "Windows update cancelled.";
                    notifytype = "warn";
                }
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Action] Initiate Windows Update-" + action + Environment.NewLine;
            }
       
            if (action == "nexthink")
            {
                string remoteactionId =  data["response"]["remoteactionId"].Value<string>();
                CallNQL(remoteactionId);
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Call NQLAPI]" + remoteactionId + Environment.NewLine;
            }
            if (action == "Reinstall MS Teams")
            {
                string remoteactionId = "reinstall_microsoft_teams_windows";
                CallNQL(remoteactionId);
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Call NQLAPI]" + remoteactionId + Environment.NewLine;
            }
            //Kill Process
            if (action == "Kill processes with high CPU utilization")
            {
                notify = "Select the right process before hit the kill button.";
                notifytype = "warn";
                //Kill process prompt;  
                ProcessCPU cpu = new ProcessCPU();
                cpu.ShowDialog();
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Kill Process] Prompt to kill proecss." + Environment.NewLine;
            }
            if (action == "recover the application from crash state")
            {
                CallNQL("reinstall_adobe_acrobat_reader");
                notify = "Nexthink will repair the application shortly.";
                notifytype = "info";
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[Call NQLAPI]" + "reinstall_adobe_acrobat_reader" + Environment.NewLine;
            }
            //update ticket with knowledge search
            if (action == "")
            {
                //update ticket
                aNote = "I have reviewed this issue and found this information which may assist you to fix this issue, " + Knotes;
                UpdateTicket(aTktNo, aNote);
                textBox1.Text += DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":[SN Ticket Updated]<" + aNote + ">" + aNote + Environment.NewLine;
            }

        }


        

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            if (setting.ShowDialog() == DialogResult.OK)
            {
                string type = setting.Type;
                typeChk = type;
                if (type == "CPU")
                {
                    setCpu = setting.val;
                }
                if (type == "HDD")
                {
                    setHDD = setting.val;
                }
                if (type == "WU")
                {
                    WUDy = setting.val;
                }
                if (type == "WIFI")
                {
                    WIFISig = setting.val;
                }
            }
        }
        /// <summary>
        /// Get Antivirus Issues
        /// </summary>
        public void getAntivirusStatus()
        {
            ManagementObjectSearcher wmiData = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM AntiVirusProduct");
            ManagementObjectCollection data = wmiData.Get();
            AVTrack.Clear();

            foreach (ManagementObject virusChecker in data)
            {
                if (virusChecker["productState"].ToString() == "262144") {
                    if (!AVTrack.ContainsKey("AVG Internet Security disabled and up to date"))
                    {
                        AVTrack.Add("AVG Internet Security disabled and up to date", DateTime.Now);
                    }                   
                }
                if (virusChecker["productState"].ToString() == "262160")
                {
                    if (!AVTrack.ContainsKey("AVG Internet Security firewall disabled")){
                        AVTrack.Add("AVG Internet Security firewall disabled", DateTime.Now);
                    }
                }
                if (virusChecker["productState"].ToString() == "397584")
                {
                    if (!AVTrack.ContainsKey("Windows Defender enabled and out of date")){
                        AVTrack.Add("Windows Defender enabled and out of date", DateTime.Now);
                    }
                }
                if (virusChecker["productState"].ToString() == "393472")
                {
                    if (!AVTrack.ContainsKey("Windows Defender disabled and up to date")) {
                        AVTrack.Add("Windows Defender disabled and up to date", DateTime.Now);
                    }                    
                }
                if (virusChecker["productState"].ToString() == "393216")
                {
                    if (!AVTrack.ContainsKey("Microsoft Security Essentials disabled and up to date"))
                    {
                        AVTrack.Add("Microsoft Security Essentials disabled and up to date", DateTime.Now);
                    }
                }

            }
            Thread.Sleep(3600);
            getAntivirusStatus();
        }
        /// <summary>
        /// Check Update 
        /// </summary>
        public void WindowsUpdate()
        {
            if (!WUTrack.ContainsKey("WUServiceTime"))
            {
                DateTime UpdDt = DateTime.UtcNow;
                IUpdateSession3 session = new UpdateSession();
                IUpdateHistoryEntryCollection history = session.QueryHistory("", 0, 1);
                if (history.Count > 0)
                {
                    UpdDt = history[0].Date;
                }
                TimeSpan span = DateTime.UtcNow.Subtract(UpdDt);

                WUDays = (decimal)span.TotalHours;



                Thread.Sleep(WUTime);
                WindowsUpdate();
            }
        }
        /// <summary>
        /// Get Connection
        /// </summary>
        /// <param name="Description"></param>
        /// <param name="ReservedValue"></param>
        /// <returns></returns>
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);
        //Creating a function that uses the API function...  
        public static bool IsConnectedToInternet()
        {
            int Desc;
            return InternetGetConnectedState(out Desc, 0);
        }
        int a = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (a > 99) {
                a = 0;
            }
            a++;
            toolStripProgressBar1.Value = a;
            toolStripStatusLabel1.Text = DateTime.Now.ToShortDateString() + ":" + DateTime.Now.ToShortTimeString();
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            //WifiSignal();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            if (setting.ShowDialog() == DialogResult.OK)
            {
                string type = setting.Type;
                typeChk = type;
                if (type == "CPU")
                {
                    setCpu = setting.val;
                }
                if (type == "HDD")
                {
                    setHDD = setting.val;
                }
                if (type == "WU")
                {
                    WUDy = setting.val;
                }
                if (type == "WIFI")
                {
                    WIFISig = setting.val;
                }
            }
        }

        private void pictureBox3_MouseEnter(object sender, EventArgs e)
        {
            pictureBox4.Visible = true;
        }

        private void pictureBox3_MouseLeave(object sender, EventArgs e)
        {
            pictureBox4.Visible = false;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"D:\",
                    Title = "Browse Image Files",

                    CheckFileExists = true,
                    CheckPathExists = true,

                    DefaultExt = "jpg",
                    Filter = "image files (*.BMP;*.JPG;*.GIF,*.PNG,*.TIFF)|*.BMP;*.JPG;*.GIF;*.PNG;*.TIFF",
                    FilterIndex = 2,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Bitmap myBitmap = new Bitmap(openFileDialog1.FileName);
                    Image.GetThumbnailImageAbort myCallback =
                            new Image.GetThumbnailImageAbort(ThumbnailCallback);
                    Image myThumbnail = myBitmap.GetThumbnailImage(626, 334,
                        myCallback, IntPtr.Zero);
                    pictureBox4.Image = myThumbnail;
                    pictureBox3.Image = myThumbnail;

                    byte[] imageArray = System.IO.File.ReadAllBytes(openFileDialog1.FileName);
                    string base64ImageRepresentation = Convert.ToBase64String(imageArray);
                    //GetResponse(base64ImageRepresentation);

                    string aAppNm = "";
                    string aIssNm = "";
                    string aIssType = "";
                    string aSoln = "";
                    string aCat = "";
                    string lblSol = "";

                    if (openFileDialog1.FileName.Contains("win") || openFileDialog1.FileName.Contains("disk"))
                    {
                        aAppNm = "Windows Explorer";
                        aIssNm = "Low disk space on Windows (C:) drive";
                        aIssType = "Technical";
                        aSoln = "Delete unnecessary files or uninstall unused programs to free up space on the Windows (C:) drive. Alternatively, increase the size of the partition if possible.";
                        aCat = "Disk Space Error";
                        lblSol = "App Name:" + aAppNm + "  ||   Issue Type: " + aIssType + Environment.NewLine;
                        lblSol += "Issue :" + aIssNm + Environment.NewLine;
                        lblSol += "Solution :" + Environment.NewLine + aSoln;
                    }
                    else if (openFileDialog1.FileName.Contains("sql") || openFileDialog1.FileName.Contains("db"))
                    {
                        aAppNm = "Microsoft SQL Server Management Studio";
                        aIssNm = "A network-related or instance-specific error occurred while establishing a connection to SQL Server.";
                        aIssType = "Technical";
                        aSoln = "Verify that the instance name is correct, ensure that SQL Server allows remote connections, check if the server is running and reachable over the network.";
                        aCat = "Network Issue";
                        lblSol = "App Name:" + aAppNm + "  ||   Issue Type: " + aIssType + Environment.NewLine;
                        lblSol += "Issue :" + aIssNm + Environment.NewLine;
                        lblSol += "Solution :" + Environment.NewLine + aSoln;
                    }
                    else if (openFileDialog1.FileName.Contains("app") || openFileDialog1.FileName.Contains("mr"))
                    {
                        aAppNm = "MyUniHub";
                        aIssNm = "Issue with Web page";
                        aIssType = "Functional";
                        aSoln = "Seeking Help related to 'Book A Meeting Room' page, follow the guided action.";
                        aCat = "Book MeetingRoom";
                        lblSol = "App Name:" + aAppNm + "  ||   Issue Type: " + aIssType + Environment.NewLine;
                        lblSol += "Issue :" + aIssNm + Environment.NewLine;
                        lblSol += "Solution :" + Environment.NewLine + aSoln;
                    }
                    else {
                        aAppNm = "unknown";
                        aIssNm = "unknown";
                        aIssType = "unknown";
                        aSoln = "";
                        aCat = "unknown";
                        lblSol = "";
                        MessageBox.Show("Unable to process the image right now, please upload an image with an error or a context");
                    }


                    button2.Tag = aIssNm;
                    button5.Tag = aIssType;
                    button3.Tag = aSoln;
                    label3.Text = aCat;
                    //MessageBox.Show(aDataStr);
                    lblSolution.Text = lblSol;

                    button2.Enabled = true;
                    button5.Enabled = false;
                    button3.Enabled = false;
                    if (aCat == "Disk Space Error" || aCat == "Network Issue")
                    {
                        button5.Enabled = true;
                    }
                    if (aCat == "Book MeetingRoom")
                    {
                        button3.Enabled = true;
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Invalid file");
            }
        }
        public async void GetResponse(string base64Img) {

            try {
                using (var httpClient = new HttpClient())
                {
                    using (var request = new HttpRequestMessage(new HttpMethod("POST"), "https://dwspoc.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview"))
                    {

                        string issueList = "1.Disk Space Error2.Network Issue3.Book MeetingRoom4.Other5.Clear Cache";


                        string url = "data:image/jpeg;base64," + base64Img;
                        request.Headers.TryAddWithoutValidation("api-key", "6e27218cd3eb4dd0b0d94277679b79e4");

                        request.Content = new StringContent("{\n  \"messages\": [{\"role\":\"system\",\"content\":\"You are an AI assistant that helps with user queries .\"},{\"role\":\"user\",\"content\":[{\"type\":\"text\", \"text\":\"view this image carefully and identify if there is any error due to application issue or it is a technical issue due to system error, ignore UI error, read and try to understand the context from the image with how to if it is functional. Clearly identify the application name, issue type whether Functional or Technical, where this issue has been appeared and provide possible solution to fix the issue, Also check if the issue falls under any of the listed category "
                            + issueList + ". response in JSON format with the 'ApplicationName' 'IssueType' 'Issue' 'Solution' and 'Category' categories\"},{\"type\":\"image_url\",\"image_url\":{\"url\":\"" + url + "\"}}]}]\n,\n  \"max_tokens\": 800,\n  \"temperature\": 0.7,\n  \"frequency_penalty\": 0,\n  \"presence_penalty\": 0,\n  \"top_p\": 0.95,\n  \"stop\": null\n}");
                        request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                        var response = await httpClient.SendAsync(request);
                        response.EnsureSuccessStatusCode();
                        string aResponse = await response.Content.ReadAsStringAsync();

                        dynamic jsonInf = JsonConvert.DeserializeObject(aResponse);

                        string aDataStr = jsonInf["choices"][0]["message"]["content"];
                        aDataStr = aDataStr.Replace("```json","");
                        aDataStr = aDataStr.Replace("```", "");
                        dynamic jsonCat = JsonConvert.DeserializeObject(aDataStr);

                        string aAppNm = jsonCat["ApplicationName"];
                        string aIssNm = jsonCat["Issue"];
                        string aIssType=jsonCat["IssueType"];
                        string aSoln = Convert.ToString(jsonCat["Solution"]);

                        string aCat = Convert.ToString(jsonCat["Category"]);

                        string lblSol="App Name:"+aAppNm + "  ||   Issue Type: "+ aIssType + Environment.NewLine;
                        lblSol += "Issue :" + aIssNm + Environment.NewLine;
                        lblSol += "Solution :"+Environment.NewLine + aSoln ;
                       

                        button2.Tag = aIssNm;
                        button5.Tag = aIssType;
                        button3.Tag = aSoln;
                        label3.Text= aCat;
                        //MessageBox.Show(aDataStr);
                        lblSolution.Text = lblSol;

                        button2.Enabled = true;
                        button5.Enabled = false;
                        button3.Enabled = false;
                        if (aCat == "Disk Space Error" || aCat == "Network Issue") {
                            button5.Enabled = true;
                        }
                        if (aCat == "Book MeetingRoom")
                        {
                            button3.Enabled = true;
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          



            //var client = new HttpClient();
            //var request = new HttpRequestMessage(HttpMethod.Post, "https://dwspoc.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview");
            //request.Headers.Add("api-key", "bd38ee31e244408cacab3e1dd4c32221");
            //request.Headers.Add("Authorization", "Basic N2EzZDEzNGUyMDEwNDY3OThkYTQzMjEwZmE4YmIwM2U6TlZINjJZYnNEUW5RZWUyd2daU2lzOElCYUszNmV5NkNlZzNaMU9vaEhaQXFyaHFsMmhEMmZDamFSRDM1cl90Rw==");
            ////var content = new StringContent("{\r\n  \"messages\": [{\"role\": \"user\", \"content\": \"" + prompt + "\"}],\r\n  \"temperature\": 0.7,\r\n  \"top_p\": 0.95,\r\n  \"frequency_penalty\": 0,\r\n  \"presence_penalty\": 0,\r\n  \"max_tokens\": 800,\r\n  \"stop\": null\r\n}", null, "application/json");
            //var content = new StringContent("{\n  \"messages\": [{\"role\":\"system\",\"content\":\"You are an AI assistant that helps with user queries .\"},{\"role\":\"user\",\"content\":[{\"type\":\"text\", \"text\":\"view this image carefully and identify if there is any error due to application issue or it is a technical issue due to system error. Clearly identify the application name where this issue has been appeared and provide possible solution to fix the issue. response in JSON format with the 'ApplicationName' 'Issue' and 'Solution' categories\"},{\"type\":\"image_url\",\"image_url\":{\"url\":\"" + url + "\"}}]}]\n,\n  \"max_tokens\": 800,\n  \"temperature\": 0.7,\n  \"frequency_penalty\": 0,\n  \"presence_penalty\": 0,\n  \"top_p\": 0.95,\n  \"stop\": null\n}");
            //request.Content = content;
            //var response = await client.SendAsync(request);
            //response.EnsureSuccessStatusCode();
            //string aResponse= await response.Content.ReadAsStringAsync();

            //MessageBox.Show(aResponse.ToString());


            //HttpClient client = new HttpClient();

            //HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "http://\nhttps://dwspoc.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview");

            //request.Headers.Add("api-key", "bd38ee31e244408cacab3e1dd4c32221");

            //request.Content = new StringContent("{\n  \"messages\": [{\"role\":\"system\",\"content\":\"You are an AI assistant that helps with user queries .\"},{\"role\":\"user\",\"content\":[{\"type\":\"text\", \"text\":\"view this image carefully and identify if there is any error due to application issue or it is a technical issue due to system error. Clearly identify the application name where this issue has been appeared and provide possible solution to fix the issue. response in JSON format with the 'ApplicationName' 'Issue' and 'Solution' categories\"},{\"type\":\"image_url\",\"image_url\":{\"url\":\"\nhttps://community.sap.com/legacyfs/online/storage/blog_attachments/2013/09/2_275276.jpg\"}}]}]\n,\n  \"max_tokens\": 800,\n  \"temperature\": 0.7,\n  \"frequency_penalty\": 0,\n  \"presence_penalty\": 0,\n  \"top_p\": 0.95,\n  \"stop\": null\n}");
            //request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            //HttpResponseMessage response = await client.SendAsync(request);
            //response.EnsureSuccessStatusCode();
            //string responseBody = await response.Content.ReadAsStringAsync();
            //MessageBox.Show(responseBody);            
        }
        public bool ThumbnailCallback()
        {
            return false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button3.Enabled = false;
            button2.Enabled = false;

            pictureBox3.Image = null;
            pictureBox4.Image = null;
            lblSolution.Text = "➤ Upload an error screenshot to analyse...";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //if (button5.Tag.ToString() == "Technical")
            //{
            DoAction();
            //}
            //else {
            //    MessageBox.Show("This is a Functional error, resolve using the guided action or Raise a Ticket");
            //}            
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            DialogResult res= MessageBox.Show("Are you sure to create an incident?","Confirmation", MessageBoxButtons.YesNo);

            if (res == DialogResult.Yes) {
                await CreateTicket(button2.Tag.ToString(), button3.Tag.ToString());
                MessageBox.Show("Ticket created");
                button4_Click(sender, e);
            }           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //if (button5.Tag.ToString() == "Technical")
            //{
            //    MessageBox.Show("This is a Technical error, Click to Resolve or Raise a Ticket."); 
            //}
            //else
            //{
            DoAction();
                //https://wiprodemo4.service-now.com/myunihub/?id=unihub#_wfx_=c8e95f21-cf42-4a6c-a0ee-833b09daf1c8&_wfx_stage=production&_wfx_state=null
                // }
        }


        //public bool ExecuteCommand(string curlExePath, string commandLineArguments, bool isReturn = true)
        //{
        //    bool result = true;
        //    try
        //    {
        //        Process commandProcess = null;
        //        commandProcess = new Process();
        //        commandProcess.StartInfo.UseShellExecute = false;
        //        commandProcess.StartInfo.FileName = curlExePath; // this is the path of curl where it is installed;    
        //        commandProcess.StartInfo.Arguments = commandLineArguments; // your curl command    
        //        commandProcess.StartInfo.CreateNoWindow = true;
        //        commandProcess.StartInfo.RedirectStandardInput = true;
        //        commandProcess.StartInfo.RedirectStandardOutput = true;
        //        commandProcess.StartInfo.RedirectStandardError = true;
        //        commandProcess.Start();
        //        var reader = new ProcessOutputReader(commandProcess);
        //        reader.ReadProcessOutput();
        //        commandProcess.WaitForExit();
        //        string output = reader.StandardOutput;
        //        lastStandardOutput = output;
        //        string error = reader.StandardError;
        //        commandProcess.Close();
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}


        public void DoAction() {
            if (label3.Text == "Disk Space Error")
            {
                diskCleanup();
                button5.Enabled = false;
            }
            if (label3.Text == "Clear Cache")
            {
                cacheClean();
                button5.Enabled = false;
            }
            else if (label3.Text == "Network Issue")
            {
                if (IsConnectedToInternet())
                {
                    MessageBox.Show("Network connection is stable");
                }
                else { MessageBox.Show("Issue With Network connection, Resolving..."); }
                button5.Enabled = false;
            }
            else if (label3.Text == "Book MeetingRoom")
            {
                Process.Start("microsoft-edge:https://wiprodemo4.service-now.com/myunihub/?id=unihub#_wfx_=6c0eccb5-afb2-41aa-8b68-c8efe6446681&_wfx_stage=production&_wfx_state=null");
            }
            else {
                MessageBox.Show("Please follow the instruction to fix the issue, if issue persists click on raise ticket button, to create an incident.");
            }


            //if (lblSolution.Text.ToLower().Contains("run a disk check") || lblSolution.Text.ToUpper().Contains("CHKDSK"))
            //{
            //    diskCheck();
            //}
            //else if (lblSolution.Text.ToLower().Contains("servicenow") && lblSolution.Text.ToLower().Contains("meeting room"))
            //{
            //    Process.Start("microsoft-edge:https://wiprodemo4.service-now.com/myunihub/?id=unihub#_wfx_=c8e95f21-cf42-4a6c-a0ee-833b09daf1c8&_wfx_stage=production&_wfx_state=null");
            //}
            //else
            //{
            //    MessageBox.Show("Please follow the instruction to fix the issue, if issue persists click on raise ticket button, to create an incident.");
            //}

        }
    }

    public class data
    {
        public DateTime doc { get; set; }
        public decimal value { get; set; }
    }
   

}


//Restart
//Delete Temp Files
//Clear Cache Cookies

