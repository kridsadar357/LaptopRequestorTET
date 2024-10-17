using System;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Runtime.InteropServices;
using MaterialSkin;
using MaterialSkin.Controls;
using System.Net.Http;
using System.Drawing;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Collections.Generic;
using Microsoft.Win32;
using IWshRuntimeLibrary;


namespace LaptopRequestorTET
{
    internal static class Program
    {
        private static NotifyIcon trayIcon;
        private static Thread monitorThread;
        private static Form dummyForm;
        private static String computerName;
        private static List<string> validSSIDs = new List<string> { "TAKANE_WiFi", "TAKANE_WH" };

        [DllImport("wlanapi.dll", SetLastError = true)]
        public static extern uint WlanOpenHandle(uint dwClientVersion, IntPtr pReserved, out uint pdwNegotiatedVersion, out IntPtr phClientHandle);

        [DllImport("wlanapi.dll", SetLastError = true)]
        public static extern uint WlanEnumInterfaces(IntPtr hClientHandle, IntPtr pReserved, out IntPtr ppInterfaceList);

        [DllImport("wlanapi.dll", SetLastError = true)]
        public static extern uint WlanQueryInterface(IntPtr hClientHandle, ref Guid interfaceGuid, WLAN_INTF_OPCODE opCode, IntPtr pReserved, out int pdwDataSize, ref IntPtr ppData, IntPtr pWlanOpcodeValueType);

        [DllImport("wlanapi.dll", SetLastError = true)]
        public static extern void WlanFreeMemory(IntPtr pMemory);

        [DllImport("wlanapi.dll", SetLastError = true)]
        public static extern uint WlanCloseHandle(IntPtr hClientHandle, IntPtr pReserved);

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hwnd, uint msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr DefWindowProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct WLAN_INTERFACE_INFO
        {
            public Guid InterfaceGuid;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string strInterfaceDescription;
            public WLAN_INTERFACE_STATE isState;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WLAN_INTERFACE_INFO_LIST
        {
            public uint dwNumberOfItems;
            public uint dwIndex;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public WLAN_INTERFACE_INFO[] InterfaceInfo;
        }

        public enum WLAN_INTERFACE_STATE
        {
            wlan_interface_state_not_ready = 0,
            wlan_interface_state_connected,
            wlan_interface_state_ad_hoc_network_formed,
            wlan_interface_state_disconnecting,
            wlan_interface_state_disconnected,
            wlan_interface_state_associating,
            wlan_interface_state_discovering,
            wlan_interface_state_authenticating
        }

        public enum WLAN_INTF_OPCODE
        {
            wlan_intf_opcode_current_connection = 7
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct WLAN_CONNECTION_ATTRIBUTES
        {
            public WLAN_INTERFACE_STATE isState;
            public WLAN_CONNECTION_MODE wlanConnectionMode;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string strProfileName;
            public WLAN_ASSOCIATION_ATTRIBUTES wlanAssociationAttributes;
            public WLAN_SECURITY_ATTRIBUTES wlanSecurityAttributes;
        }

        public enum WLAN_CONNECTION_MODE
        {
            wlan_connection_mode_profile = 0,
            wlan_connection_mode_temporary_profile,
            wlan_connection_mode_discovery_secure,
            wlan_connection_mode_discovery_unsecure,
            wlan_connection_mode_auto,
            wlan_connection_mode_invalid
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WLAN_ASSOCIATION_ATTRIBUTES
        {
            public DOT11_SSID dot11Ssid;
            public DOT11_BSS_TYPE dot11BssType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
            public byte[] dot11Bssid;
            public DOT11_PHY_TYPE dot11PhyType;
            public uint uDot11PhyIndex;
            public uint wlanSignalQuality;
            public uint ulRxRate;
            public uint ulTxRate;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DOT11_SSID
        {
            public uint uSSIDLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
            public byte[] ucSSID;
        }

        public enum DOT11_BSS_TYPE
        {
            dot11_BSS_type_infrastructure = 1,
            dot11_BSS_type_independent = 2,
            dot11_BSS_type_any = 3
        }

        public enum DOT11_PHY_TYPE : uint
        {
            dot11_phy_type_unknown = 0,
            dot11_phy_type_any = 0,
            dot11_phy_type_fhss = 1,
            dot11_phy_type_dsss = 2,
            dot11_phy_type_irbaseband = 3,
            dot11_phy_type_ofdm = 4,
            dot11_phy_type_hrdsss = 5,
            dot11_phy_type_erp = 6,
            dot11_phy_type_ht = 7,
            dot11_phy_type_vht = 8,
            dot11_phy_type_IHV_start = 0x80000000,
            dot11_phy_type_IHV_end = 0xffffffff
        }

        public struct WLAN_SECURITY_ATTRIBUTES
        {
            [MarshalAs(UnmanagedType.Bool)]
            public bool bSecurityEnabled;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bOneXEnabled;
            public DOT11_AUTH_ALGORITHM dot11AuthAlgorithm;
            public DOT11_CIPHER_ALGORITHM dot11CipherAlgorithm;
        }

        public enum DOT11_AUTH_ALGORITHM : uint
        {
            DOT11_AUTH_ALGO_80211_OPEN = 1,
            DOT11_AUTH_ALGO_80211_SHARED_KEY = 2,
            DOT11_AUTH_ALGO_WPA = 3,
            DOT11_AUTH_ALGO_WPA_PSK = 4,
            DOT11_AUTH_ALGO_WPA_NONE = 5,
            DOT11_AUTH_ALGO_RSNA = 6,
            DOT11_AUTH_ALGO_RSNA_PSK = 7,
            DOT11_AUTH_ALGO_IHV_START = 0x80000000,
            DOT11_AUTH_ALGO_IHV_END = 0xffffffff
        }

        public enum DOT11_CIPHER_ALGORITHM : uint
        {
            DOT11_CIPHER_ALGO_NONE = 0x00,
            DOT11_CIPHER_ALGO_WEP40 = 0x01,
            DOT11_CIPHER_ALGO_TKIP = 0x02,
            DOT11_CIPHER_ALGO_CCMP = 0x04,
            DOT11_CIPHER_ALGO_WEP104 = 0x05,
            DOT11_CIPHER_ALGO_WPA_USE_GROUP = 0x100,
            DOT11_CIPHER_ALGO_RSN_USE_GROUP = 0x100,
            DOT11_CIPHER_ALGO_WEP = 0x101,
            DOT11_CIPHER_ALGO_IHV_START = 0x80000000,
            DOT11_CIPHER_ALGO_IHV_END = 0xffffffff
        }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            CreateStartupShortcut();
            computerName = Environment.MachineName;

            dummyForm = new Form
            {
                ShowInTaskbar = false,
                Opacity = 0,
                WindowState = FormWindowState.Minimized
            };

            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.laptoplock_83c_icon,
                Visible = true,
                Text = "TET Asset Control"
            };


            monitorThread = new Thread(MonitorWiFi);
            monitorThread.Start();
            Application.Run(dummyForm);
        }

        static void AddApplicationToStartup()
        {
            try
            {
                string appName = "TETAssetControlOutside";
                string appPath = Application.ExecutablePath;

                RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
                if (registryKey != null)
                {
                    registryKey.SetValue(appName, appPath);
                    registryKey.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding application to startup: {ex.Message}");
            }
        }

        static void CreateStartupShortcut()
        {
            try
            {
                string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string shortcutPath = Path.Combine(startupFolderPath, "TETAssetControlOutside.lnk");

                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.Description = "TET Asset Control Outside Application";
                shortcut.TargetPath = Application.ExecutablePath;
                shortcut.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating startup shortcut: {ex.Message}");
            }
        }

        static void MonitorWiFi()
        {
            while (true)
            {
                string ssid = GetCurrentSSID();
                if (validSSIDs.Contains(ssid) == false)
                {
                    if (dummyForm.IsHandleCreated)
                    {
                        dummyForm.Invoke((MethodInvoker)delegate
                        {
                            CheckActivationStatus();
                        });
                    }
                }
                Thread.Sleep(5000);
            }
        }

        static string GetCurrentSSID()
        {
            IntPtr clientHandle = IntPtr.Zero;
            IntPtr interfaceList = IntPtr.Zero;
            try
            {
                uint negotiatedVersion;
                uint result = WlanOpenHandle(2, IntPtr.Zero, out negotiatedVersion, out clientHandle);
                if (result != 0)
                    return string.Empty;

                result = WlanEnumInterfaces(clientHandle, IntPtr.Zero, out interfaceList);
                if (result != 0)
                    return string.Empty;

                WLAN_INTERFACE_INFO_LIST listInfo = (WLAN_INTERFACE_INFO_LIST)Marshal.PtrToStructure(interfaceList, typeof(WLAN_INTERFACE_INFO_LIST));

                IntPtr currentPtr = (IntPtr)((long)interfaceList + Marshal.OffsetOf(typeof(WLAN_INTERFACE_INFO_LIST), "InterfaceInfo").ToInt64());

                for (int i = 0; i < listInfo.dwNumberOfItems; i++)
                {
                    WLAN_INTERFACE_INFO interfaceInfo = (WLAN_INTERFACE_INFO)Marshal.PtrToStructure(currentPtr, typeof(WLAN_INTERFACE_INFO));

                    IntPtr connAttrPtr = IntPtr.Zero;
                    int connAttrSize = 0;
                    result = WlanQueryInterface(clientHandle, ref interfaceInfo.InterfaceGuid, WLAN_INTF_OPCODE.wlan_intf_opcode_current_connection, IntPtr.Zero, out connAttrSize, ref connAttrPtr, IntPtr.Zero);

                    if (result == 0)
                    {
                        WLAN_CONNECTION_ATTRIBUTES connAttr = (WLAN_CONNECTION_ATTRIBUTES)Marshal.PtrToStructure(connAttrPtr, typeof(WLAN_CONNECTION_ATTRIBUTES));
                        if (connAttr.isState == WLAN_INTERFACE_STATE.wlan_interface_state_connected)
                        {
                            string ssid = new string(connAttr.wlanAssociationAttributes.dot11Ssid.ucSSID
                                .Take((int)connAttr.wlanAssociationAttributes.dot11Ssid.uSSIDLength)
                                .Select(b => (char)b)
                                .ToArray());
                            return ssid;
                        }
                    }

                    currentPtr = IntPtr.Add(currentPtr, Marshal.SizeOf(typeof(WLAN_INTERFACE_INFO)));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting SSID: {ex.Message}");
            }
            finally
            {
                if (interfaceList != IntPtr.Zero)
                    WlanFreeMemory(interfaceList);
                if (clientHandle != IntPtr.Zero)
                    WlanCloseHandle(clientHandle, IntPtr.Zero);
            }
            return string.Empty;
        }

        static void CheckActivationStatus()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string filePath = Path.Combine(appDataPath, "TET Asset Control Outside", "activation_state.txt");

                if (System.IO.File.Exists(filePath))
                {
                    string activationData = System.IO.File.ReadAllText(filePath);
                    string[] parts = activationData.Split('-');
                    if (parts.Length >= 2)
                    {
                        string datePart = parts[1];
                        DateTime activationDate;
                        if (DateTime.TryParseExact(datePart, "yyyyMMddHHmmss", null, System.Globalization.DateTimeStyles.None, out activationDate))
                        {
                            DateTime currentDate = DateTime.Now;
                            TimeSpan difference = currentDate - activationDate;
                            if (difference.TotalDays < 1)
                            {
                                Console.WriteLine("Activation is still valid.");
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading activation state: {ex.Message}");
            }
            ShowInputForm(null);
        }
        public class MaterialFormWithNoMove : MaterialForm
        {
            protected override void WndProc(ref Message message)
            {
                const int WM_NCRBUTTONDOWN = 0x00A4;
                const int WM_NCRBUTTONUP = 0x00A5;
                const int WM_NCRBUTTONDBLCLK = 0x00A6;
                const int WM_SYSCOMMAND = 0x0112;
                const int SC_MOVE = 0xF010;
                const int HTCAPTION = 0x2;

                switch (message.Msg)
                {
                    case WM_NCRBUTTONDOWN:
                    case WM_NCRBUTTONUP:
                    case WM_NCRBUTTONDBLCLK:
                        return; // Ignore all right-click messages in the non-client area

                    case WM_SYSCOMMAND:
                        int command = message.WParam.ToInt32() & 0xfff0;
                        if (command == SC_MOVE)
                            return; // Ignore the move command
                        break;

                    case 0x84: // WM_NCHITTEST
                        message.Result = (IntPtr)HTCAPTION; // Prevent dragging by disabling caption hit test
                        return;
                }

                base.WndProc(ref message);
            }
        }

        static async void ShowInputForm(Form dummyForm)
        {
            var skinManager = MaterialSkinManager.Instance;
            skinManager.AddFormToManage(new MaterialFormWithNoMove());
            skinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            skinManager.ColorScheme = new ColorScheme(Primary.Blue600, Primary.Blue700, Primary.Blue200, Accent.LightBlue200, TextShade.WHITE);

            using (MaterialFormWithNoMove inputForm = new MaterialFormWithNoMove())
            {
                inputForm.Width = 400;
                inputForm.Height = 300;
                inputForm.MinimumSize = new Size(400, 300);
                inputForm.MaximumSize = new Size(400, 300);
                inputForm.MinimizeBox = false;
                inputForm.MaximizeBox = false;
                inputForm.TopMost = true;
                inputForm.ShowInTaskbar = false;
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputForm.ControlBox = false;
                inputForm.Text = "Laptop activation for outside";
                inputForm.StartPosition = FormStartPosition.CenterScreen;
                inputForm.Sizable = false;

                PictureBox logo = new PictureBox()
                {
                    Image = Properties.Resources.logo,
                    Left = inputForm.Width / 2 - 40,
                    Top = 10,
                    Width = 70,
                    Height = 60,
                    SizeMode = PictureBoxSizeMode.StretchImage
                };
                MaterialLabel label = new MaterialLabel() { Left = 80, Top = 80, Text = "Laptop activation to outside TET :", Width = 300 };
                MaterialTextBox textBox1 = new MaterialTextBox()
                {
                    Location = new System.Drawing.Point(80, 120),

                    Width = 50,
                    Height = 50,
                    Font = new System.Drawing.Font("Roboto", 11F, System.Drawing.FontStyle.Bold),
                    MaxLength = 1
                };
                MaterialTextBox textBox2 = new MaterialTextBox()
                {
                    Location = new System.Drawing.Point(140, 120),
                    Width = 50,
                    Height = 50,
                    Font = new System.Drawing.Font("Roboto", 11F, System.Drawing.FontStyle.Bold),
                    MaxLength = 1
                };
                MaterialTextBox textBox3 = new MaterialTextBox()
                {
                    Location = new System.Drawing.Point(205, 120),
                    Width = 50,
                    Height = 50,
                    Font = new System.Drawing.Font("Roboto", 11F, System.Drawing.FontStyle.Bold),
                    MaxLength = 1
                };
                MaterialTextBox textBox4 = new MaterialTextBox()
                {
                    Location = new System.Drawing.Point(265, 120),
                    Width = 50,
                    Height = 50,
                    Font = new System.Drawing.Font("Roboto", 11F, System.Drawing.FontStyle.Bold),
                    MaxLength = 1
                };
                MaterialButton confirmButton = new MaterialButton() { Text = "Confirm", Location = new System.Drawing.Point(150, 200), Width = 300 };
                MaterialLabel errorLabel = new MaterialLabel()
                {
                    Location = new System.Drawing.Point(80, 250),
                    ForeColor = System.Drawing.Color.Red,
                    Text = "ERROR: Code cannot be empty!",
                    Visible = false,
                    Font = new System.Drawing.Font("Roboto", 11F, System.Drawing.FontStyle.Bold),
                    AutoSize = true
                };

                confirmButton.Click += async (sender, e) =>
                {
                    string code = textBox1.Text + textBox2.Text + textBox3.Text + textBox4.Text;
                    confirmButton.Enabled = false;
                    if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                 string.IsNullOrWhiteSpace(textBox2.Text) ||
                 string.IsNullOrWhiteSpace(textBox3.Text) ||
                 string.IsNullOrWhiteSpace(textBox4.Text))
                    {
                        errorLabel.Visible = true;
                        errorLabel.Text = "ERROR: All codes must be entered!";
                    }
                    else
                    {
                        errorLabel.Visible = true;
                        errorLabel.Text = "Loading...";
                        string secretKey = "takane304";
                        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string token = code + "-" + timestamp + "-" + secretKey;
                        string comName = Environment.MachineName;
                        using (HttpClient client = new HttpClient())
                        {
                            try
                            {
                                var values = new { code = token, pc = comName };
                                var content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(values), System.Text.Encoding.UTF8, "application/json");

                                HttpResponseMessage response = await client.PostAsync("http://takane304.ruijieddns.com:8080/app/active_outside.php", content);
                                string responseString = await response.Content.ReadAsStringAsync();

                                dynamic jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject(responseString);
                                //Console.WriteLine(jsonResponse);
                                if (jsonResponse != null && jsonResponse.st == 10)
                                {
                                    inputForm.DialogResult = DialogResult.OK;
                                    errorLabel.Visible = false;
                                    string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                                    string filePath = Path.Combine(appDataPath, "TET Asset Control Outside", "activation_state.txt");
                                    //MessageBox.Show(filePath);
                                    Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                                    System.IO.File.WriteAllText(filePath, token);
                                }
                                else
                                {
                                    textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = "";
                                    errorLabel.Text = "ERROR: Incorrect code.";
                                    textBox1.Focus();
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                errorLabel.Text = $"ERROR: {ex.Message}";
                            }
                        }
                        confirmButton.Enabled = true;
                    }
                };
                inputForm.Controls.Add(logo);
                inputForm.Controls.Add(label);
                inputForm.Controls.Add(textBox1);
                inputForm.Controls.Add(textBox2);
                inputForm.Controls.Add(textBox3);
                inputForm.Controls.Add(textBox4);
                inputForm.Controls.Add(confirmButton);
                inputForm.Controls.Add(errorLabel);
                inputForm.AcceptButton = confirmButton;

                inputForm.MouseClick += DisabledMovement;

                textBox1.TextChanged += (s, e) => MoveToNextTextBox(textBox1, textBox2);
                textBox2.TextChanged += (s, e) => MoveToNextTextBox(textBox2, textBox3);
                textBox3.TextChanged += (s, e) => MoveToNextTextBox(textBox3, textBox4);
                textBox4.TextChanged += (s, e) => MoveToNextTextBox(textBox4, null);

                textBox2.KeyDown += (s, e) => MoveToPreviousTextBoxOnBackspace(e, textBox1, textBox2);
                textBox3.KeyDown += (s, e) => MoveToPreviousTextBoxOnBackspace(e, textBox2, textBox3);
                textBox4.KeyDown += (s, e) => MoveToPreviousTextBoxOnBackspace(e, textBox3, textBox4);


                void DisabledMovement(Object sender, MouseEventArgs e)
                {
                    ReleaseCapture();
                    SendMessage(inputForm.Handle, 0x00A1, 2, 1);
                }

                void MoveToNextTextBox(MaterialTextBox currentTextBox, MaterialTextBox nextTextBox)
                {
                    if (currentTextBox.Text.Length == currentTextBox.MaxLength)
                    {
                        nextTextBox?.Focus();
                    }
                }

                void MoveToPreviousTextBoxOnBackspace(KeyEventArgs e, MaterialTextBox previousTextBox, MaterialTextBox currentTextBox)
                {
                    if (e.KeyCode == Keys.Back && currentTextBox.Text.Length == 0)
                    {
                        previousTextBox.Focus();
                        e.Handled = true;
                    }
                }

                inputForm.FormClosing += (sender, e) =>
                {
                    if (inputForm.DialogResult != DialogResult.OK)
                    {
                        e.Cancel = true;
                        //MessageBox.Show("You must enter a valid code before closing.");
                    }
                };

                inputForm.Shown += (s, e) =>
                {
                    Thread monitorThread = new Thread(() =>
                    {
                        while (true)
                        {
                            if (inputForm.IsHandleCreated)
                            {
                                string ssid = GetCurrentSSID();
                                if (validSSIDs.Contains(ssid))
                                {
                                    inputForm.Invoke((MethodInvoker)delegate
                                    {
                                        if (inputForm.Visible)
                                        {
                                            inputForm.Hide();
                                        }
                                    });
                                    break;
                                }
                            }
                            Thread.Sleep(2000);
                        }
                    });
                    monitorThread.IsBackground = true;
                    monitorThread.Start();
                };

                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    Console.WriteLine($"Code entered: {textBox1.Text}");
                }
            }
        }
    }
}
