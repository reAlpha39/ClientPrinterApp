using Microsoft.Win32;
using System;
using System.IO;
using System.Text;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Drawing;

namespace ClientPrinterApp
{
    public partial class MainForm : Form
    {
        private HttpListener listener;
        private bool isRunning = false;
        private string apiPort = "8085";
        private NotifyIcon trayIcon;
        private bool startupEnabled = false;
        private readonly string appName = "PrinterServiceHelper";

        public MainForm()
        {
            InitializeComponent();
            InitializeTrayIcon();
            CheckStartupStatus();
            txtPort.Text = apiPort;
        }

        private void CheckStartupStatus()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false))
            {
                startupEnabled = key?.GetValue(appName) != null;
                updateToggleBtnAutoStartText();
            }
        }

        private void updateToggleBtnAutoStartText()
        {


            if (startupEnabled)
            {
                btnAutoStart.Text = "Auto Start ON";
            }
            else
            {
                btnAutoStart.Text = "Auto Start OFF";
            }

        }

        private void btnAutoStart_Click(object sender, EventArgs e)
        {
            if (startupEnabled)
            {
                DisableStartup();
            }
            else
            {
                EnableStartup();
            }

            startupEnabled = !startupEnabled;

            updateToggleBtnAutoStartText();

            

        }

        private void EnableStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                // Add the application to the startup list
                key?.SetValue(appName, Application.ExecutablePath);
            }

            MessageBox.Show("Helper added to the startup!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DisableStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                // Remove the application from the startup list
                key?.DeleteValue(appName, false);
            }

            MessageBox.Show("Helper removed from startup!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void InitializeTrayIcon()
        {
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Text = "Printer Service Helper";
            trayIcon.Visible = true;

            // Create context menu for tray icon
            ContextMenu menu = new ContextMenu();
            menu.MenuItems.Add("Show", OnShow);
            menu.MenuItems.Add("Exit", OnExit);
            trayIcon.ContextMenu = menu;

            trayIcon.DoubleClick += (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; };
        }

        private void OnShow(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void OnExit(object sender, EventArgs e)
        {
            StopServer();
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                trayIcon.Visible = false;
                base.OnFormClosing(e);
            }
        }

        private void btnStartStop_Click(object sender, EventArgs e)
        {
            if (isRunning)
            {
                StopServer();
                btnStartStop.Text = "Start Server";
                LogMessage("Server stopped");
                txtPort.Enabled = true;
            }
            else
            {
                apiPort = txtPort.Text;
                StartServer();
                btnStartStop.Text = "Stop Server";
                LogMessage($"Server started on port {apiPort}");
                txtPort.Enabled = false;
            }
        }

        private void StartServer()
        {
            try
            {
                listener = new HttpListener();
                listener.Prefixes.Add($"http://localhost:{apiPort}/");
                listener.Start();
                isRunning = true;

                Task.Run(() => HandleRequests());
            }
            catch (Exception ex)
            {
                LogMessage($"Error starting server: {ex.Message}");
            }
        }

        private void StopServer()
        {
            if (listener != null && listener.IsListening)
            {
                listener.Stop();
                listener.Close();
                isRunning = false;
            }
        }

        private async Task HandleRequests()
        {
            while (isRunning)
            {
                try
                {
                    var context = await listener.GetContextAsync();
                    ProcessRequest(context);
                }
                catch (Exception ex)
                {
                    if (isRunning)
                    {
                        LogMessage($"Error handling request: {ex.Message}");
                    }
                }
            }
        }

        private void ProcessRequest(HttpListenerContext context)
        {
            try
            {
                var request = context.Request;
                var response = context.Response;

                // Enable CORS
                response.Headers.Add("Access-Control-Allow-Origin", "*");
                response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");

                if (request.HttpMethod == "OPTIONS")
                {
                    response.StatusCode = 200;
                    response.Close();
                    return;
                }

                if (request.HttpMethod == "GET" && request.Url.AbsolutePath == "/printers")
                {
                    // Get list of printers and return as JSON
                    var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
                    var printerList = new List<string>();
                    foreach (string printer in printers)
                    {
                        printerList.Add(printer);
                    }

                    SendJsonResponse(response, printerList);
                    LogMessage("Returned list of printers");
                }
                else if (request.HttpMethod == "POST" && request.Url.AbsolutePath == "/print")
                {
                    // Read the print request
                    string requestBody;
                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                    {
                        requestBody = reader.ReadToEnd();
                    }

                    var printData = JsonConvert.DeserializeObject<PrintRequest>(requestBody);

                    // Decode base64 data if provided
                    byte[] rawData;
                    if (printData.IsBase64)
                    {
                        rawData = Convert.FromBase64String(printData.Data);
                    }
                    else
                    {
                        rawData = Encoding.UTF8.GetBytes(printData.Data);
                    }

                    // Print the data
                    string errorMessage = "";
                    bool success = PrintRawData(printData.PrinterName, printData.DocumentName, rawData, out errorMessage);

                    LogMessage($"Print request to {printData.PrinterName}: {(success ? "Success" : "Failed - " + errorMessage)}");

                    var result = new
                    {
                        Success = success,
                        Message = success ? "Print job sent successfully" : errorMessage
                    };

                    SendJsonResponse(response, result);
                }
                else if (request.HttpMethod == "POST" && request.Url.AbsolutePath == "/print-sato")
                {
                    // Read the print request
                    string requestBody;
                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                    {
                        requestBody = reader.ReadToEnd();
                    }

                    var satoRequest = JsonConvert.DeserializeObject<SatoLabelRequest>(requestBody);

                    // Print using the SATO format
                    string errorMessage = "";
                    bool success = SendSatoLabel(satoRequest, out errorMessage);

                    LogMessage($"SATO print request to {satoRequest.PrinterName}: {(success ? "Success" : "Failed - " + errorMessage)}");

                    var result = new
                    {
                        Success = success,
                        Message = success ? "SATO print job sent successfully" : errorMessage
                    };

                    SendJsonResponse(response, result);
                }
                else
                {
                    // Unknown endpoint
                    response.StatusCode = 404;
                    byte[] buffer = Encoding.UTF8.GetBytes("Not found");
                    response.ContentLength64 = buffer.Length;
                    response.OutputStream.Write(buffer, 0, buffer.Length);
                    response.Close();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error processing request: {ex.Message}");
                try
                {
                    var response = context.Response;
                    response.StatusCode = 500;
                    byte[] buffer = Encoding.UTF8.GetBytes($"Server error: {ex.Message}");
                    response.ContentLength64 = buffer.Length;
                    response.OutputStream.Write(buffer, 0, buffer.Length);
                    response.Close();
                }
                catch { /* Ignore errors in error handling */ }
            }
        }

        private bool SendSatoLabel(SatoLabelRequest labelData, out string errorMessage)
        {
            try
            {
                // Define control characters
                char STX = (char)0x02; // Start of Text
                char ESC = (char)0x1b; // Escape
                char ETX = (char)0x03; // End of Text

                // First part of the command (up to logo editing)
                string EditWk = "";

                // Set STX
                EditWk = STX.ToString();

                // Start data transmission
                EditWk += ESC + "A";

                // Set paper size (width=200, height=600)
                EditWk += ESC + "A1V00200H0600";

                // Title editing
                // Position: H=220, V=30, Magnification: 1x1
                EditWk += ESC + "%0" + ESC + "V0030" + ESC + "H0220" + ESC + "P00" + ESC + "L0101";
                EditWk += ESC + "XM " + labelData.Title;

                // Barcode editing
                // Position: H=260, V=60
                EditWk += ESC + "V0060" + ESC + "H0260";
                // Barcode type (CODE39), magnification (2x)
                EditWk += ESC + "D101060*" + labelData.Barcode + "*";

                // Human-readable barcode text
                // Position: H=230, V=130, Magnification: 1x1
                EditWk += ESC + "%0" + ESC + "V00130" + ESC + "H0230" + ESC + "P00" + ESC + "L0101";
                EditWk += ESC + "X22,*" + labelData.Barcode + "*";

                // Convert first part to byte array using Shift-JIS encoding
                byte[] bStData = Encoding.GetEncoding("Shift-JIS").GetBytes(EditWk);

                // Second part of the command (print quantity to end)
                EditWk = ESC + "Q1";

                // End data transmission
                EditWk += ESC + "Z";

                // Set ETX
                EditWk += ETX.ToString();

                // Convert second part to byte array using Shift-JIS encoding
                byte[] bEnData = Encoding.GetEncoding("Shift-JIS").GetBytes(EditWk);

                // Combine byte arrays
                byte[] bPrData = new byte[bStData.Length + bEnData.Length];
                Array.Copy(bStData, 0, bPrData, 0, bStData.Length);
                Array.Copy(bEnData, 0, bPrData, bStData.Length, bEnData.Length);

                // Print the data
                return PrintRawData(labelData.PrinterName, "SATO Label", bPrData, out errorMessage);
            }
            catch (Exception ex)
            {
                errorMessage = $"Exception in SATO label printing: {ex.Message}";
                return false;
            }
        }

        private void SendJsonResponse(HttpListenerResponse response, object data)
        {
            response.ContentType = "application/json";
            string json = JsonConvert.SerializeObject(data);
            byte[] buffer = Encoding.UTF8.GetBytes(json);
            response.ContentLength64 = buffer.Length;
            response.OutputStream.Write(buffer, 0, buffer.Length);
            response.Close();
        }

        private void LogMessage(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(LogMessage), message);
                return;
            }

            txtLog.AppendText($"[{DateTime.Now.ToString("HH:mm:ss")}] {message}{Environment.NewLine}");
            txtLog.ScrollToCaret();
        }

        // Printer API implementations
        #region Printer API

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DOCINFOW
        {
            [MarshalAs(UnmanagedType.LPWStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPWStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPWStr)] public string pDataType;
        }

        [DllImport("winspool.drv", EntryPoint = "OpenPrinterW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool OpenPrinter(string pPrinterName, out IntPtr hPrinter, IntPtr pDefault);

        [DllImport("winspool.drv", EntryPoint = "ClosePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", EntryPoint = "StartDocPrinterW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int level, ref DOCINFOW pDI);

        [DllImport("winspool.drv", EntryPoint = "EndDocPrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", EntryPoint = "WritePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        public bool PrintRawData(string printerName, string documentName, byte[] rawData, out string errorMessage)
        {
            IntPtr hPrinter = IntPtr.Zero;
            errorMessage = string.Empty;

            try
            {
                // Open printer
                if (!OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
                {
                    errorMessage = $"Failed to open printer. Error: {Marshal.GetLastWin32Error()}";
                    return false;
                }

                // Start document
                DOCINFOW di = new DOCINFOW
                {
                    pDocName = documentName,
                    pOutputFile = null,
                    pDataType = null
                };

                if (!StartDocPrinter(hPrinter, 1, ref di))
                {
                    errorMessage = $"Failed to start document. Error: {Marshal.GetLastWin32Error()}";
                    ClosePrinter(hPrinter);
                    return false;
                }

                // Allocate unmanaged memory and copy data
                IntPtr pUnmanagedBytes = Marshal.AllocCoTaskMem(rawData.Length);
                Marshal.Copy(rawData, 0, pUnmanagedBytes, rawData.Length);

                // Write data to printer
                int bytesWritten = 0;
                if (!WritePrinter(hPrinter, pUnmanagedBytes, rawData.Length, out bytesWritten))
                {
                    errorMessage = $"Failed to write to printer. Error: {Marshal.GetLastWin32Error()}";
                    Marshal.FreeCoTaskMem(pUnmanagedBytes);
                    EndDocPrinter(hPrinter);
                    ClosePrinter(hPrinter);
                    return false;
                }

                // Free memory and close printer
                Marshal.FreeCoTaskMem(pUnmanagedBytes);
                EndDocPrinter(hPrinter);
                ClosePrinter(hPrinter);

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"Exception: {ex.Message}";
                if (hPrinter != IntPtr.Zero)
                {
                    ClosePrinter(hPrinter);
                }
                return false;
            }
        }

        #endregion
    }

    public class PrintRequest
    {
        public string PrinterName { get; set; }
        public string DocumentName { get; set; }
        public string Data { get; set; }
        public bool IsBase64 { get; set; }
    }

    public class SatoLabelRequest
    {
        public string PrinterName { get; set; }
        public string Title { get; set; }
        public string Barcode { get; set; }
    }

    // Designer-generated code for the form
    public partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private TextBox txtLog;
        private Button btnStartStop;
        private Button btnAutoStart;
        private Label lblPort;
        private TextBox txtPort;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnStartStop = new System.Windows.Forms.Button();
            this.btnAutoStart = new System.Windows.Forms.Button();
            this.lblPort = new System.Windows.Forms.Label();
            this.txtPort = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtLog
            // 
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Location = new System.Drawing.Point(12, 41);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(460, 208);
            this.txtLog.TabIndex = 0;
            // 
            // btnStartStop
            // 
            this.btnStartStop.Location = new System.Drawing.Point(372, 12);
            this.btnStartStop.Name = "btnStartStop";
            this.btnStartStop.Size = new System.Drawing.Size(100, 23);
            this.btnStartStop.TabIndex = 1;
            this.btnStartStop.Text = "Start Server";
            this.btnStartStop.UseVisualStyleBackColor = true;
            this.btnStartStop.Click += new System.EventHandler(this.btnStartStop_Click);
            //
            // btnAutoStart
            //
            this.btnAutoStart.Location = new System.Drawing.Point(248, 12);
            this.btnAutoStart.Name = "btnAutoStart";
            this.btnAutoStart.Size = new System.Drawing.Size(112, 23);
            this.btnAutoStart.TabIndex = 2;
            this.btnAutoStart.Text = "Auto Start OFF";
            this.btnAutoStart.UseVisualStyleBackColor = true;
            this.btnAutoStart.Click += new System.EventHandler(this.btnAutoStart_Click);
            // 
            // lblPort
            // 
            this.lblPort.AutoSize = true;
            this.lblPort.Location = new System.Drawing.Point(12, 17);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(29, 13);
            this.lblPort.TabIndex = 3;
            this.lblPort.Text = "Port:";
            // 
            // txtPort
            // 
            this.txtPort.Location = new System.Drawing.Point(47, 14);
            this.txtPort.Name = "txtPort";
            this.txtPort.Size = new System.Drawing.Size(100, 20);
            this.txtPort.TabIndex = 4;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 261);
            this.Controls.Add(this.txtPort);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.btnAutoStart);
            this.Controls.Add(this.btnStartStop);
            this.Controls.Add(this.txtLog);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Printer Service Helper";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}