using Fleck;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DesktopAgent
{
    public partial class DesktopAgentForm : Form
    {
        public DesktopAgentForm()
        {
            InitializeComponent();
            SetupServer();
        }

        private void DesktopAgentForm_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
        }

        private void ExecuteActions(IWebSocketConnection webSocket, string message)
        {
            Invoke(new Action(() => txtLog.AppendText(message)));

            try
            {
                var action = JObject.Parse(message);
                var command = action["action"].Value<string>();
                switch (command)
                {
                    //Add your actions in this switch
                    case "WSH":
                        this.RunWSH(webSocket, action);
                        break;
                    case "ExecProgram":
                        this.RunProgram(webSocket, action);
                        break;
                    case "Excel":
                        this.ExcelAction(webSocket, action);
                        break;
                    case "Word":
                        this.WordAction(webSocket, action);
                        break;
                }
            }
            catch
            {
                // Just ignore
            }
        }

        private void RunWSH(IWebSocketConnection webSocket, JObject action)
        {

            var wshAction = action["command"].Value<string>();
            switch (wshAction)
            {
                case "ComputerName":
                    dynamic wsh = Microsoft.VisualBasic.Interaction.CreateObject("WScript.Network");
                    JObject res = new JObject();
                    res["action"] = "wsh";
                    res["result"] = wsh.ComputerName;
                    webSocket.Send(res.ToString());
                    break;
            }
        }

        Excel.Application xlApp;
        private void ExcelAction(IWebSocketConnection webSocket, JObject action)
        {
            var excelAction = action["command"].Value<string>();
            switch (excelAction)
            {
                case "open":
                    {
                        xlApp = new Excel.Application();
                        xlApp.Visible = true;
                        JObject res = new JObject();
                        res["action"] = "excel";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());
                    }
                    break;
                case "close":
                    {
                        xlApp.Quit();
                        Marshal.FinalReleaseComObject(xlApp);
                        xlApp = null;
                        JObject res = new JObject();
                        res["action"] = "excel";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());
                    }
                    break;
                case "setcell":
                    {
                        var row = action["row"].Value<int>();
                        var column = action["row"].Value<int>();
                        var activeSheet = xlApp.Application.ActiveWorkbook.ActiveSheet;
                        var range = activeSheet.Rows.Cells[row, column];
                        range.Value = action["value"].Value<string>();
                        JObject res = new JObject();
                        res["action"] = "excel";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());
                    }
                    break;
                case "getcell":
                    {
                        var row = action["row"].Value<int>();
                        var column = action["row"].Value<int>();
                        var activeSheet = xlApp.Application.ActiveWorkbook.ActiveSheet;
                        var range = activeSheet.Rows.Cells[row, column];
                        JObject res = new JObject();
                        res["action"] = "excel";
                        res["result"] = range.Value;
                        webSocket.Send(res.ToString());
                    }
                    break;
            }
        }


        Word.Application wordApp;
        private void WordAction(IWebSocketConnection webSocket, JObject action)
        {
            var wordAction = action["command"].Value<string>();
            switch (wordAction)
            {
                case "open":
                    {
                        wordApp = new Word.Application();
                        wordApp.Visible = true;
                        JObject res = new JObject();
                        res["action"] = "word";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());
                    }
                    break;
                case "close":
                    {
                        wordApp.Quit();
                        Marshal.FinalReleaseComObject(wordApp);
                        wordApp = null;
                        JObject res = new JObject();
                        res["action"] = "word";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());
                    }
                    break;
                case "settext":
                    {
                        //Create a missing variable for missing value  
                        object missing = System.Reflection.Missing.Value;
                        //Create a new document  
                        var document = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                        //Add header into the document  
                        foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                        {
                            //Get the header range and add the header details.  
                            Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                            headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                            headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                            headerRange.Font.Size = 10;
                            headerRange.Text = "Header text goes here";
                        }

                        //Add the footers into the document  
                        foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                        {
                            //Get the footer range and add the footer details.  
                            Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                            footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                            footerRange.Font.Size = 10;
                            footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            footerRange.Text = "Footer text goes here";
                        }

                        //adding text to document  
                        document.Content.SetRange(0, 0);
                        document.Content.Text = "This is test document " + action["text"].Value<string>() + Environment.NewLine;

                        JObject res = new JObject();
                        res["action"] = "word";
                        res["result"] = "OK";
                        webSocket.Send(res.ToString());

                    }
                    break;
            }
        }

        /// <summary>
        /// Runs a command line program and returns the output thru the websocket
        /// </summary>
        /// <param name="webSocket"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        private void RunProgram(IWebSocketConnection webSocket, JObject action)
        {
            var command = action["command"].Value<string>();

            var commandLineParameters = action["parameters"].Value<string>();


            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = command,
                    Arguments = commandLineParameters,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            proc.Start();
            string line = "";
            while (!proc.StandardOutput.EndOfStream)
            {
                line += proc.StandardOutput.ReadLine();
                // do something with line
            }

            JObject res = new JObject();
            res["action"] = "execprogram";
            res["result"] = line;
            webSocket.Send(res.ToString());
        }

        /// <summary>
        /// Some COM References can only run on an STA thread
        /// </summary>
        /// <param name="webSocket"></param>
        /// <param name="action"></param>
        private void RunOnSTATread(IWebSocketConnection webSocket, JObject action, Action<IWebSocketConnection, JObject> staThreadAction)
        {
            var thread = new Thread(() => staThreadAction(webSocket, action));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private readonly IList<IWebSocketConnection> _allSockets = new List<IWebSocketConnection>();

        private void SetupServer()
        {
            lblStatus.Text = "0 clients connected";
            var server = new WebSocketServer("wss://0.0.0.0:996");
            server.Certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(@"M:\Program Files\OpenSSL-Win64\bin\secondtest.pfx", "test");
            server.Start(c =>
            {
                c.OnOpen = () =>
                {
                    Invoke(new Action(() => txtLog.AppendText("Open!" + Environment.NewLine)));
                    _allSockets.Add(c);
                    Invoke(new Action(() => lblStatus.Text = $"{ _allSockets.Count} clients connected"));
                };
                c.OnClose = () =>
                {
                    Invoke(new Action(() => txtLog.AppendText("Closed!" + Environment.NewLine)));
                    _allSockets.Remove(c);
                    Invoke(new Action(() => lblStatus.Text = $"{ _allSockets.Count} clients connected"));
                };
                c.OnMessage = m => ExecuteActions(c, m);
            });
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            Restore(this);
        }

        [DllImport("user32.dll")]
        private static extern int ShowWindow(IntPtr hWnd, uint Msg);

        private const uint SW_RESTORE = 0x09;

        public static void Restore(Form form)
        {
            if (form.WindowState == FormWindowState.Minimized)
            {
                ShowWindow(form.Handle, SW_RESTORE);
            }
        }
    }
}
