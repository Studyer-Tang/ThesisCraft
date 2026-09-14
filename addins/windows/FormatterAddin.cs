using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("ThesisCraft")]
[assembly: AssemblyCompany("Study-Tang")]
[assembly: AssemblyProduct("ThesisCraft")]
[assembly: AssemblyVersion("3.0.0.0")]
[assembly: ComVisible(true)]

namespace StudyTang.WordFormatter
{
    // Office's stable COM contracts, declared locally to avoid requiring PIAs.
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI(string ribbonId);
    }

    [ComVisible(true), Guid("5D769C85-5D2E-4301-A55F-459C482B4FC4"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFormatterCommands
    {
        void FormatCurrent(object control);
        void OpenSettings(object control);
        void About(object control);
        string Status();
        string FormatSavedFile(string inputPath);
    }

    [ComVisible(false)]
    public class FormatResult
    {
        public string[] outputs { get; set; }
        public object[] failures { get; set; }
        public object[] skipped { get; set; }
        public int exit_code { get; set; }
    }

    [ComVisible(true), Guid("6B7621A4-AC12-47DE-AF27-F7AEB5A34B7C"), ProgId("StudyTang.WordFormatter"),
        ClassInterface(ClassInterfaceType.AutoDual)]
    public class FormatterAddin : IDTExtensibility2, IRibbonExtensibility, IFormatterCommands
    {
        private object application;
        private bool busy;

        public void OnConnection(object app, int mode, object addin, ref Array custom)
        {
            application = app;
            try { ((dynamic)addin).Object = this; } catch { }
        }
        public void OnDisconnection(int mode, ref Array custom) { application = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { application = null; }
        public string Status() { return "Study-Tang / 3.0.0 / " + (application == null ? "disconnected" : "connected"); }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon><tabs><tab id='StudyTangFormatterTab' label='学研排版'>
                <group id='StudyTangFormatterGroup' label='ThesisCraft'>
                  <button id='FormatCopy' label='排版为新副本' size='large' imageMso='FormatPainter'
                    onAction='FormatCurrent' screentip='排版当前已保存的文档'
                    supertip='请先保存修改。输出新文件，保留原文档。'/>
                  <button id='Settings' label='排版设置 / 桌面版' onAction='OpenSettings' imageMso='FilePageSetup'/>
                  <button id='About' label='关于 Study-Tang' onAction='About' imageMso='Help'/>
                </group></tab></tabs></ribbon></customUI>";
        }

        private static string Setting(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Study-Tang\ThesisCraft"))
            {
                string value = key == null ? null : key.GetValue(name) as string;
                if (String.IsNullOrWhiteSpace(value)) throw new InvalidOperationException("插件尚未安装完整，请重新运行 install.ps1。");
                return value;
            }
        }

        // Windows CommandLineToArgvW quoting, including trailing backslashes.
        private static string Quote(string value)
        {
            StringBuilder result = new StringBuilder("\"");
            int slashes = 0;
            foreach (char c in value)
            {
                if (c == '\\') { slashes++; continue; }
                result.Append('\\', c == '"' ? slashes * 2 + 1 : slashes);
                result.Append(c);
                slashes = 0;
            }
            result.Append('\\', slashes * 2);
            return result.Append('"').ToString();
        }

        private static ProcessStartInfo StartInfo(string arguments)
        {
            return new ProcessStartInfo(Setting("PythonPath"), arguments) {
                WorkingDirectory = Setting("ProjectPath"), UseShellExecute = false,
                CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
            };
        }

        public string FormatSavedFile(string inputPath)
        {
            inputPath = Path.GetFullPath(inputPath);
            if (!File.Exists(inputPath)) throw new FileNotFoundException("请先把文档保存到本地。", inputPath);
            string arguments = "-X utf8 -m word_formatter.office " + Quote(inputPath);
            using (Process process = Process.Start(StartInfo(arguments)))
            {
                Task<string> output = process.StandardOutput.ReadToEndAsync();
                Task<string> error = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(300000))
                {
                    // Keep the child alive to finish its transactional save.
                    throw new TimeoutException("处理超过五分钟。后台仍会完成安全保存，请检查原文件旁的输出；暂勿重复点击。");
                }
                Task.WaitAll(output, error);
                if (process.ExitCode != 0) throw new InvalidOperationException("排版未完成：" + error.Result + output.Result);
                FormatResult result = new JavaScriptSerializer().Deserialize<FormatResult>(output.Result);
                if (result == null || result.outputs == null || result.outputs.Length != 1 || !File.Exists(result.outputs[0]))
                    throw new InvalidOperationException("排版核心没有返回可用的新文档。");
                return result.outputs[0];
            }
        }

        public void FormatCurrent(object control)
        {
            if (busy || application == null) return;
            object document = null;
            try
            {
                dynamic app = application;
                if ((int)app.Documents.Count == 0) throw new InvalidOperationException("请先打开需要排版的文档。");
                document = app.ActiveDocument;
                dynamic active = document;
                if (String.IsNullOrEmpty((string)active.Path) || !(bool)active.Saved)
                    throw new InvalidOperationException("请先按 Ctrl+S 保存当前修改，然后再排版为新副本。");
                string input = (string)active.FullName;
                busy = true;
                using (Form progress = new Form())
                {
                    progress.Text = "Study-Tang · 正在排版";
                    progress.Width = 430; progress.Height = 150;
                    progress.StartPosition = FormStartPosition.CenterScreen;
                    progress.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progress.ControlBox = false;
                    progress.Controls.Add(new Label { Text = "正在生成排版副本，完成后自动打开。\n请等待当前任务结束。", Dock = DockStyle.Fill, Padding = new Padding(20) });
                    Task<string> work = Task.Run(() => FormatSavedFile(input));
                    using (Timer timer = new Timer { Interval = 150 })
                    {
                        timer.Tick += (sender, args) => { if (work.IsCompleted) { timer.Stop(); progress.Close(); } };
                        timer.Start();
                        progress.ShowDialog();
                    }
                    if (work.IsFaulted) throw work.Exception.GetBaseException();
                    if (application != null) ((dynamic)application).Documents.Open(work.Result);
                }
            }
            catch (Exception error) { MessageBox.Show(error.Message, "学研排版", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            finally
            {
                busy = false;
                if (document != null && Marshal.IsComObject(document)) Marshal.ReleaseComObject(document);
            }
        }

        public void OpenSettings(object control)
        {
            try
            {
                ProcessStartInfo info = StartInfo("-X utf8 -m word_formatter");
                info.RedirectStandardOutput = false; info.RedirectStandardError = false;
                Process.Start(info);
            }
            catch (Exception error) { MessageBox.Show(error.Message, "学研排版"); }
        }
        public void About(object control)
        {
            MessageBox.Show("ThesisCraft 4.0.0\n作者：Study-Tang\n\n当前已保存文档 → 本地排版 → 新副本\n使用桌面版“保存为默认”可同步插件配置。", "关于 Study-Tang");
        }
    }
}
