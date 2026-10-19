using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// Python/Paddle are intentionally not frozen into a PyInstaller import graph.
// A verified core is installed once; the pinned CPU wheels are provisioned by
// its private Python. The public artifact remains one ordinary Windows EXE.
static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        AppContext.SetSwitch("Switch.System.IO.UseLegacyPathHandling", false);
        AppContext.SetSwitch("Switch.System.IO.BlockLongPaths", false);
        var form = new InstallerWindow(args);
        Application.Run(form);
        return form.ExitCode;
    }
}

sealed class InstallerWindow : Form
{
    readonly Label message = new Label();
    readonly ProgressBar progress = new ProgressBar();
    readonly Button cancel = new Button();
    readonly string[] arguments;
    string root, cancelFile, dataDirectory;
    bool offline, prepareOnly, quiet, finished, customInstallDirectory;
    volatile bool cancelled;
    DateTime lastUpdate = DateTime.MinValue;
    public int ExitCode = 1;

    public InstallerWindow(string[] args)
    {
        arguments = args;
        Text = "Документ → Excel · подготовка";
        ClientSize = new Size(570, 190);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        message.SetBounds(24, 20, 520, 65);
        message.Text = "Подготовка приложения. При первом запуске потребуется интернет для PaddleOCR CPU и моделей.";
        progress.SetBounds(24, 100, 520, 22);
        cancel.SetBounds(424, 142, 120, 32);
        cancel.Text = "Отмена";
        cancel.Click += delegate { RequestCancel(); };
        Controls.AddRange(new Control[] { message, progress, cancel });
        FormClosing += delegate(object sender, FormClosingEventArgs e) {
            if (!finished) { RequestCancel(); e.Cancel = true; }
        };
        Shown += async delegate {
            try {
                ParseOptions();
                if (quiet) { Opacity = 0; ShowInTaskbar = false; }
                await Task.Run((Action)InstallAndLaunch);
                ExitCode = cancelled ? 3 : 0;
            } catch (Exception error) {
                ExitCode = cancelled ? 3 : 2;
                try { if (root != null) File.WriteAllText(root + "-launcher-error.log", error.ToString(), Encoding.UTF8); } catch { }
                if (!quiet && !cancelled)
                    MessageBox.Show("Подготовка не завершена. " + error.Message + "\n\nПроверьте интернет и свободное место, затем повторите запуск этого EXE.\nДокументы пользователя не изменяются.", "Документ → Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally { finished = true; Close(); }
        };
    }

    void ParseOptions()
    {
        // Keep SDK paths below legacy Windows MAX_PATH even for long profile
        // names; document/settings directories remain separate.
        root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DXOCR", PackageInfo.Id);
        for (int i = 0; i < arguments.Length && arguments[i] != "--"; i++) {
            switch (arguments[i]) {
                case "--install-dir": root = arguments[++i]; customInstallDirectory = true; break;
                case "--data-dir": dataDirectory = Path.GetFullPath(arguments[++i]); break;
                case "--offline": offline = true; break;
                case "--prepare-only": prepareOnly = true; break;
                case "--quiet": quiet = true; break;
                default: throw new ArgumentException("Неизвестный параметр запуска");
            }
        }
        root = Path.GetFullPath(root);
        if (root == Path.GetPathRoot(root)) throw new ArgumentException("Нужна отдельная папка приложения");
        if (!customInstallDirectory) SelectDefaultDisk();
        Directory.CreateDirectory(Path.GetDirectoryName(root));
    }

    void SelectDefaultDisk()
    {
        if (CoreReady(root)) return;
        string user;
        using (var hash = SHA256.Create()) user = "user-" + Hex(hash.ComputeHash(Encoding.UTF8.GetBytes(Environment.UserDomainName + "\\" + Environment.UserName))).Substring(0, 12);
        var disks = DriveInfo.GetDrives().Where(d => d.DriveType == DriveType.Fixed && d.IsReady).ToArray();
        // Reuse the same installation even after system-disk space changes.
        foreach (var disk in disks) {
            string candidate = Path.Combine(disk.RootDirectory.FullName, "DXOCR", user, PackageInfo.Id);
            if (CoreReady(candidate)) { root = candidate; return; }
        }
        if (new DriveInfo(Path.GetPathRoot(root)).AvailableFreeSpace >= 2500L * 1048576) return;
        foreach (var disk in disks.OrderByDescending(d => d.AvailableFreeSpace)) {
            if (disk.AvailableFreeSpace < 2500L * 1048576) continue;
            string parent = Path.Combine(disk.RootDirectory.FullName, "DXOCR", user);
            try {
                Directory.CreateDirectory(parent);
                string probe = Path.Combine(parent, ".write-" + Guid.NewGuid().ToString("N"));
                using (var file = new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose)) { }
                root = Path.Combine(parent, PackageInfo.Id);
                return;
            } catch (UnauthorizedAccessException) { }
              catch (IOException) { }
        }
    }

    void RequestCancel()
    {
        cancelled = true;
        cancel.Enabled = false;
        message.Text = "Отмена подготовки…";
        if (cancelFile != null) { try { File.WriteAllText(cancelFile, "cancel"); } catch { } }
    }

    void CheckCancel() { if (cancelled) throw new OperationCanceledException(); }

    void Update(int value, string text)
    {
        if ((DateTime.UtcNow - lastUpdate).TotalMilliseconds < 120 && value != 100) return;
        lastUpdate = DateTime.UtcNow;
        BeginInvoke((Action)delegate { progress.Value = Math.Max(0, Math.Min(100, value)); message.Text = text; });
    }

    static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant(); }

    bool CoreReady() { return CoreReady(root); }

    bool CoreReady(string folder)
    {
        string marker = Path.Combine(folder, ".core-ready");
        return File.Exists(marker) && File.ReadAllText(marker) == PackageInfo.Sha256
            && new string[] { "runtime/python.exe", "runtime/pythonw.exe", "app/run.py", "installer/provision.py", "models/rapidocr/manifest.json" }
                .All(p => File.Exists(Path.Combine(folder, p)));
    }

    void InstallCore()
    {
        if (CoreReady()) return;
        string exe = Application.ExecutablePath;
        string stageName = ".setup-" + Guid.NewGuid().ToString("N").Substring(0, 12);
        string stage = Path.Combine(Path.GetDirectoryName(root), stageName);
        if (Directory.Exists(stage)) throw new IOException("Временная папка занята");
        Directory.CreateDirectory(stage);
        try {
            using (var file = File.OpenRead(exe)) {
                file.Seek(-24, SeekOrigin.End);
                byte[] magic = new byte[16]; file.Read(magic, 0, 16);
                byte[] size = new byte[8]; file.Read(size, 0, 8);
                if (Encoding.ASCII.GetString(magic) != "DOCEXCEL-ZIP-V1!") throw new InvalidDataException("EXE-пакет повреждён");
                long length = BitConverter.ToInt64(size, 0);
                long start = file.Length - 24 - length;
                if (length <= 0 || start < 0) throw new InvalidDataException("Неверный размер EXE-пакета");
                using (var slice = new SliceStream(file, start, length)) {
                    Update(1, "Проверка EXE-пакета…");
                    using (var hash = SHA256.Create()) {
                        if (Hex(hash.ComputeHash(slice)) != PackageInfo.Sha256) throw new InvalidDataException("Контрольная сумма EXE не совпала");
                    }
                    slice.Position = 0;
                    using (var zip = new ZipArchive(slice, ZipArchiveMode.Read, true)) {
                        long total = zip.Entries.Sum(e => e.Length), done = 0;
                        long free = new DriveInfo(Path.GetPathRoot(stage)).AvailableFreeSpace;
                        if (free < total + 1000L * 1048576) throw new IOException("Для установки нужно не менее " + Math.Ceiling(total / 1048576.0 + 1000) + " МБ свободного места");
                        byte[] buffer = new byte[262144];
                        foreach (var entry in zip.Entries) {
                            CheckCancel();
                            string target = Path.GetFullPath(Path.Combine(stage, entry.FullName.Replace('/', Path.DirectorySeparatorChar)));
                            if (!target.StartsWith(stage + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Неверный путь в EXE-пакете");
                            if (entry.FullName.EndsWith("/")) { Directory.CreateDirectory(target); continue; }
                            Directory.CreateDirectory(Path.GetDirectoryName(target));
                            using (var source = entry.Open()) using (var output = File.Create(target)) {
                                int count;
                                while ((count = source.Read(buffer, 0, buffer.Length)) > 0) { CheckCancel(); output.Write(buffer, 0, count); done += count; }
                            }
                            Update((int)(35 * done / Math.Max(1, total)), "Установка приложения: " + (100 * done / Math.Max(1, total)) + "%");
                        }
                    }
                }
            }
            File.WriteAllText(Path.Combine(stage, ".core-ready"), PackageInfo.Sha256);
            if (Directory.Exists(root)) Directory.Move(root, root + ".previous-" + DateTime.UtcNow.Ticks);
            Directory.Move(stage, root);
        } finally {
            // Only our freshly generated staging directory can be removed.
            string parent = Path.GetFullPath(Path.GetDirectoryName(root)) + Path.DirectorySeparatorChar;
            if (Directory.Exists(stage) && Path.GetFullPath(stage).StartsWith(parent, StringComparison.OrdinalIgnoreCase)
                && Path.GetFileName(stage) == stageName)
                Directory.Delete(stage, true);
        }
    }

    ProcessStartInfo Python(bool windowed, string args)
    {
        string temporary = Path.Combine(root, "temp");
        Directory.CreateDirectory(temporary);
        var info = new ProcessStartInfo(Path.Combine(root, windowed ? "runtime/pythonw.exe" : "runtime/python.exe"), args);
        info.UseShellExecute = false; info.CreateNoWindow = true; info.WorkingDirectory = root;
        info.EnvironmentVariables.Remove("PYTHONHOME"); info.EnvironmentVariables.Remove("VIRTUAL_ENV");
        info.EnvironmentVariables["PYTHONPATH"] = Path.Combine(root, "app/src");
        info.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
        info.EnvironmentVariables["PYTHONUTF8"] = "1";
        info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
        info.EnvironmentVariables["TEMP"] = temporary;
        info.EnvironmentVariables["TMP"] = temporary;
        foreach (string key in new string[] { "OMP_NUM_THREADS", "OPENBLAS_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS" }) info.EnvironmentVariables[key] = "2";
        info.EnvironmentVariables["HF_HOME"] = Path.Combine(root, "cache/huggingface");
        info.EnvironmentVariables["MODELSCOPE_CACHE"] = Path.Combine(root, "cache/modelscope");
        info.EnvironmentVariables["PADDLE_PDX_CACHE_HOME"] = Path.Combine(root, "cache/paddlex");
        info.EnvironmentVariables["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True";
        info.EnvironmentVariables["HF_HUB_OFFLINE"] = "1";
        info.EnvironmentVariables["PIP_CONFIG_FILE"] = "nul";
        info.EnvironmentVariables["PIP_NO_INPUT"] = "1";
        // The copied Python retains uv's externally-managed marker. An explicit
        // private target is supported by pip without changing a system runtime.
        info.EnvironmentVariables["PIP_TARGET"] = Path.Combine(root, "runtime/Lib/site-packages");
        info.EnvironmentVariables["PIP_UPGRADE"] = "1";
        info.EnvironmentVariables["DOCUMENT_EXCEL_RAPID_MODELS"] = Path.Combine(root, "models/rapidocr");
        info.EnvironmentVariables["DOCUMENT_EXCEL_PADDLE_MODELS"] = Path.Combine(root, "models/paddle");
        if (dataDirectory != null) info.EnvironmentVariables["DOCUMENT_EXCEL_DATA_DIR"] = dataDirectory;
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        info.EnvironmentVariables["PATH"] = Path.Combine(root, "runtime") + ";" + Path.Combine(windows, "System32") + ";" + windows;
        return info;
    }

    void InstallAndLaunch()
    {
        string identity;
        using (var hash = SHA256.Create()) identity = Hex(hash.ComputeHash(Encoding.UTF8.GetBytes(root))).Substring(0, 20);
        using (var gate = new Mutex(false, "Local\\DocumentExcelInstall-" + identity)) {
            bool locked = false;
            try {
                while (!locked) { CheckCancel(); try { locked = gate.WaitOne(250); } catch (AbandonedMutexException) { locked = true; } }
                InstallCore(); CheckCancel();
                cancelFile = Path.Combine(root, ".cancel-" + Guid.NewGuid().ToString("N"));
                string args = Quote(Path.Combine(root, "installer/provision.py")) + " --root " + Quote(root) + " --cancel-file " + Quote(cancelFile) + (offline ? " --offline" : "");
                var info = Python(false, args);
                info.RedirectStandardOutput = true; info.RedirectStandardError = true;
                info.StandardOutputEncoding = Encoding.UTF8; info.StandardErrorEncoding = Encoding.UTF8;
                string failure = "";
                using (var child = new Process()) using (var log = new StreamWriter(Path.Combine(root, "launcher.log"), false, Encoding.UTF8)) {
                    log.AutoFlush = true;
                    object sync = new object();
                    child.StartInfo = info;
                    child.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e) {
                        if (e.Data == null) return;
                        lock (sync) log.WriteLine(e.Data);
                        string[] parts = e.Data.Split(new char[] {'|'}, 3);
                        int percent;
                        if (parts.Length == 3 && parts[0] == "PROGRESS" && Int32.TryParse(parts[1], out percent)) Update(percent, parts[2]);
                        if (e.Data.StartsWith("ERROR|")) failure = e.Data.Substring(6);
                    };
                    child.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) lock (sync) log.WriteLine(e.Data); };
                    child.Start(); child.BeginOutputReadLine(); child.BeginErrorReadLine(); child.WaitForExit();
                    if (child.ExitCode != 0) throw new InvalidOperationException(failure == "" ? "Ошибка подготовки. Подробности: launcher.log в папке приложения." : failure);
                }
                File.Delete(cancelFile); CheckCancel();
                if (!prepareOnly) {
                    int split = Array.IndexOf(arguments, "--");
                    string extra = split < 0 ? "" : " " + String.Join(" ", arguments.Skip(split + 1).Select(Quote));
                    using (var application = Process.Start(Python(true, Quote(Path.Combine(root, "app/run.py")) + extra))) {
                        if (application.WaitForExit(1200) && application.ExitCode != 0) throw new InvalidOperationException("Приложение не смогло запуститься");
                    }
                }
            } finally { if (locked) gate.ReleaseMutex(); }
        }
    }

    static string Quote(string value)
    {
        var output = new StringBuilder("\""); int slash = 0;
        foreach (char c in value) {
            if (c == '\\') { slash++; continue; }
            if (c == '"') output.Append('\\', slash * 2 + 1); else output.Append('\\', slash);
            output.Append(c); slash = 0;
        }
        output.Append('\\', slash * 2); output.Append('"'); return output.ToString();
    }
}

sealed class SliceStream : Stream
{
    readonly Stream source; readonly long start, length; long position;
    public SliceStream(Stream input, long offset, long size) { source = input; start = offset; length = size; }
    public override bool CanRead { get { return true; } }
    public override bool CanSeek { get { return true; } }
    public override bool CanWrite { get { return false; } }
    public override long Length { get { return length; } }
    public override long Position { get { return position; } set { Seek(value, SeekOrigin.Begin); } }
    public override int Read(byte[] buffer, int offset, int count) { source.Position = start + position; int read = source.Read(buffer, offset, (int)Math.Min(count, length - position)); position += read; return read; }
    public override long Seek(long offset, SeekOrigin origin) { long next = (origin == SeekOrigin.Begin ? 0 : origin == SeekOrigin.Current ? position : length) + offset; if (next < 0 || next > length) throw new IOException(); return position = next; }
    public override void Flush() { }
    public override void SetLength(long value) { throw new NotSupportedException(); }
    public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
}
