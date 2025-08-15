using Microsoft.Win32; // <-- đọc DisplayVersion/Build từ Registry
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System.Windows.Forms;

//version 009.2 nang cap windows+change kms
namespace IT_Support_Toolkit
{
    public partial class Homepage : Form
    {
        public Homepage()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Bạn có thể thêm code khi form load ở đây
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Lấy tên hệ điều hành thân thiện (có 24H2/22H2 nếu có)
                string osFriendlyName = GetFriendlyOSName();

                // Lấy thông tin cơ bản
                string osVersion = Environment.OSVersion.VersionString; // GIỮ NGUYÊN osVersion như bạn yêu cầu
                string machineName = Environment.MachineName;
                string userName = Environment.UserName;

                // Lấy thông tin CPU
                string cpuName = "";
                using (ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_Processor"))
                {
                    foreach (ManagementObject mo in mos.Get())
                    {
                        cpuName = mo["Name"].ToString();
                    }
                }

                // Lấy dung lượng RAM
                double totalRAM = 0;
                using (ManagementObjectSearcher mos = new ManagementObjectSearcher("Select * from Win32_ComputerSystem"))
                {
                    foreach (ManagementObject mo in mos.Get())
                    {
                        totalRAM = Math.Round(
                            Convert.ToDouble(mo["TotalPhysicalMemory"]) / (1024 * 1024 * 1024),
                            2
                        );
                    }
                }

                // Lấy thông tin chi tiết về ổ đĩa
                string diskInfo = GetDiskInformation();

                // Hiển thị thông tin
                string info = string.Join(Environment.NewLine, new string[]
                {
                    $"Tên máy: {machineName}",
                    $"Người dùng: {userName}",
                    $"Hệ điều hành: {osFriendlyName} ({osVersion})", // Đã bao gồm cả build
					$"CPU: {cpuName}",
                    $"RAM: {totalRAM} GB",
                    "", // Dòng trống để phân cách
					"=== THÔNG TIN Ổ ĐĨA ===",
                    diskInfo
                });

                // Sử dụng form riêng để hiển thị thông tin dài
                //ShowDetailedInfo(info);

                // MessageBox đơn giản có nút Copy tùy chỉnh
                ShowCustomMessageBox(info);

                // MessageBox mặc định của Windows
                //MessageBox.Show(info, "Thông tin máy tính", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        // Hàm lấy thông tin chi tiết về ổ đĩa
        private string GetDiskInformation()
        {
            StringBuilder diskInfo = new StringBuilder();

            try
            {
                // Lấy thông tin Physical Disk (ổ cứng vật lý)
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive"))
                {
                    int diskIndex = 1;
                    foreach (ManagementObject disk in searcher.Get())
                    {
                        string model = disk["Model"]?.ToString() ?? "Không xác định";
                        string interfaceType = disk["InterfaceType"]?.ToString() ?? "Không xác định";
                        string mediaType = disk["MediaType"]?.ToString() ?? "Không xác định";

                        // Chuyển đổi size từ bytes sang GB
                        long sizeBytes = Convert.ToInt64(disk["Size"] ?? "0");
                        double sizeGB = Math.Round(sizeBytes / (1024.0 * 1024.0 * 1024.0), 2);

                        // Xác định loại ổ đĩa (SSD/HDD)
                        string diskType = GetDiskType(disk["Index"]?.ToString());

                        diskInfo.AppendLine($"Ổ đĩa {diskIndex}: {model}");
                        diskInfo.AppendLine($"  • Dung lượng: {sizeGB} GB");
                        diskInfo.AppendLine($"  • Loại: {diskType}");
                        diskInfo.AppendLine($"  • Giao tiếp: {interfaceType}");
                        diskInfo.AppendLine($"  • Media Type: {mediaType}");

                        diskInfo.AppendLine();
                        diskIndex++;
                    }
                }

                // Thêm thông tin tóm tắt về các ổ đĩa logic
                diskInfo.AppendLine("=== TÓM TẮT CÁC Ổ ĐĨA LOGIC ===");
                DriveInfo[] drives = DriveInfo.GetDrives();
                foreach (DriveInfo drive in drives)
                {
                    if (drive.IsReady && drive.DriveType == DriveType.Fixed)
                    {
                        double totalGB = Math.Round(drive.TotalSize / (1024.0 * 1024.0 * 1024.0), 2);
                        double freeGB = Math.Round(drive.AvailableFreeSpace / (1024.0 * 1024.0 * 1024.0), 2);
                        double usedGB = Math.Round(totalGB - freeGB, 2);
                        double percentUsed = Math.Round((usedGB / totalGB) * 100, 1);

                        diskInfo.AppendLine($"Ổ {drive.Name} ({drive.DriveFormat})");
                        diskInfo.AppendLine($"  • Tổng: {totalGB} GB");
                        diskInfo.AppendLine($"  • Đã dùng: {usedGB} GB ({percentUsed}%)");
                        diskInfo.AppendLine($"  • Còn trống: {freeGB} GB");
                        diskInfo.AppendLine();
                    }
                }
            }
            catch (Exception ex)
            {
                diskInfo.AppendLine($"Lỗi khi lấy thông tin ổ đĩa: {ex.Message}");
            }

            return diskInfo.ToString();
        }

        // Hàm xác định loại ổ đĩa (SSD/HDD)
        private string GetDiskType(string diskIndex)
        {
            try
            {
                // Phương pháp 1: Kiểm tra qua Win32_PhysicalMedia
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_PhysicalMedia WHERE Tag='\\\\.\\PHYSICALDRIVE{diskIndex}'"))
                {
                    foreach (ManagementObject media in searcher.Get())
                    {
                        string mediaType = media["MediaType"]?.ToString() ?? "";
                        if (mediaType.ToLower().Contains("ssd") || mediaType.ToLower().Contains("solid"))
                            return "SSD";
                    }
                }

                // Phương pháp 2: Kiểm tra qua MSFT_PhysicalDisk (Windows 8+)
                try
                {
                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\Microsoft\\Windows\\Storage", $"SELECT * FROM MSFT_PhysicalDisk WHERE DeviceId='{diskIndex}'"))
                    {
                        foreach (ManagementObject disk in searcher.Get())
                        {
                            var mediaType = disk["MediaType"];
                            if (mediaType != null)
                            {
                                // MediaType: 3=HDD, 4=SSD, 5=SCM
                                switch (Convert.ToInt32(mediaType))
                                {
                                    case 3: return "HDD (Cơ học)";
                                    case 4: return "SSD (Thể rắn)";
                                    case 5: return "SCM (Storage Class Memory)";
                                    default: return "Không xác định";
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Windows cũ không hỗ trợ MSFT_PhysicalDisk
                }

                // Phương pháp 3: Dự đoán dựa trên tốc độ seek time (không chính xác 100%)
                return "Không xác định (có thể là HDD)";
            }
            catch
            {
                return "Không xác định";
            }
        }

        // Hàm tìm drive letter từ partition
        private string GetDriveLetterFromPartition(string partitionDeviceID)
        {
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_LogicalDiskToPartition WHERE PartName='{partitionDeviceID}'"))
                {
                    foreach (ManagementObject logicalDisk in searcher.Get())
                    {
                        string dependent = logicalDisk["Dependent"]?.ToString();
                        if (!string.IsNullOrEmpty(dependent))
                        {
                            // Extract drive letter from path like "Win32_LogicalDisk.DeviceID=\"C:\""
                            int startIndex = dependent.IndexOf("\"") + 1;
                            int endIndex = dependent.LastIndexOf("\"");
                            if (startIndex > 0 && endIndex > startIndex)
                            {
                                return dependent.Substring(startIndex, endIndex - startIndex);
                            }
                        }
                    }
                }
            }
            catch { }
            return "";
        }

        // Hàm hiển thị thông tin chi tiết trong cửa sổ riêng
        private void ShowDetailedInfo(string info)
        {
            Form detailForm = new Form()
            {
                Text = "Thông tin chi tiết hệ thống",
                Size = new System.Drawing.Size(600, 500),
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = true,
                MinimizeBox = true
            };

            TextBox textBox = new TextBox()
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Text = info,
                //Text = info.Replace("\n", Environment.NewLine), // FIX: Chuyển \n thành Windows newline
                WordWrap = true, // Thêm word wrap
                Font = new System.Drawing.Font("Consolas", 9)
            };

            Button copyButton = new Button()
            {
                Text = "Copy to Clipboard",
                Dock = DockStyle.Bottom,
                Height = 30
            };

            copyButton.Click += (s, e) => {
                Clipboard.SetText(info);
                MessageBox.Show("Đã copy thông tin vào clipboard!");
            };

            detailForm.Controls.Add(textBox);
            detailForm.Controls.Add(copyButton);
            detailForm.ShowDialog();
        }

        // MessageBox tùy chỉnh với nút Copy và Close
        private void ShowCustomMessageBox(string info)
        {
            // Tạo form nhỏ gọn như MessageBox
            Form msgForm = new Form()
            {
                Text = "Thông tin máy tính",
                Size = new System.Drawing.Size(500, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = false
            };

            // TextBox để hiển thị thông tin
            TextBox textBox = new TextBox()
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Text = info,
                Font = new System.Drawing.Font("Segoe UI", 9),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(msgForm.Width - 40, msgForm.Height - 80)
            };

            // Nút Copy
            Button copyButton = new Button()
            {
                Text = "Copy",
                Size = new System.Drawing.Size(80, 30),
                Location = new System.Drawing.Point(msgForm.Width - 180, msgForm.Height - 60),
                UseVisualStyleBackColor = true
            };

            // Nút Close
            Button closeButton = new Button()
            {
                Text = "Đóng",
                Size = new System.Drawing.Size(80, 30),
                Location = new System.Drawing.Point(msgForm.Width - 90, msgForm.Height - 60),
                UseVisualStyleBackColor = true,
                DialogResult = DialogResult.OK
            };

            // Xử lý sự kiện Copy
            copyButton.Click += (s, e) => {
                Clipboard.SetText(info);
                copyButton.Text = "✓ Copied";
                copyButton.BackColor = System.Drawing.Color.LightGreen;

                // Reset lại sau 2 giây
                Timer timer = new Timer();
                timer.Interval = 2000;
                timer.Tick += (sender, args) => {
                    copyButton.Text = "Copy";
                    copyButton.BackColor = System.Drawing.SystemColors.Control;
                    timer.Stop();
                };
                timer.Start();
            };

            // Xử lý sự kiện Close
            closeButton.Click += (s, e) => msgForm.Close();

            // Thêm controls vào form
            msgForm.Controls.Add(textBox);
            msgForm.Controls.Add(copyButton);
            msgForm.Controls.Add(closeButton);

            // Hiển thị form
            msgForm.ShowDialog();
        }
        private string GetFriendlyOSName()
        {
            // Lấy Caption từ WMI (ví dụ: "Microsoft Windows 11 Pro")
            string caption = "Không xác định";
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject os in searcher.Get())
                    {
                        caption = os["Caption"]?.ToString().Trim() ?? caption;
                    }
                }
            }
            catch { /* bỏ qua, sẽ dùng caption mặc định nếu WMI lỗi */ }

            // Đọc DisplayVersion và Build từ Registry để thêm 24H2/22H2…
            string displayVersion = null;
            string build = null;

            try
            {
                using (var cv = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    displayVersion = cv?.GetValue("DisplayVersion") as string; // ví dụ "24H2", "22H2"
                    build = cv?.GetValue("CurrentBuild") as string;            // ví dụ "26100", "22631", "19045"
                }
            }
            catch { /* người dùng không có quyền ghi, nhưng đọc HKLM thường được phép */ }

            // Nếu không có DisplayVersion (Windows 10/11 cũ) thì suy ra từ build
            if (string.IsNullOrWhiteSpace(displayVersion))
            {
                displayVersion = MapBuildToRelease(build);
            }

            // Ghép chuỗi: "Windows 11 Pro 24H2 (Build 26100)"
            string tail = "";
            if (!string.IsNullOrWhiteSpace(displayVersion))
                tail = " " + displayVersion;

            if (!string.IsNullOrWhiteSpace(build))
                tail += $" (Build {build})";

            return (caption + tail).Trim();
        }

        // Fallback khi không có DisplayVersion trong Registry
        private string MapBuildToRelease(string buildNumber)
        {
            if (!int.TryParse(buildNumber, out int build)) return "";

            // Windows 11
            if (build >= 26100) return "24H2";
            if (build >= 22631) return "23H2";
            if (build >= 22621) return "22H2";
            if (build >= 22000) return "21H2";

            // Windows 10 (một vài mốc chính)
            if (build >= 19045) return "22H2";
            if (build >= 19044) return "21H2";
            if (build >= 19043) return "21H1";
            if (build >= 19042) return "20H2";
            if (build >= 19041) return "2004";
            if (build >= 18363) return "1909";
            if (build >= 18362) return "1903";
            if (build >= 17763) return "1809";

            return "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Code cho button2
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // Code cho button14
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                // Các thư mục cần dọn cho user hiện tại
                string[] paths = {
            Environment.GetEnvironmentVariable("TEMP"),
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Temp",
            //Environment.GetFolderPath(Environment.SpecialFolder.Recent)
        };

                int deletedFiles = 0;
                long freedBytes = 0;

                foreach (string path in paths)
                {
                    if (System.IO.Directory.Exists(path))
                    {
                        // Xóa file
                        foreach (string file in System.IO.Directory.GetFiles(path))
                        {
                            try
                            {
                                long size = new System.IO.FileInfo(file).Length;
                                System.IO.File.Delete(file);
                                deletedFiles++;
                                freedBytes += size;
                            }
                            catch { /* Bỏ qua nếu file đang được sử dụng */ }
                        }

                        // Xóa thư mục con
                        foreach (string dir in System.IO.Directory.GetDirectories(path))
                        {
                            try
                            {
                                long dirSize = GetDirectorySize(dir);
                                System.IO.Directory.Delete(dir, true);
                                deletedFiles++;
                                freedBytes += dirSize;
                            }
                            catch { /* Bỏ qua nếu thư mục đang được sử dụng */ }
                        }
                    }
                }

                double freedMB = Math.Round(freedBytes / (1024.0 * 1024.0), 2);
                MessageBox.Show(
                    $"Đã xóa {deletedFiles} mục, giải phóng {freedMB} MB.",
                    "Dọn file rác",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi dọn file rác: " + ex.Message);
            }
        }

        // Hàm tính dung lượng thư mục
        private long GetDirectorySize(string folderPath)
        {
            long size = 0;
            try
            {
                foreach (string file in System.IO.Directory.GetFiles(folderPath, "*", System.IO.SearchOption.AllDirectories))
                {
                    try
                    {
                        size += new System.IO.FileInfo(file).Length;
                    }
                    catch { }
                }
            }
            catch { }
            return size;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string exportDir = Path.Combine(appDir, "Wifi_Export");
                Directory.CreateDirectory(exportDir);

                // Lấy danh sách Wi-Fi
                var psiList = new ProcessStartInfo("netsh", "wlan show profiles")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };
                string output;
                using (var p = Process.Start(psiList))
                {
                    output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                }

                // Ghi log danh sách Wi-Fi ra file UTF-8
                File.WriteAllText(Path.Combine(exportDir, "Wifi_List.txt"), output, Encoding.UTF8);

                // Lấy tên profile từ output
                var lines = output.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.Contains(":"))
                    {
                        string name = line.Split(':')[1].Trim();
                        if (!string.IsNullOrEmpty(name))
                        {
                            // Export từng profile kèm key=clear
                            var psiExport = new ProcessStartInfo("netsh", $"wlan export profile name=\"{name}\" key=clear folder=\"{exportDir}\"")
                            {
                                UseShellExecute = false,
                                CreateNoWindow = true
                            };
                            using (var p = Process.Start(psiExport))
                            {
                                p.WaitForExit();
                            }
                        }
                    }
                }

                // Nén thư mục export thành zip
                string zipPath = Path.Combine(appDir, "Wifi_Backup.zip");
                if (File.Exists(zipPath))
                    File.Delete(zipPath);
                ZipFile.CreateFromDirectory(exportDir, zipPath);

                MessageBox.Show($"Đã xuất và nén Wi-Fi profiles vào:\n{zipPath}", "Hoàn tất");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        // Hàm kiểm tra quyền admin
        private bool IsRunAsAdmin()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        // Hàm khởi động lại app với quyền admin
        private void RelaunchAsAdmin()
        {
            ProcessStartInfo proc = new ProcessStartInfo
            {
                UseShellExecute = true,
                WorkingDirectory = Environment.CurrentDirectory,
                FileName = Application.ExecutablePath,
                Verb = "runas"
            };
            try
            {
                Process.Start(proc);
            }
            catch
            {
                MessageBox.Show("Bạn đã từ chối quyền Administrator.");
            }
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        //Wifi import button
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                string currentDir = AppDomain.CurrentDomain.BaseDirectory;
                string zipPath = Path.Combine(currentDir, "Wifi_Backup.zip");
                string tempFolder = Path.Combine(currentDir, "Wifi_Import");

                if (!File.Exists(zipPath))
                {
                    MessageBox.Show("Không tìm thấy file Wifi_Backup.zip", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);

                ZipFile.ExtractToDirectory(zipPath, tempFolder);

                // Import từng file XML
                foreach (string xmlFile in Directory.GetFiles(tempFolder, "*.xml"))
                {
                    ProcessStartInfo psi = new ProcessStartInfo("netsh", $"wlan add profile filename=\"{xmlFile}\" user=all")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false
                    };
                    Process.Start(psi)?.WaitForExit();
                }

                MessageBox.Show("✅ Đã import xong tất cả Wi-Fi profiles.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // Kiểm tra quyền admin
            if (!IsRunAsAdmin())
            {
                MessageBox.Show("Ứng dụng cần chạy với quyền Administrator để sao lưu driver.",
                    "Cần quyền Admin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RelaunchAsAdmin();
                return;
            }

            try
            {
                // Tạo folder backup
                string backupPath = CreateDriverBackupFolder();

                // Hiển thị progress form
                ShowDriverBackupProgress(backupPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi sao lưu driver: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Tạo thư mục backup driver
        private string CreateDriverBackupFolder()
        {
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            string backupDir = Path.Combine(currentDir, "Driver_Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            if (!Directory.Exists(backupDir))
            {
                Directory.CreateDirectory(backupDir);
            }

            return backupDir;
        }

        // Hiển thị progress form trong khi backup
        private void ShowDriverBackupProgress(string backupPath)
        {
            Form progressForm = new Form()
            {
                Text = "Đang sao lưu driver...",
                Size = new System.Drawing.Size(500, 250),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = false
            };

            Label statusLabel = new Label()
            {
                Text = "Đang quét danh sách driver...",
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(460, 20),
                Font = new System.Drawing.Font("Segoe UI", 9)
            };

            ProgressBar progressBar = new ProgressBar()
            {
                Location = new System.Drawing.Point(10, 50),
                Size = new System.Drawing.Size(460, 25),
                Style = ProgressBarStyle.Continuous
            };

            TextBox logTextBox = new TextBox()
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Location = new System.Drawing.Point(10, 85),
                Size = new System.Drawing.Size(460, 100),
                Font = new System.Drawing.Font("Consolas", 8)
            };

            Button cancelButton = new Button()
            {
                Text = "Hủy",
                Size = new System.Drawing.Size(80, 30),
                Location = new System.Drawing.Point(400, 195),
                DialogResult = DialogResult.Cancel
            };

            progressForm.Controls.AddRange(new Control[] { statusLabel, progressBar, logTextBox, cancelButton });

            // Background worker để backup driver
            System.ComponentModel.BackgroundWorker worker = new System.ComponentModel.BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;

            worker.DoWork += (s, e) => {
                BackupDriversWorker(worker, e, backupPath);
            };

            worker.ProgressChanged += (s, e) => {
                progressBar.Value = e.ProgressPercentage;
                if (e.UserState != null)
                {
                    string message = e.UserState.ToString();
                    statusLabel.Text = message;
                    logTextBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + " - " + message + Environment.NewLine);
                    logTextBox.ScrollToCaret();
                }
            };

            worker.RunWorkerCompleted += (s, e) => {
                if (e.Error != null)
                {
                    MessageBox.Show($"Lỗi: {e.Error.Message}", "Lỗi sao lưu",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (e.Cancelled)
                {
                    statusLabel.Text = "Đã hủy sao lưu";
                    MessageBox.Show("Đã hủy quá trình sao lưu driver.", "Đã hủy",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    var result = (BackupResult)e.Result;
                    statusLabel.Text = "Hoàn thành!";
                    progressBar.Value = 100;

                    string completionMessage = $"Sao lưu driver hoàn thành!\n\n" +
                        $"• Tổng số driver: {result.TotalDrivers}\n" +
                        $"• Đã sao lưu: {result.BackedUpDrivers}\n" +
                        $"• Bỏ qua: {result.SkippedDrivers}\n" +
                        $"• Thư mục: {backupPath}\n\n" +
                        $"Bạn có muốn mở thư mục backup không?";

                    DialogResult openFolder = MessageBox.Show(completionMessage, "Hoàn thành",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                    if (openFolder == DialogResult.Yes)
                    {
                        Process.Start("explorer.exe", backupPath);
                    }
                }

                cancelButton.Text = "Đóng";
                cancelButton.DialogResult = DialogResult.OK;
            };

            cancelButton.Click += (s, e) => {
                if (worker.IsBusy && !worker.CancellationPending)
                {
                    worker.CancelAsync();
                    cancelButton.Text = "Đang hủy...";
                    cancelButton.Enabled = false;
                }
                else
                {
                    progressForm.Close();
                }
            };

            progressForm.Shown += (s, e) => worker.RunWorkerAsync();
            progressForm.ShowDialog();
        }

        // Class để lưu kết quả backup
        private class BackupResult
        {
            public int TotalDrivers { get; set; }
            public int BackedUpDrivers { get; set; }
            public int SkippedDrivers { get; set; }
        }


        // Worker function tối ưu - sử dụng PnPUtil để backup trực tiếp
        private void BackupDriversWorker(System.ComponentModel.BackgroundWorker worker,
            System.ComponentModel.DoWorkEventArgs e, string backupPath)
        {
            BackupResult result = new BackupResult();
            List<DriverInfo> drivers = new List<DriverInfo>();

            try
            {
                // Bước 1: Lấy danh sách driver từ DISM
                worker.ReportProgress(5, "Đang lấy danh sách driver từ hệ thống...");
                drivers = GetDriverListFromDISM();
                result.TotalDrivers = drivers.Count;

                worker.ReportProgress(10, $"Tìm thấy {drivers.Count} driver. Bắt đầu sao lưu...");

                // Bước 2: Backup từng driver trực tiếp vào thư mục tên thân thiện
                int currentDriver = 0;
                foreach (var driver in drivers)
                {
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    currentDriver++;
                    int progressPercent = 10 + (currentDriver * 85 / drivers.Count);

                    // Tạo tên thư mục thân thiện
                    string friendlyName = GetFriendlyDriverName(driver);
                    string driverFolder = Path.Combine(backupPath, SanitizeFileName(friendlyName));
                    driverFolder = GetUniqueDirectoryName(driverFolder);

                    worker.ReportProgress(progressPercent,
                        $"Đang backup ({currentDriver}/{drivers.Count}): {friendlyName}");

                    try
                    {
                        // Tạo thư mục đích
                        Directory.CreateDirectory(driverFolder);

                        // Sử dụng PnPUtil để export driver trực tiếp
                        bool success = BackupDriverWithPnPUtil(driver, driverFolder);

                        if (success)
                        {
                            result.BackedUpDrivers++;

                            // Tạo file thông tin chi tiết
                            CreateDriverInfoFile(driverFolder, driver);

                            worker.ReportProgress(progressPercent,
                                $"✅ Đã backup: {friendlyName}");
                        }
                        else
                        {
                            result.SkippedDrivers++;
                            worker.ReportProgress(progressPercent,
                                $"⚠️ Bỏ qua: {friendlyName} (Inbox driver hoặc không thể copy)");

                            // Xóa thư mục rỗng
                            try { Directory.Delete(driverFolder, false); } catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        result.SkippedDrivers++;
                        worker.ReportProgress(progressPercent,
                            $"❌ Lỗi: {friendlyName} - {ex.Message}");

                        // Ghi log lỗi
                        try
                        {
                            string errorLog = Path.Combine(backupPath, $"error_{SanitizeFileName(driver.PublishedName)}.txt");
                            File.WriteAllText(errorLog,
                                $"Failed to backup driver: {ex.Message}\n" +
                                $"Driver: {driver.PublishedName}\n" +
                                $"Friendly Name: {friendlyName}\n" +
                                $"Time: {DateTime.Now}");
                        }
                        catch { }
                    }
                }

                worker.ReportProgress(95, "Đang tạo file tổng hợp...");
                CreateSummaryFile(backupPath, result, drivers);

                worker.ReportProgress(100, "Hoàn thành sao lưu driver!");
                e.Result = result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi trong quá trình backup: {ex.Message}");
            }
        }

        // Hàm backup driver bằng PnPUtil và file copy - FIXED
        private bool BackupDriverWithPnPUtil(DriverInfo driver, string destinationFolder)
        {
            try
            {
                string foundDriverPath = null;

                // Phương pháp 1: Tìm trong Driver Store theo Published Name
                string driverStore = @"C:\Windows\System32\DriverStore\FileRepository";

                if (Directory.Exists(driverStore))
                {
                    // Tìm thư mục chứa driver
                    var driverFolders = Directory.GetDirectories(driverStore, "*", SearchOption.TopDirectoryOnly)
                        .Where(folder =>
                        {
                            // Tìm theo pattern thông thường
                            return Directory.GetFiles(folder, driver.PublishedName, SearchOption.TopDirectoryOnly).Any() ||
                                   Directory.GetFiles(folder, "*.inf", SearchOption.TopDirectoryOnly)
                                       .Any(inf => Path.GetFileNameWithoutExtension(inf).Equals(
                                           Path.GetFileNameWithoutExtension(driver.OriginalFileName ?? ""),
                                           StringComparison.OrdinalIgnoreCase));
                        })
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(driverFolders))
                    {
                        foundDriverPath = driverFolders;
                    }
                }

                // Phương pháp 2: Copy driver files từ Driver Store
                if (!string.IsNullOrEmpty(foundDriverPath) && Directory.Exists(foundDriverPath))
                {
                    CopyDriverFiles(foundDriverPath, destinationFolder);
                    return true;
                }

                // Phương pháp 3: Fallback - tìm trong Windows\INF
                string windowsInfPath = @"C:\Windows\INF";
                string infPath = Path.Combine(windowsInfPath, driver.PublishedName);

                if (File.Exists(infPath))
                {
                    // Copy INF file
                    string destInf = Path.Combine(destinationFolder, driver.PublishedName);
                    File.Copy(infPath, destInf, true);

                    // Tìm và copy PNF file (nếu có)
                    string pnfPath = Path.Combine(windowsInfPath,
                        Path.GetFileNameWithoutExtension(driver.PublishedName) + ".pnf");
                    if (File.Exists(pnfPath))
                    {
                        string destPnf = Path.Combine(destinationFolder, Path.GetFileName(pnfPath));
                        File.Copy(pnfPath, destPnf, true);
                    }

                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // Copy toàn bộ driver files từ source sang destination - FIXED for .NET Framework
        private void CopyDriverFiles(string sourceFolder, string destinationFolder)
        {
            // Copy tất cả files trong thư mục driver
            foreach (string file in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
            {
                try
                {
                    // Tạo relative path manually cho .NET Framework
                    string relativePath = GetRelativePathManual(sourceFolder, file);
                    string destFile = Path.Combine(destinationFolder, relativePath);

                    // Tạo thư mục đích nếu cần
                    string destDir = Path.GetDirectoryName(destFile);
                    if (!Directory.Exists(destDir))
                    {
                        Directory.CreateDirectory(destDir);
                    }

                    File.Copy(file, destFile, true);
                }
                catch
                {
                    // Bỏ qua file không copy được (có thể đang được sử dụng)
                }
            }
        }

        // Hàm tạo relative path cho .NET Framework
        private string GetRelativePathManual(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) throw new ArgumentNullException("fromPath");
            if (string.IsNullOrEmpty(toPath)) throw new ArgumentNullException("toPath");

            fromPath = Path.GetFullPath(fromPath);
            toPath = Path.GetFullPath(toPath);

            // Nếu file nằm trực tiếp trong thư mục gốc
            if (toPath.StartsWith(fromPath, StringComparison.OrdinalIgnoreCase))
            {
                string relative = toPath.Substring(fromPath.Length);
                if (relative.StartsWith(Path.DirectorySeparatorChar.ToString()) ||
                    relative.StartsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    relative = relative.Substring(1);
                }
                return relative;
            }

            // Fallback: chỉ lấy tên file
            return Path.GetFileName(toPath);
        }
        // Lấy danh sách driver từ DISM - tách riêng để tối ưu
        private List<DriverInfo> GetDriverListFromDISM()
        {
            ProcessStartInfo psi = new ProcessStartInfo("dism", "/online /get-drivers")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(errorOutput))
                {
                    throw new Exception($"DISM Error: {errorOutput}");
                }

                return ParseDriverList(output);
            }
        }

        // Parse danh sách driver từ DISM output - sửa lại logic
        private List<DriverInfo> ParseDriverList(string dismOutput)
        {
            List<DriverInfo> drivers = new List<DriverInfo>();
            string[] lines = dismOutput.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            DriverInfo currentDriver = null;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();

                // Tìm dòng bắt đầu driver info
                if (trimmedLine.StartsWith("Published name", StringComparison.OrdinalIgnoreCase))
                {
                    if (currentDriver != null)
                    {
                        drivers.Add(currentDriver);
                    }
                    currentDriver = new DriverInfo();

                    // Extract published name
                    int colonIndex = trimmedLine.IndexOf(':');
                    if (colonIndex > 0 && colonIndex < trimmedLine.Length - 1)
                    {
                        currentDriver.PublishedName = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                }
                else if (currentDriver != null)
                {
                    if (trimmedLine.StartsWith("Original file name", StringComparison.OrdinalIgnoreCase))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.OriginalFileName = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                    else if (trimmedLine.StartsWith("Inbox", StringComparison.OrdinalIgnoreCase))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.InboxDriver = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                    else if (trimmedLine.StartsWith("Class Name", StringComparison.OrdinalIgnoreCase))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.ClassName = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                    else if (trimmedLine.StartsWith("Provider Name", StringComparison.OrdinalIgnoreCase))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.ProviderName = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                    else if (trimmedLine.StartsWith("Date", StringComparison.OrdinalIgnoreCase) && !trimmedLine.Contains("Updated"))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.Date = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                    else if (trimmedLine.StartsWith("Version", StringComparison.OrdinalIgnoreCase))
                    {
                        int colonIndex = trimmedLine.IndexOf(':');
                        if (colonIndex > 0) currentDriver.Version = trimmedLine.Substring(colonIndex + 1).Trim();
                    }
                }
            }

            // Thêm driver cuối cùng
            if (currentDriver != null && !string.IsNullOrEmpty(currentDriver.PublishedName))
            {
                drivers.Add(currentDriver);
            }

            return drivers;
        }

        // Cải tiến hàm tạo file thông tin driver
        private void CreateDriverInfoFile(string driverFolder, DriverInfo driver)
        {
            string infoFile = Path.Combine(driverFolder, "driver_info.txt");
            string friendlyName = GetFriendlyDriverName(driver);

            string[] infoLines = {
                $"DRIVER INFORMATION",
                new string('=', 60),
                $"Backup Date: {DateTime.Now}",
                $"Friendly Name: {friendlyName}",
                "",
                "TECHNICAL DETAILS:",
                new string('-', 30),
                $"Published Name: {driver.PublishedName ?? "N/A"}",
                $"Original File Name: {driver.OriginalFileName ?? "N/A"}",
                $"Class Name: {driver.ClassName ?? "N/A"}",
                $"Provider Name: {driver.ProviderName ?? "N/A"}",
                $"Version: {driver.Version ?? "N/A"}",
                $"Date: {driver.Date ?? "N/A"}",
                $"Inbox Driver: {driver.InboxDriver ?? "N/A"}",
                "",
                "RESTORE INSTRUCTIONS:",
                new string('-', 30),
                "Method 1 - Using PnPUtil (Recommended):",
                $"  pnputil /add-driver *.inf /install",
                "",
                "Method 2 - Using Device Manager:",
                "  1. Open Device Manager",
                "  2. Right-click on the device",
                "  3. Select 'Update driver'",
                "  4. Choose 'Browse my computer for drivers'",
                "  5. Point to this folder",
                "",
                $"Backup Location: {driverFolder}",
                $"Computer: {Environment.MachineName}",
                $"User: {Environment.UserName}"
            };

            File.WriteAllLines(infoFile, infoLines, Encoding.UTF8);
        }

        // Tạo file tổng hợp
        private void CreateSummaryFile(string backupPath, BackupResult result, List<DriverInfo> drivers)
        {
            string summaryFile = Path.Combine(backupPath, "BACKUP_SUMMARY.txt");

            List<string> summaryLines = new List<string>
    {
        $"DRIVER BACKUP SUMMARY - {DateTime.Now}",
        new string('=', 60),
        $"Computer: {Environment.MachineName}",
        $"User: {Environment.UserName}",
        $"OS: {Environment.OSVersion}",
        "",
        $"Total Drivers Found: {result.TotalDrivers}",
        $"Successfully Backed Up: {result.BackedUpDrivers}",
        $"Skipped: {result.SkippedDrivers}",
        "",
        "DRIVER LIST:",
        new string('-', 30)
    };

            foreach (var driver in drivers)
            {
                summaryLines.Add($"• {driver.PublishedName} - {driver.ProviderName} ({driver.Version})");
            }

            summaryLines.AddRange(new[] {
        "",
        "RESTORE INSTRUCTIONS:",
        new string('-', 30),
        "1. Copy driver folders to target computer",
        "2. Open Command Prompt as Administrator",
        "3. Navigate to driver folder",
        "4. Run: pnputil /add-driver *.inf /install",
        "",
        "Or use Device Manager > Update Driver > Browse for drivers"
    });

            File.WriteAllLines(summaryFile, summaryLines, Encoding.UTF8);
        }

        // Class thông tin driver
        private class DriverInfo
        {
            public string PublishedName { get; set; }
            public string OriginalFileName { get; set; }
            public string InboxDriver { get; set; }
            public string ClassName { get; set; }
            public string ProviderName { get; set; }
            public string Date { get; set; }
            public string Version { get; set; }
        }

        // Các hàm khác giữ nguyên từ code trước
        private string GetUniqueDirectoryName(string basePath)
        {
            string originalPath = basePath;
            int counter = 1;

            while (Directory.Exists(basePath))
            {
                basePath = $"{originalPath}_{counter}";
                counter++;
            }

            return basePath;
        }

        // Hàm tạo tên thân thiện cho driver
        private string GetFriendlyDriverName(DriverInfo driver)
        {
            bool isInboxDriver = !string.IsNullOrEmpty(driver.InboxDriver) &&
                                driver.InboxDriver.Equals("Yes", StringComparison.OrdinalIgnoreCase);

            // ✨ MỚI: Ưu tiên 1 - Tìm tên thiết bị thực tế từ Device Manager qua WMI
            string deviceFriendlyName = GetDeviceFriendlyNameFromWMI(driver);
            if (!string.IsNullOrEmpty(deviceFriendlyName))
            {
                // Thêm thông tin provider nếu cần
                if (!string.IsNullOrEmpty(driver.ProviderName) &&
                    !driver.ProviderName.Equals("Microsoft", StringComparison.OrdinalIgnoreCase) &&
                    !deviceFriendlyName.ToLower().Contains(driver.ProviderName.ToLower().Split(' ')[0]))
                {
                    deviceFriendlyName += $" - {driver.ProviderName}";
                }

                // Thêm version nếu có
                if (!string.IsNullOrEmpty(driver.Version) && !driver.Version.Equals("0.0.0.0"))
                {
                    deviceFriendlyName += $" (v{driver.Version})";
                }

                if (isInboxDriver)
                {
                    deviceFriendlyName = $"[INBOX] {deviceFriendlyName}";
                }

                return deviceFriendlyName; // ✨ RETURN TÊN TỪ DEVICE MANAGER
            }

            // ✨ MỚI: Ưu tiên 2 - Tìm từ Hardware ID patterns trong INF file
            string hardwareBasedName = GetFriendlyNameFromHardwarePattern(driver);
            if (!string.IsNullOrEmpty(hardwareBasedName))
            {
                if (!string.IsNullOrEmpty(driver.Version) && !driver.Version.Equals("0.0.0.0"))
                {
                    hardwareBasedName += $" (v{driver.Version})";
                }

                if (isInboxDriver)
                {
                    hardwareBasedName = $"[INBOX] {hardwareBasedName}";
                }

                return hardwareBasedName; // ✨ RETURN TÊN TỪ INF FILE
            }

            // ⚠️ FALLBACK: Logic cũ (không thay đổi)
            string friendlyName = "";

            if (!string.IsNullOrEmpty(driver.ClassName) &&
                !driver.ClassName.Equals("Unknown", StringComparison.OrdinalIgnoreCase) &&
                !driver.ClassName.Equals("System", StringComparison.OrdinalIgnoreCase))
            {
                friendlyName = driver.ClassName;
                // ... logic cũ không đổi
            }

            return friendlyName;
        }

        // ✨ HÀM MỚI 1: Lấy tên từ WMI (Device Manager)
        // =================================================================
        private string GetDeviceFriendlyNameFromWMI(DriverInfo driver)
        {
            try
            {
                // 🔍 Truy vấn Win32_PnPEntity (thiết bị Plug and Play)
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity"))
                {
                    foreach (ManagementObject device in searcher.Get())
                    {
                        try
                        {
                            string deviceName = device["Name"]?.ToString();        // ✨ TÊN HIỂN THỊ TRONG DEVICE MANAGER
                            string deviceDriver = device["DriverName"]?.ToString();
                            string deviceINF = device["InfName"]?.ToString();      // ✨ TÊN FILE INF

                            // 🎯 So khớp driver với thiết bị
                            if ((!string.IsNullOrEmpty(deviceINF) && deviceINF.Equals(driver.PublishedName, StringComparison.OrdinalIgnoreCase)) ||
                                (!string.IsNullOrEmpty(deviceDriver) && deviceDriver.Equals(driver.OriginalFileName, StringComparison.OrdinalIgnoreCase)))
                            {
                                if (!string.IsNullOrEmpty(deviceName) &&
                                    !deviceName.Contains("Unknown") &&
                                    !deviceName.Contains("Generic"))
                                {
                                    return deviceName; // ✨ RETURN TÊN CHÍNH XÁC TỪ DEVICE MANAGER
                                }
                            }
                        }
                        catch { continue; }
                    }
                }

                // 🔍 Fallback: Tìm từ Win32_SystemDriver
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SystemDriver"))
                {
                    foreach (ManagementObject sysDriver in searcher.Get())
                    {
                        try
                        {
                            string sysDriverDisplayName = sysDriver["DisplayName"]?.ToString(); // ✨ TÊN HIỂN THỊ
                            string sysDriverPathName = sysDriver["PathName"]?.ToString();

                            if (!string.IsNullOrEmpty(sysDriverPathName) &&
                                (sysDriverPathName.Contains(Path.GetFileNameWithoutExtension(driver.OriginalFileName ?? "")) ||
                                 sysDriverPathName.Contains(driver.PublishedName)))
                            {
                                return sysDriverDisplayName;
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch { }

            return null; // ✨ KHÔNG TÌM THẤY
        }

        // =================================================================
        // ✨ HÀM MỚI 2: Lấy tên từ INF file parsing
        // =================================================================
        private string GetFriendlyNameFromHardwarePattern(DriverInfo driver)
        {
            try
            {
                // 📄 Đọc file INF để tìm device descriptions
                string infPath = Path.Combine(@"C:\Windows\INF", driver.PublishedName);
                if (File.Exists(infPath))
                {
                    string[] infLines = File.ReadAllLines(infPath);
                    string deviceDesc = null;

                    // 🔍 Tìm section [Strings] và các device descriptions
                    bool inStringsSection = false;
                    foreach (string line in infLines)
                    {
                        string trimmedLine = line.Trim();

                        if (trimmedLine.StartsWith("[Strings]", StringComparison.OrdinalIgnoreCase))
                        {
                            inStringsSection = true; // ✨ VÀO SECTION STRINGS
                            continue;
                        }

                        if (inStringsSection)
                        {
                            if (trimmedLine.StartsWith("[") && !trimmedLine.StartsWith("[Strings"))
                            {
                                break; // ✨ KẾT THÚC SECTION STRINGS
                            }

                            // 🎯 Tìm device description
                            if (trimmedLine.Contains("=") && !trimmedLine.StartsWith(";"))
                            {
                                string[] parts = trimmedLine.Split('=');
                                if (parts.Length == 2)
                                {
                                    string value = parts[1].Trim().Trim('"'); // ✨ LẤY GIÁ TRỊ STRING

                                    // ✅ Kiểm tra xem có phải là device name không
                                    if (IsDeviceDescription(value))
                                    {
                                        if (string.IsNullOrEmpty(deviceDesc) || value.Length > deviceDesc.Length)
                                        {
                                            deviceDesc = value; // ✨ LƯU TÊN DÀI NHẤT
                                        }
                                    }
                                }
                            }
                        }
                    }

                    return deviceDesc;
                }
            }
            catch { }

            return null;
        }

        // =================================================================
        // ✨ HÀM MỚI 3: Kiểm tra device description hợp lệ
        // =================================================================
        private bool IsDeviceDescription(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length < 5)
                return false;

            // ❌ Loại bỏ các chuỗi không phải device name
            string lowerText = text.ToLower();
            string[] excludePatterns = {
                "microsoft", "corp", "inc", "ltd", "co.", "company",
                "driver", "inf", "sys", "dll", "exe",
                "version", "copyright", "(c)", "©",
                "description", "manufacturer"
            };

            foreach (string pattern in excludePatterns)
            {
                if (lowerText == pattern || (lowerText.Contains(pattern) && lowerText.Replace(pattern, "").Trim().Length < 3))
                {
                    return false; // ❌ KHÔNG PHẢI DEVICE NAME
                }
            }

            // ✅ Device name thường chứa các keywords này
            string[] deviceKeywords = {
                "graphics", "audio", "sound", "network", "ethernet", "wifi", "wireless",
                "bluetooth", "usb", "pci", "hd", "camera", "webcam", "mouse", "keyboard",
                "controller", "adapter", "card", "device", "chipset", "intel", "amd",
                "nvidia", "realtek", "broadcom", "qualcomm"
            };

            foreach (string keyword in deviceKeywords)
            {
                if (lowerText.Contains(keyword))
                {
                    return true; // ✅ CÓ THỂ LÀ DEVICE NAME
                }
            }

            // ✅ Nếu có format như "ABC(R) XYZ 123" thì có thể là device name
            if (text.Contains("(R)") || text.Contains("®") || text.Contains("™"))
            {
                return true;
            }

            return false;
        }

        // Cải tiến hàm làm sạch tên file
        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Unknown_Driver";

            // Thay thế các ký tự không hợp lệ
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string result = fileName;

            foreach (char c in invalidChars)
            {
                result = result.Replace(c, '_');
            }

            // Thay thế một số ký tự đặc biệt khác
            result = result.Replace(":", "_")
                           .Replace("*", "_")
                           .Replace("?", "_")
                           .Replace("\"", "_")
                           .Replace("<", "_")
                           .Replace(">", "_")
                           .Replace("|", "_");

            // Loại bỏ khoảng trắng thừa và thay bằng dấu gạch dưới
            result = Regex.Replace(result, @"\s+", " ").Trim();

            // Giới hạn độ dài tên thư mục (Windows có giới hạn 260 ký tự cho đường dẫn)
            if (result.Length > 100)
            {
                result = result.Substring(0, 100).Trim();
            }

            // Đảm bảo không kết thúc bằng dấu chấm hoặc khoảng trắng
            result = result.TrimEnd('.', ' ');

            // Nếu sau khi làm sạch mà rỗng, dùng tên mặc định
            if (string.IsNullOrEmpty(result))
                return "Unknown_Driver";

            return result;
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string[] lines =
            {
            "Phần mềm: IT Support Toolkit",
            "Phiên bản: 0.0.10.0",
            "Ngày phát hành: 15/08/2025",
            "Tác giả: Harry Hoang Le",
            "","",
            "Phần mềm thử nghiệm, các nút ẩn mờ đi là tính năng dự kiến sẽ phát triển, các nút màu xám và xanh nhạt là tính năng chưa được test đầy đủ, các nút màu xanh là đã chạy được cơ bản."
             };

            string versionInfo = string.Join(Environment.NewLine, lines);

            MessageBox.Show(versionInfo,
                            "Thông tin phiên bản",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            try
            {
                string edition = GetWindowsEdition();

                if (edition.Contains("Pro"))
                {
                    MessageBox.Show("Máy đã chạy Windows Pro. Không cần nâng cấp.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Nếu chưa phải bản Pro
                DialogResult result = MessageBox.Show(
                    $"Phiên bản hiện tại: {edition}\n\nBạn muốn nâng cấp lên Windows Pro bằng key mặc định không?",
                    "Nâng cấp Windows",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);

                // Yes = dùng key mặc định
                if (result == DialogResult.Yes)
                {
                    string defaultKey = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"; // Key mặc định Windows Pro
                    UpgradeWindows(defaultKey);
                }
                // No = nhập key tùy chỉnh
                else if (result == DialogResult.No)
                {
                    string userKey = Microsoft.VisualBasic.Interaction.InputBox("Nhập key Windows Pro của bạn:", "Nhập key", "");

                    if (string.IsNullOrWhiteSpace(userKey) || userKey.Length < 25)
                    {
                        MessageBox.Show("Key không hợp lệ. Vui lòng thử lại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    UpgradeWindows(userKey);
                }
                // Cancel = không làm gì
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private string GetWindowsEdition()
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject os in searcher.Get())
                {
                    return os["Caption"].ToString();
                }
            }
            return "Không xác định";
        }

        private void UpgradeWindows(string productKey)
        {
            try
            {
                string command = $"/C changepk.exe /productkey {productKey}";
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", command)
                {
                    Verb = "runas",
                    CreateNoWindow = true,
                    UseShellExecute = true
                };

                Process.Start(psi);
                MessageBox.Show("Đang tiến hành nâng cấp. Vui lòng chờ và khởi động lại máy nếu cần.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi nâng cấp: " + ex.Message);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                string edition = GetWindowsEdition();
                string licenseChannel = GetLicenseChannel();

                string message = $"Phiên bản hiện tại: {edition}\nLoại kích hoạt: {licenseChannel}";

                if (licenseChannel.Contains("Volume") && licenseChannel.Contains("KMS"))
                {
                    MessageBox.Show(message + "\n\nWindows đã được kích hoạt qua KMS. Không cần chuyển.", "Thông tin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    message + "\n\nBạn có muốn chuyển sang Volume KMS không?",
                    "Chuyển sang KMS",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string kmsKey = "W269N-WFGWX-YVC9B-4F6C9-T83GX"; // KMS Client Setup Key cho Windows Pro
                    RunCmd($"/c slmgr /ipk {kmsKey}");
                    RunCmd("/c slmgr /ato");

                    MessageBox.Show("Đã chuyển sang Volume KMS và kích hoạt lại Windows.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private string GetLicenseChannel()
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("powershell", "-Command \"(Get-WmiObject -query 'select * from SoftwareLicensingProduct where PartialProductKey is not null').Description\"")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                string result = "";
                using (Process proc = Process.Start(psi))
                {
                    result = proc.StandardOutput.ReadToEnd();
                }

                result = result.ToLower();

                if (result.Contains("kms"))
                    return "Volume - KMS";
                else if (result.Contains("mak"))
                    return "Volume - MAK";
                else if (result.Contains("retail"))
                    return "Retail";
                else
                    return "Không xác định";
            }
            catch
            {
                return "Không xác định";
            }
        }

        private void RunCmd(string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", arguments)
            {
                Verb = "runas",
                CreateNoWindow = true,
                UseShellExecute = true
            };
            Process.Start(psi);
        }

        // Biến lưu thông tin KMS server hiện tại (sẽ được lấy từ hệ thống)
        private string currentKmsServer = "";
        private string currentKmsPort = "";

        // Nút hiển thị thông tin KMS server hiện tại
        private void button17_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Lấy thông tin KMS server từ hệ thống
                GetCurrentKmsServer();

                string currentInfo;
                if (string.IsNullOrEmpty(currentKmsServer))
                {
                    currentInfo = "Chưa cấu hình KMS Server hoặc không thể truy xuất thông tin.";
                }
                else
                {
                    currentInfo = $"KMS Server hiện tại:\n" +
                                 $"Server: {currentKmsServer}\n" +
                                 $"Port: {(string.IsNullOrEmpty(currentKmsPort) ? "1688 (mặc định)" : currentKmsPort)}";
                }

                DialogResult result = MessageBox.Show(
                    currentInfo + "\n\nBạn có muốn thay đổi KMS Server không?",
                    "Thông tin KMS Server",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information
                );

                // Nếu người dùng muốn thay đổi
                if (result == DialogResult.Yes)
                {
                    ShowChangeKmsServerDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi truy xuất thông tin KMS: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Form để người dùng nhập KMS server mới
        private void ShowChangeKmsServerDialog()
        {
            Form changeForm = new Form()
            {
                Width = 390,
                Height = 190,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Thay đổi KMS Server",
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            Label serverLabel = new Label() { Left = 20, Top = 20, Width = 40, Text = "Server:" };
            TextBox serverTextBox = new TextBox() { Left = 80, Top = 17, Width = 250, Text = currentKmsServer };

            Label portLabel = new Label() { Left = 20, Top = 50, Width = 40, Text = "Port:" };
            TextBox portTextBox = new TextBox() { Left = 80, Top = 47, Width = 100, Text = "1688" }; // Port mặc định

            Button confirmButton = new Button() { Text = "Xác nhận", Left = 100, Width = 80, Top = 90, DialogResult = DialogResult.OK };
            Button cancelButton = new Button() { Text = "Hủy", Left = 220, Width = 80, Top = 90, DialogResult = DialogResult.Cancel };

            // Sự kiện xác nhận
            confirmButton.Click += (sender, e) => {
                string newServer = serverTextBox.Text.Trim();
                string newPort = portTextBox.Text.Trim();

                if (ValidateKmsInput(newServer, newPort))
                {
                    currentKmsServer = newServer;
                    currentKmsPort = newPort;

                    MessageBox.Show($"Đã thay đổi KMS Server thành công!\n" +
                                  $"Server: {currentKmsServer}\n" +
                                  $"Port: {currentKmsPort}",
                                  "Thành công",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);

                    // Gọi hàm áp dụng thay đổi (nếu cần)
                    ApplyKmsServerChange();

                    changeForm.Close();
                }
            };

            changeForm.Controls.Add(serverLabel);
            changeForm.Controls.Add(serverTextBox);
            changeForm.Controls.Add(portLabel);
            changeForm.Controls.Add(portTextBox);
            changeForm.Controls.Add(confirmButton);
            changeForm.Controls.Add(cancelButton);

            changeForm.AcceptButton = confirmButton;
            changeForm.CancelButton = cancelButton;

            changeForm.ShowDialog();
        }

        // Kiểm tra tính hợp lệ của input
        private bool ValidateKmsInput(string server, string port)
        {
            // Kiểm tra server không được rỗng
            if (string.IsNullOrWhiteSpace(server))
            {
                MessageBox.Show("Vui lòng nhập địa chỉ server!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Kiểm tra port
            if (string.IsNullOrWhiteSpace(port))
            {
                MessageBox.Show("Vui lòng nhập port!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Kiểm tra port là số và trong khoảng hợp lệ
            if (!int.TryParse(port, out int portNumber) || portNumber < 1 || portNumber > 65535)
            {
                MessageBox.Show("Port phải là số từ 1 đến 65535!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // Parse thông tin KMS từ output của slmgr (cải tiến)
        private bool ParseKmsInfoFromOutput(string output)
        {
            try
            {
                if (string.IsNullOrEmpty(output))
                    return false;

                // Các pattern có thể có trong output
                string[] patterns = {
                "KMS machine name",
                "KMS server name",
                "Key Management Service machine name",
                "Tên máy KMS"
            };

                string[] lines = output.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string line in lines)
                {
                    foreach (string pattern in patterns)
                    {
                        if (line.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Tìm phần sau dấu ":"
                            int colonIndex = line.IndexOf(':');
                            if (colonIndex > 0 && colonIndex < line.Length - 1)
                            {
                                string kmsInfo = line.Substring(colonIndex + 1).Trim();

                                if (!string.IsNullOrEmpty(kmsInfo) && kmsInfo != "<not set>" && kmsInfo.ToLower() != "not set")
                                {
                                    // Kiểm tra nếu có port trong cùng string
                                    if (kmsInfo.Contains(':') && !kmsInfo.StartsWith("http"))
                                    {
                                        string[] parts = kmsInfo.Split(':');
                                        currentKmsServer = parts[0].Trim();
                                        currentKmsPort = parts[1].Trim();
                                    }
                                    else
                                    {
                                        currentKmsServer = kmsInfo;
                                        currentKmsPort = "1688"; // Port mặc định
                                    }
                                    return true;
                                }
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi parse KMS info: {ex.Message}");
                return false;
            }
        }
        private void ApplyKmsServerChange()
        {
            try
            {
                // Sử dụng lệnh slmgr để cài đặt KMS server mới
                string kmsAddress = string.IsNullOrEmpty(currentKmsPort) || currentKmsPort == "1688"
                                  ? currentKmsServer
                                  : $"{currentKmsServer}:{currentKmsPort}";

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "slmgr.vbs",
                    Arguments = $"/skms {kmsAddress}",
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    Verb = "runas" // Chạy với quyền admin
                };

                using (Process process = Process.Start(psi))
                {
                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        Console.WriteLine($"Áp dụng KMS Server thành công: {currentKmsServer}:{currentKmsPort}");
                    }
                    else
                    {
                        throw new Exception($"Lệnh slmgr trả về mã lỗi: {process.ExitCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi áp dụng thay đổi KMS server:\n{ex.Message}\n\nVui lòng chạy ứng dụng với quyền Administrator.",
                              "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Truy xuất thông tin KMS server từ hệ thống Windows
        private void GetCurrentKmsServer()
        {
            bool success = false;
            string debugInfo = "";

            // Phương pháp 1: Sử dụng cscript để chạy slmgr.vbs
            try
            {
                debugInfo += "Thử phương pháp 1: cscript slmgr.vbs...\n";
                if (GetKmsUsingCScript())
                {
                    success = true;
                    debugInfo += "Thành công với cscript!\n";
                }
            }
            catch (Exception ex)
            {
                debugInfo += $"Phương pháp 1 lỗi: {ex.Message}\n";
            }

            // Phương pháp 2: Đọc từ Registry nếu phương pháp 1 thất bại
            if (!success)
            {
                try
                {
                    debugInfo += "Thử phương pháp 2: Registry...\n";
                    if (GetKmsFromRegistry())
                    {
                        success = true;
                        debugInfo += "Thành công với Registry!\n";
                    }
                }
                catch (Exception ex)
                {
                    debugInfo += $"Phương pháp 2 lỗi: {ex.Message}\n";
                }
            }

            // Phương pháp 3: Sử dụng WMI
            if (!success)
            {
                try
                {
                    debugInfo += "Thử phương pháp 3: WMI...\n";
                    if (GetKmsUsingWMI())
                    {
                        success = true;
                        debugInfo += "Thành công với WMI!\n";
                    }
                }
                catch (Exception ex)
                {
                    debugInfo += $"Phương pháp 3 lỗi: {ex.Message}\n";
                }
            }

            // Phương pháp 4: Sử dụng PowerShell
            if (!success)
            {
                try
                {
                    debugInfo += "Thử phương pháp 4: PowerShell...\n";
                    if (GetKmsUsingPowerShell())
                    {
                        success = true;
                        debugInfo += "Thành công với PowerShell!\n";
                    }
                }
                catch (Exception ex)
                {
                    debugInfo += $"Phương pháp 4 lỗi: {ex.Message}\n";
                }
            }

            if (!success)
            {
                currentKmsServer = "";
                currentKmsPort = "";
                // Hiển thị thông tin debug để biết nguyên nhân
                MessageBox.Show($"Debug Info:\n{debugInfo}", "Thông tin Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Phương pháp 1: Sử dụng cscript để chạy slmgr.vbs
        private bool GetKmsUsingCScript()
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "cscript.exe",
                    Arguments = "//nologo C:\\Windows\\System32\\slmgr.vbs /dlv",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (!string.IsNullOrEmpty(error))
                    {
                        Console.WriteLine($"CScript Error: {error}");
                    }

                    return ParseKmsInfoFromOutput(output);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetKmsUsingCScript Error: {ex.Message}");
                return false;
            }
        }

        // Phương pháp 2: Đọc từ Registry (cải tiến)
        private bool GetKmsFromRegistry()
        {
            try
            {
                // Thử đọc từ các vị trí Registry khác nhau
                string[] registryPaths = {
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\SoftwareProtectionPlatform",
                @"SYSTEM\CurrentControlSet\Services\sppsvc\Parameters"
            };

                foreach (string path in registryPaths)
                {
                    using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(path))
                    {
                        if (key != null)
                        {
                            object kmsServer = key.GetValue("KeyManagementServiceName");
                            object kmsPort = key.GetValue("KeyManagementServicePort");

                            if (kmsServer != null && !string.IsNullOrEmpty(kmsServer.ToString()))
                            {
                                currentKmsServer = kmsServer.ToString();
                                currentKmsPort = kmsPort?.ToString() ?? "1688";
                                return true;
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetKmsFromRegistry Error: {ex.Message}");
                return false;
            }
        }

        // Phương pháp 3: Sử dụng WMI
        private bool GetKmsUsingWMI()
        {
            try
            {
                using (var searcher = new System.Management.ManagementObjectSearcher("SELECT * FROM SoftwareLicensingService"))
                {
                    foreach (System.Management.ManagementObject obj in searcher.Get())
                    {
                        object kmsServer = obj["KeyManagementServiceMachine"];
                        object kmsPort = obj["KeyManagementServicePort"];

                        if (kmsServer != null && !string.IsNullOrEmpty(kmsServer.ToString()))
                        {
                            currentKmsServer = kmsServer.ToString();
                            currentKmsPort = kmsPort?.ToString() ?? "1688";
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetKmsUsingWMI Error: {ex.Message}");
                return false;
            }
        }

        // Phương pháp 4: Sử dụng PowerShell
        private bool GetKmsUsingPowerShell()
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-Command \"Get-WmiObject -Query 'SELECT * FROM SoftwareLicensingService' | Select-Object KeyManagementServiceMachine, KeyManagementServicePort\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    return ParsePowerShellOutput(output);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetKmsUsingPowerShell Error: {ex.Message}");
                return false;
            }
        }

        // Parse thông tin từ PowerShell output
        private bool ParsePowerShellOutput(string output)
        {
            try
            {
                string[] lines = output.Split('\n');
                foreach (string line in lines)
                {
                    if (line.Contains("KeyManagementServiceMachine") && line.Contains(":"))
                    {
                        string server = line.Split(':')[1].Trim();
                        if (!string.IsNullOrEmpty(server) && server != "null")
                        {
                            currentKmsServer = server;
                            currentKmsPort = "1688"; // Sẽ được cập nhật nếu tìm thấy port
                            return true;
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
