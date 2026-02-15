// ==================================================
// 文件批量打印工具
// 作者: 王国强  Rev.A01
// 版本: A0 (完整版，增强授权 + 联网时间验证)
// 搜索关系：或（任意关键词匹配即可）
// 图标说明：请将 Print.ico 放在程序同目录，否则使用默认图标
// 创建日期: 2026-02-14
// ==================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace FileBatchPrinterGUI
{
    // 授权窗口（提示文字绿色）
    public class LicenseForm : Form
    {
        private TextBox txtLicenseCode;
        private Button btnOK;
        private Label lblPrompt;
        private Label lblContact;

        public string InputCode { get; private set; }

        public LicenseForm()
        {
            this.Text = "授权验证";
            this.Size = new Size(350, 180);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lblPrompt = new Label
            {
                Text = "请输入本月授权码:",
                Location = new Point(20, 20),
                Size = new Size(300, 20)
            };

            txtLicenseCode = new TextBox
            {
                Location = new Point(20, 45),
                Size = new Size(200, 23)
            };

            btnOK = new Button
            {
                Text = "验证",
                Location = new Point(230, 43),
                Size = new Size(75, 27)
            };
            btnOK.Click += BtnOK_Click;

            // 联系提示（绿色）
            lblContact = new Label
            {
                Text = "授权码请联络王国强获取",
                Location = new Point(20, 80),
                Size = new Size(300, 20),
                ForeColor = Color.Green,
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                TextAlign = ContentAlignment.MiddleLeft
            };

            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtLicenseCode);
            this.Controls.Add(btnOK);
            this.Controls.Add(lblContact);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            InputCode = txtLicenseCode.Text.Trim();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    // 主窗体
    public class Form1 : Form
    {
        // 存储可打印的文件列表（完整路径）
        private List<string> filesList = new List<string>();

        // 支持打印的文件后缀
        private readonly List<string> supportPrintExts = new List<string>
        {
            ".txt", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff"
        };

        // 控件声明
        private TextBox txtDirectory;
        private Button btnScan;
        private TextBox txtSearch;
        private Button btnSearch;
        private CheckedListBox clbFiles;
        private Button btnSelectAll;
        private Button btnPrint;
        private Label lblStatus;
        private Label lblSelectedCount;
        private Label lblAuthorVersion;
        private Label lblEmail;
        private Label lblExpireDate;

        // 授权相关常量
        private const string LicenseFileName = "license.dat";
        private const string SecretKey = "Wang2026Secret";
        private DateTime authorizedDate;

        // NTP 服务器列表
        private readonly string[] ntpServers = { "time.windows.com", "pool.ntp.org", "time.nist.gov" };

        public Form1()
        {
            if (!CheckLicense())
            {
                Environment.Exit(0);
                return;
            }

            // 加载自定义图标
            LoadCustomIcon();

            SetupUI();
        }

        /// <summary>
        /// 从网络获取标准时间（NTP 协议）
        /// </summary>
        private DateTime? GetNetworkTime()
        {
            // NTP 消息长度 48 字节
            byte[] ntpData = new byte[48];
            ntpData[0] = 0x1B; // 设置模式为客户端

            foreach (string ntpServer in ntpServers)
            {
                try
                {
                    using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                    {
                        socket.Connect(ntpServer, 123);
                        socket.ReceiveTimeout = 3000; // 3秒超时
                        socket.Send(ntpData);
                        socket.Receive(ntpData);
                    }

                    // 解析时间戳（从第40字节开始）
                    ulong intPart = (ulong)ntpData[40] << 24 | (ulong)ntpData[41] << 16 | (ulong)ntpData[42] << 8 | (ulong)ntpData[43];
                    ulong fractPart = (ulong)ntpData[44] << 24 | (ulong)ntpData[45] << 16 | (ulong)ntpData[46] << 8 | (ulong)ntpData[47];

                    ulong milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);
                    DateTime networkDateTime = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds((long)milliseconds);

                    return networkDateTime.ToLocalTime();
                }
                catch
                {
                    // 尝试下一个服务器
                }
            }
            return null; // 所有服务器都失败
        }

        /// <summary>
        /// 加载自定义图标（从程序目录加载 Print.ico）
        /// </summary>
        private void LoadCustomIcon()
        {
            try
            {
                string iconPath = Path.Combine(Application.StartupPath, "Print.ico");
                if (File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("图标加载失败: " + ex.Message);
            }
        }

        // 增强版授权检查（防时间回退 + 联网验证）
        private bool CheckLicense()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string licenseDir = Path.Combine(appData, "FileBatchPrinter");
            string licenseFile = Path.Combine(licenseDir, LicenseFileName);

            // 尝试获取网络时间
            DateTime? networkNow = GetNetworkTime();
            DateTime currentTime;
            if (networkNow.HasValue)
            {
                currentTime = networkNow.Value;
                // 可在状态栏显示（但此时UI尚未创建，可暂不处理）
            }
            else
            {
                currentTime = DateTime.Now;
                MessageBox.Show("无法连接到时间服务器，将使用本地系统时间进行授权验证。\n若本地时间不准确，可能影响授权。",
                    "网络时间同步失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            DateTime lastAuthDate;
            if (File.Exists(licenseFile))
            {
                string content = File.ReadAllText(licenseFile).Trim();
                if (DateTime.TryParse(content, out lastAuthDate))
                {
                    // 检测时间回退：如果当前时间小于上次授权日期，说明用户把系统时间调前了
                    if (currentTime.Date < lastAuthDate.Date)
                    {
                        // 删除旧授权文件，要求重新授权
                        File.Delete(licenseFile);
                        MessageBox.Show("检测到系统时间异常，授权已失效，请重新获取授权码。",
                            "授权失效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // 继续执行授权流程（弹出窗口）
                    }
                    else if ((currentTime.Date - lastAuthDate.Date).Days <= 30)
                    {
                        authorizedDate = lastAuthDate;
                        return true;
                    }
                }
            }

            // 无有效授权，弹出授权窗口（使用网络时间或本地时间生成授权码？应与之前保持一致，仍用本地时间生成，因为授权码基于本地年月）
            using (LicenseForm licenseForm = new LicenseForm())
            {
                if (licenseForm.ShowDialog() == DialogResult.OK)
                {
                    // 注意：生成授权码仍使用本地时间，以保证与之前逻辑一致
                    string expectedCode = GenerateLicenseCode();
                    if (licenseForm.InputCode == expectedCode)
                    {
                        if (!Directory.Exists(licenseDir))
                            Directory.CreateDirectory(licenseDir);
                        // 写入当前时间（使用网络时间或本地时间，但以网络时间为准）
                        File.WriteAllText(licenseFile, currentTime.ToString("yyyy-MM-dd"));
                        authorizedDate = currentTime;
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("授权码错误！", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            return false;
        }

        // 生成当月授权码（基于本地时间，与之前一致）
        private string GenerateLicenseCode()
        {
            string yearMonth = DateTime.Now.ToString("yyyyMM");
            string raw = yearMonth + SecretKey;
            using (MD5 md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(raw));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 4; i++)
                {
                    sb.Append(hash[i].ToString("x2"));
                }
                return sb.ToString().ToUpper();
            }
        }

        // 设置界面布局
        private void SetupUI()
        {
            // 窗体标题
            this.Text = "检验表批量打印工具";
            this.Size = new Size(800, 590);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 目录选择行
            Label lblDir = new Label { Text = "目录路径:", Location = new Point(12, 15), Size = new Size(70, 23) };
            txtDirectory = new TextBox { Location = new Point(88, 12), Size = new Size(540, 23) };
            btnScan = new Button { Text = "扫描", Location = new Point(634, 10), Size = new Size(75, 27) };
            btnScan.Click += BtnScan_Click;

            // 搜索行
            Label lblSearch = new Label { Text = "关键词:", Location = new Point(12, 50), Size = new Size(70, 23) };
            txtSearch = new TextBox { Location = new Point(88, 47), Size = new Size(460, 23) };
            btnSearch = new Button { Text = "搜索", Location = new Point(554, 45), Size = new Size(75, 27) };
            btnSearch.Click += BtnSearch_Click;

            // 文件列表标题行（左侧标题 + 右侧Email）
            Label lblFiles = new Label
            {
                Text = "文件列表（勾选要打印的文件）:",
                Location = new Point(12, 85),
                Size = new Size(200, 23)
            };

            // 右侧Email标签（绿色）
            lblEmail = new Label
            {
                Text = "Email: guoqiang.w@cn.interplex.com",
                Location = new Point(500, 85),
                Size = new Size(250, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Green,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            // 文件列表（带复选框）
            clbFiles = new CheckedListBox
            {
                Location = new Point(12, 110),
                Size = new Size(760, 300),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            clbFiles.Format += (s, e) =>
            {
                if (e.ListItem is string path)
                    e.Value = Path.GetFileName(path);
            };
            clbFiles.ItemCheck += ClbFiles_ItemCheck;

            // ========== 底部一行布局 ==========
            int bottomY = 420;

            // 全选按钮
            btnSelectAll = new Button { Text = "全选", Location = new Point(12, bottomY), Size = new Size(75, 27) };
            btnSelectAll.Click += BtnSelectAll_Click;

            // 已选计数标签
            lblSelectedCount = new Label
            {
                Text = "已选择 0 个文件",
                Location = new Point(100, bottomY + 2),
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 作者版本标签
            lblAuthorVersion = new Label
            {
                Text = "作者:王国强 Rev.A01",
                Location = new Point(230, bottomY + 2),
                Size = new Size(150, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            // 授权过期标签
            DateTime expireDate = authorizedDate.AddDays(30);
            lblExpireDate = new Label
            {
                Text = string.Format("授权至:{0:yyyy-MM-dd}", expireDate),
                Location = new Point(400, bottomY + 2),
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            // 批量打印按钮
            btnPrint = new Button { Text = "批量打印", Location = new Point(690, bottomY), Size = new Size(90, 27) };
            btnPrint.Click += BtnPrint_Click;

            // 状态栏
            lblStatus = new Label
            {
                Text = "就绪",
                Location = new Point(12, bottomY + 40),
                Size = new Size(760, 50),
                BorderStyle = BorderStyle.Fixed3D,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 将所有控件添加到窗体
            this.Controls.AddRange(new Control[] {
                lblDir, txtDirectory, btnScan,
                lblSearch, txtSearch, btnSearch,
                lblFiles, lblEmail, clbFiles,
                btnSelectAll, lblSelectedCount, lblAuthorVersion, lblExpireDate,
                btnPrint,
                lblStatus
            });
        }

        // 更新已选计数
        private void ClbFiles_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)(() =>
            {
                lblSelectedCount.Text = string.Format("已选择 {0} 个文件", clbFiles.CheckedItems.Count);
            }));
        }

        // 全选/取消全选
        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0) return;

            bool allChecked = clbFiles.CheckedItems.Count == clbFiles.Items.Count;
            for (int i = 0; i < clbFiles.Items.Count; i++)
            {
                clbFiles.SetItemChecked(i, !allChecked);
            }
            btnSelectAll.Text = allChecked ? "全选" : "取消全选";
        }

        // 扫描目录
        private void BtnScan_Click(object sender, EventArgs e)
        {
            string dirPath = txtDirectory.Text.Trim();
            if (string.IsNullOrEmpty(dirPath))
            {
                MessageBox.Show("请输入目录路径！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!Directory.Exists(dirPath))
            {
                MessageBox.Show("目录不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            filesList.Clear();
            clbFiles.Items.Clear();
            lblStatus.Text = "正在扫描...";

            try
            {
                string[] files = Directory.GetFiles(dirPath, "*.*", SearchOption.AllDirectories);
                int count = 0;
                foreach (string file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (supportPrintExts.Contains(ext))
                    {
                        filesList.Add(file);
                        clbFiles.Items.Add(file);
                        count++;
                    }
                    if (count % 10 == 0)
                        Application.DoEvents();
                }
                lblStatus.Text = string.Format("扫描完成！共找到 {0} 个可打印文件。", filesList.Count);
                lblSelectedCount.Text = "已选择 0 个文件";
                btnSelectAll.Text = "全选";
            }
            catch (Exception ex)
            {
                filesList.Clear();
                clbFiles.Items.Clear();
                lblStatus.Text = string.Format("扫描出错: {0}", ex.Message);
            }
        }

        // 搜索文件（多关键词空格分隔，或关系）
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (filesList.Count == 0)
            {
                MessageBox.Show("请先扫描目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string keywords = txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(keywords))
            {
                clbFiles.Items.Clear();
                foreach (string file in filesList)
                    clbFiles.Items.Add(file);
                lblStatus.Text = "已显示全部文件。";
                return;
            }

            string[] keywordList = keywords.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> results = new List<string>();

            // 或关系：只要文件包含任意一个关键词即加入结果
            foreach (string file in filesList)
            {
                bool match = false;
                foreach (string keyword in keywordList)
                {
                    if (file.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        match = true;
                        break;
                    }
                }
                if (match)
                    results.Add(file);
            }

            clbFiles.Items.Clear();
            foreach (string file in results)
                clbFiles.Items.Add(file);

            lblStatus.Text = string.Format("搜索完成，共 {0} 个文件匹配。", results.Count);
            lblSelectedCount.Text = string.Format("已选择 {0} 个文件", clbFiles.CheckedItems.Count);
        }

        /// <summary>
        /// 打印文本文件（静默打印，不弹出任何窗口）
        /// </summary>
        private void PrintTextFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath, Encoding.Default);

            PrintDocument pd = new PrintDocument();
            pd.PrintPage += (sender, e) =>
            {
                float yPos = 0;
                int count = 0;
                float leftMargin = e.MarginBounds.Left;
                float topMargin = e.MarginBounds.Top;
                string line = null;

                while (count < lines.Length)
                {
                    line = lines[count];
                    yPos = topMargin + (count * pd.DefaultPageSettings.PrinterResolution.Y / 100 * 0.3f);
                    e.Graphics.DrawString(line, new Font("宋体", 10), Brushes.Black, leftMargin, yPos, new StringFormat());
                    count++;
                    if (yPos > e.MarginBounds.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }
                }
                e.HasMorePages = false;
            };

            try
            {
                pd.Print();
            }
            catch (Exception ex)
            {
                throw new Exception("打印文本失败: " + ex.Message);
            }
        }

        // 批量打印
        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (clbFiles.CheckedItems.Count == 0)
            {
                MessageBox.Show("请先勾选要打印的文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int success = 0, fail = 0;
            lblStatus.Text = "开始提交打印任务...";
            Application.DoEvents();

            foreach (var item in clbFiles.CheckedItems)
            {
                string file = item as string;
                try
                {
                    string ext = Path.GetExtension(file).ToLower();

                    if (ext == ".txt")
                    {
                        PrintTextFile(file);
                    }
                    else
                    {
                        ProcessStartInfo psi = new ProcessStartInfo
                        {
                            FileName = file,
                            Verb = "print",
                            CreateNoWindow = true,
                            UseShellExecute = true,
                            WindowStyle = ProcessWindowStyle.Hidden
                        };
                        Process.Start(psi);
                    }
                    success++;
                    System.Threading.Thread.Sleep(1500);
                }
                catch (Exception ex)
                {
                    fail++;
                    lblStatus.Text = string.Format("打印失败: {0} - {1}", Path.GetFileName(file), ex.Message);
                    Application.DoEvents();
                }
            }

            lblStatus.Text = string.Format("打印任务提交完成！成功: {0}, 失败: {1}。请查看系统打印机队列。", success, fail);
            MessageBox.Show(string.Format("打印完成！成功 {0} 个，失败 {1} 个。", success, fail), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // 应用程序入口点
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
