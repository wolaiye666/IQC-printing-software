// ==================================================
// 文件批量打印工具
// 作者: 王国强
// 版本: A0 (完整版，含每月授权)
// 创建日期: 2026-02-14
// ==================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace FileBatchPrinterGUI
{
    // 授权窗口
    public class LicenseForm : Form
    {
        private TextBox txtLicenseCode;
        private Button btnOK;
        private Label lblPrompt;

        public string InputCode { get; private set; }

        public LicenseForm()
        {
            this.Text = "授权验证";
            this.Size = new Size(350, 150);
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
                Location = new Point(20, 50),
                Size = new Size(200, 23)
            };

            btnOK = new Button
            {
                Text = "验证",
                Location = new Point(230, 48),
                Size = new Size(75, 27)
            };
            btnOK.Click += BtnOK_Click;

            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtLicenseCode);
            this.Controls.Add(btnOK);
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
        private Label lblExpireDate;

        // 授权相关常量
        private const string LicenseFileName = "license.dat";
        private const string SecretKey = "Wang2026Secret";
        private DateTime authorizedDate;

        public Form1()
        {
            if (!CheckLicense())
            {
                Environment.Exit(0);
                return;
            }
            SetupUI();
        }

        // 授权检查
        private bool CheckLicense()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string licenseDir = Path.Combine(appData, "FileBatchPrinter");
            string licenseFile = Path.Combine(licenseDir, LicenseFileName);

            DateTime lastAuthDate;
            if (File.Exists(licenseFile))
            {
                string content = File.ReadAllText(licenseFile).Trim();
                if (DateTime.TryParse(content, out lastAuthDate))
                {
                    if ((DateTime.Now.Date - lastAuthDate.Date).Days <= 30)
                    {
                        authorizedDate = lastAuthDate;
                        return true;
                    }
                }
            }

            using (LicenseForm licenseForm = new LicenseForm())
            {
                if (licenseForm.ShowDialog() == DialogResult.OK)
                {
                    string expectedCode = GenerateLicenseCode();
                    if (licenseForm.InputCode == expectedCode)
                    {
                        if (!Directory.Exists(licenseDir))
                            Directory.CreateDirectory(licenseDir);
                        File.WriteAllText(licenseFile, DateTime.Now.ToString("yyyy-MM-dd"));
                        authorizedDate = DateTime.Now;
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

        // 生成当月授权码
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
            this.Text = "文件批量打印工具 (授权版)";
            this.Size = new Size(750, 560);
            this.StartPosition = FormStartPosition.CenterScreen;

            // 目录选择行
            Label lblDir = new Label { Text = "目录路径:", Location = new Point(12, 15), Size = new Size(60, 23) };
            txtDirectory = new TextBox { Location = new Point(78, 12), Size = new Size(500, 23) };
            btnScan = new Button { Text = "扫描", Location = new Point(584, 10), Size = new Size(75, 27) };
            btnScan.Click += BtnScan_Click;

            // 搜索行
            Label lblSearch = new Label { Text = "关键词:", Location = new Point(12, 50), Size = new Size(60, 23) };
            txtSearch = new TextBox { Location = new Point(78, 47), Size = new Size(420, 23) };
            btnSearch = new Button { Text = "搜索", Location = new Point(504, 45), Size = new Size(75, 27) };
            btnSearch.Click += BtnSearch_Click;

            // 文件列表（带复选框）
            Label lblFiles = new Label { Text = "文件列表（勾选要打印的文件）:", Location = new Point(12, 85), Size = new Size(200, 23) };
            clbFiles = new CheckedListBox
            {
                Location = new Point(12, 110),
                Size = new Size(710, 300),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            clbFiles.Format += (s, e) =>
            {
                if (e.ListItem is string path)
                    e.Value = Path.GetFileName(path);
            };
            clbFiles.ItemCheck += ClbFiles_ItemCheck;

            // 全选/取消全选按钮
            btnSelectAll = new Button { Text = "全选", Location = new Point(12, 420), Size = new Size(75, 27) };
            btnSelectAll.Click += BtnSelectAll_Click;

            // 已选文件计数标签
            lblSelectedCount = new Label
            {
                Text = "已选择 0 个文件",
                Location = new Point(100, 424),
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 作者版本标签
            lblAuthorVersion = new Label
            {
                Text = "作者:王国强 版本A0",
                Location = new Point(230, 424),
                Size = new Size(200, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            // 授权过期时间标签
            DateTime expireDate = authorizedDate.AddDays(30);
            lblExpireDate = new Label
            {
                Text = string.Format("授权至:{0:yyyy-MM-dd}", expireDate),
                Location = new Point(440, 424),
                Size = new Size(180, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            // 批量打印按钮
            btnPrint = new Button { Text = "批量打印", Location = new Point(620, 420), Size = new Size(90, 27) };
            btnPrint.Click += BtnPrint_Click;

            // 状态栏
            lblStatus = new Label
            {
                Text = "就绪",
                Location = new Point(12, 470),
                Size = new Size(710, 50),
                BorderStyle = BorderStyle.Fixed3D,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 将所有控件添加到窗体
            this.Controls.AddRange(new Control[] {
                lblDir, txtDirectory, btnScan,
                lblSearch, txtSearch, btnSearch,
                lblFiles, clbFiles,
                btnSelectAll, lblSelectedCount, lblAuthorVersion, lblExpireDate, btnPrint,
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

        // 搜索文件
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

            foreach (string file in filesList)
            {
                bool match = true;
                foreach (string keyword in keywordList)
                {
                    if (file.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        match = false;
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
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = file,
                        Verb = "print",
                        CreateNoWindow = true,
                        UseShellExecute = true
                    });
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
    static class Program
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
            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtLicenseCode);
            this.Controls.Add(btnOK);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            InputCode = txtLicenseCode.Text.Trim();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    public partial class Form1 : Form
    {
        private List<string> filesList = new List<string>();
        private readonly List<string> supportPrintExts = new List<string>
        {
            ".txt", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
            ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff"
        };

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
        private Label lblExpireDate;

        private const string LicenseFileName = "license.dat";
        private const string SecretKey = "Wang2026Secret";
        private DateTime authorizedDate;

        public Form1()
{
    if (!CheckLicense()) { Environment.Exit(0); return; }
    // 删除 InitializeComponent(); 因为 UI 已由 SetupUI 初始化
    SetupUI();
}

        private bool CheckLicense()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string licenseDir = Path.Combine(appData, "FileBatchPrinter");
            string licenseFile = Path.Combine(licenseDir, LicenseFileName);
            DateTime lastAuthDate;
            if (File.Exists(licenseFile))
            {
                string content = File.ReadAllText(licenseFile).Trim();
                if (DateTime.TryParse(content, out lastAuthDate))
                {
                    if ((DateTime.Now.Date - lastAuthDate.Date).Days <= 30)
                    {
                        authorizedDate = lastAuthDate;
                        return true;
                    }
                }
            }
            using (LicenseForm licenseForm = new LicenseForm())
            {
                if (licenseForm.ShowDialog() == DialogResult.OK)
                {
                    string expectedCode = GenerateLicenseCode();
                    if (licenseForm.InputCode == expectedCode)
                    {
                        if (!Directory.Exists(licenseDir)) Directory.CreateDirectory(licenseDir);
                        File.WriteAllText(licenseFile, DateTime.Now.ToString("yyyy-MM-dd"));
                        authorizedDate = DateTime.Now;
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

        private string GenerateLicenseCode()
        {
            string yearMonth = DateTime.Now.ToString("yyyyMM");
            string raw = yearMonth + SecretKey;
            using (MD5 md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(raw));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < 4; i++) sb.Append(hash[i].ToString("x2"));
                return sb.ToString().ToUpper();
            }
        }

        private void SetupUI()
        {
            this.Text = "文件批量打印工具 (授权版)";
            this.Size = new Size(750, 560);
            this.StartPosition = FormStartPosition.CenterScreen;

            Label lblDir = new Label { Text = "目录路径:", Location = new Point(12, 15), Size = new Size(60, 23) };
            txtDirectory = new TextBox { Location = new Point(78, 12), Size = new Size(500, 23) };
            btnScan = new Button { Text = "扫描", Location = new Point(584, 10), Size = new Size(75, 27) };
            btnScan.Click += BtnScan_Click;

            Label lblSearch = new Label { Text = "关键词:", Location = new Point(12, 50), Size = new Size(60, 23) };
            txtSearch = new TextBox { Location = new Point(78, 47), Size = new Size(420, 23) };
            btnSearch = new Button { Text = "搜索", Location = new Point(504, 45), Size = new Size(75, 27) };
            btnSearch.Click += BtnSearch_Click;

            Label lblFiles = new Label { Text = "文件列表（勾选要打印的文件）:", Location = new Point(12, 85), Size = new Size(200, 23) };
            clbFiles = new CheckedListBox
            {
                Location = new Point(12, 110),
                Size = new Size(710, 300),
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            clbFiles.Format += (s, e) => { if (e.ListItem is string path) e.Value = Path.GetFileName(path); };
            clbFiles.ItemCheck += ClbFiles_ItemCheck;

            btnSelectAll = new Button { Text = "全选", Location = new Point(12, 420), Size = new Size(75, 27) };
            btnSelectAll.Click += BtnSelectAll_Click;

            lblSelectedCount = new Label
            {
                Text = "已选择 0 个文件",
                Location = new Point(100, 424),
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            lblAuthorVersion = new Label
            {
                Text = "作者:王国强 版本A0",
                Location = new Point(230, 424),
                Size = new Size(200, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            DateTime expireDate = authorizedDate.AddDays(30);
            lblExpireDate = new Label
            {
                Text = string.Format("授权至:{0:yyyy-MM-dd}", expireDate),
                Location = new Point(440, 424),
                Size = new Size(180, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            btnPrint = new Button { Text = "批量打印", Location = new Point(620, 420), Size = new Size(90, 27) };
            btnPrint.Click += BtnPrint_Click;

            lblStatus = new Label
            {
                Text = "就绪",
                Location = new Point(12, 470),
                Size = new Size(710, 50),
                BorderStyle = BorderStyle.Fixed3D,
                TextAlign = ContentAlignment.MiddleLeft
            };

            this.Controls.AddRange(new Control[] {
                lblDir, txtDirectory, btnScan,
                lblSearch, txtSearch, btnSearch,
                lblFiles, clbFiles,
                btnSelectAll, lblSelectedCount, lblAuthorVersion, lblExpireDate, btnPrint,
                lblStatus
            });
        }

        private void ClbFiles_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)(() =>
            {
                lblSelectedCount.Text = string.Format("已选择 {0} 个文件", clbFiles.CheckedItems.Count);
            }));
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            if (clbFiles.Items.Count == 0) return;
            bool allChecked = clbFiles.CheckedItems.Count == clbFiles.Items.Count;
            for (int i = 0; i < clbFiles.Items.Count; i++)
                clbFiles.SetItemChecked(i, !allChecked);
            btnSelectAll.Text = allChecked ? "全选" : "取消全选";
        }

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
                    if (count % 10 == 0) Application.DoEvents();
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
                foreach (string file in filesList) clbFiles.Items.Add(file);
                lblStatus.Text = "已显示全部文件。";
                return;
            }
            string[] keywordList = keywords.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> results = new List<string>();
            foreach (string file in filesList)
            {
                bool match = true;
                foreach (string keyword in keywordList)
                {
                    if (file.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) == -1)
                    { match = false; break; }
                }
                if (match) results.Add(file);
            }
            clbFiles.Items.Clear();
            foreach (string file in results) clbFiles.Items.Add(file);
            lblStatus.Text = string.Format("搜索完成，共 {0} 个文件匹配。", results.Count);
            lblSelectedCount.Text = string.Format("已选择 {0} 个文件", clbFiles.CheckedItems.Count);
        }

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
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = file,
                        Verb = "print",
                        CreateNoWindow = true,
                        UseShellExecute = true
                    });
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
}
