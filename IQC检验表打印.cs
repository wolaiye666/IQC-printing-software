// ==================================================
// 文件批量打印工具
// 作者: 王国强  Rev.A01
// 版本: A0 (完整版，两排底部布局)
// 搜索关系：或（任意关键词匹配即可）
// 图标说明：请将 Print.ico 放在程序同目录，否则使用默认图标
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
        private Label lblMovingNote;          // 移动备注标签
        private Timer moveTimer;               // 定时器

        // 授权相关常量
        private const string LicenseFileName = "license.dat";
        private const string SecretKey = "Wang2026Secret";
        private DateTime authorizedDate;

        // 移动参数
        private int moveDirection = 1;          // 1表示向右，-1表示向左
        private const int moveSpeed = 2;         // 每 Tick 移动像素
        private int leftMin = 12;
        private int leftMax;

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

            // 启动移动定时器
            moveTimer = new Timer();
            moveTimer.Interval = 30;             // 约33fps
            moveTimer.Tick += MoveTimer_Tick;
            moveTimer.Start();
        }

        /// <summary>
        /// 移动标签动画
        /// </summary>
        private void MoveTimer_Tick(object sender, EventArgs e)
        {
            if (lblMovingNote == null) return;

            int newLeft = lblMovingNote.Left + moveDirection * moveSpeed;
            if (newLeft < leftMin)
            {
                newLeft = leftMin;
                moveDirection = 1;
            }
            else if (newLeft > leftMax)
            {
                newLeft = leftMax;
                moveDirection = -1;
            }
            lblMovingNote.Left = newLeft;
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

        // 设置界面布局（两排底部）
        private void SetupUI()
        {
            // 窗体标题
            this.Text = "检验表批量打印工具";
            this.Size = new Size(750, 650);          // 高度增加以容纳两排
            this.StartPosition = FormStartPosition.CenterScreen;

            // 目录选择行
            Label lblDir = new Label { Text = "目录路径:", Location = new Point(12, 15), Size = new Size(70, 23) };
            txtDirectory = new TextBox { Location = new Point(88, 12), Size = new Size(490, 23) };
            btnScan = new Button { Text = "扫描", Location = new Point(584, 10), Size = new Size(75, 27) };
            btnScan.Click += BtnScan_Click;

            // 搜索行
            Label lblSearch = new Label { Text = "关键词:", Location = new Point(12, 50), Size = new Size(70, 23) };
            txtSearch = new TextBox { Location = new Point(88, 47), Size = new Size(410, 23) };
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

            // ========== 底部第一行（全选、已选计数、作者、邮箱、授权过期）==========
            int firstRowY = 420; // 第一行 Y 坐标

            btnSelectAll = new Button { Text = "全选", Location = new Point(12, firstRowY), Size = new Size(75, 27) };
            btnSelectAll.Click += BtnSelectAll_Click;

            lblSelectedCount = new Label
            {
                Text = "已选择 0 个文件",
                Location = new Point(100, firstRowY + 2), // 垂直居中微调
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };

            lblAuthorVersion = new Label
            {
                Text = "作者:王国强 Rev.A01",
                Location = new Point(230, firstRowY + 2),
                Size = new Size(150, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            lblEmail = new Label
            {
                Text = "Email: guoqiang.w@cn.interplex.com",
                Location = new Point(390, firstRowY + 2),
                Size = new Size(250, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Green,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            DateTime expireDate = authorizedDate.AddDays(30);
            lblExpireDate = new Label
            {
                Text = string.Format("授权至:{0:yyyy-MM-dd}", expireDate),
                Location = new Point(620, firstRowY + 2),
                Size = new Size(120, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Red,
                Font = new Font("Segoe UI", 9, FontStyle.Regular)
            };

            // ========== 底部第二行（批量打印按钮 + 移动备注标签）==========
            int secondRowY = firstRowY + 35; // 第一行下方35像素

            btnPrint = new Button { Text = "批量打印", Location = new Point(12, secondRowY), Size = new Size(90, 27) };
            btnPrint.Click += BtnPrint_Click;

            // 移动备注标签（棕色）
            lblMovingNote = new Label
            {
                Text = "备注:可批量搜索,中间空格区分.",
                Location = new Point(110, secondRowY + 2), // 紧挨按钮右侧
                Size = new Size(280, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Brown,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            // 计算最大左边界（避免超出窗体）
            leftMax = this.ClientSize.Width - lblMovingNote.Width - 10;

            // ========== 状态栏（底部） ==========
            lblStatus = new Label
            {
                Text = "就绪",
                Location = new Point(12, secondRowY + 40),
                Size = new Size(710, 50),
                BorderStyle = BorderStyle.Fixed3D,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // 将所有控件添加到窗体
            this.Controls.AddRange(new Control[] {
                lblDir, txtDirectory, btnScan,
                lblSearch, txtSearch, btnSearch,
                lblFiles, clbFiles,
                btnSelectAll, lblSelectedCount, lblAuthorVersion, lblEmail, lblExpireDate,
                btnPrint, lblMovingNote,
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

        // 批量打印（直接使用默认打印机，静默模式）
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
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = file,
                        Verb = "print",
                        CreateNoWindow = true,
                        UseShellExecute = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };
                    Process.Start(psi);
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
