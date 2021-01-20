using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;

namespace JustinSoft
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.SandC.SelectedIndex = 0;
            this.CompanyList.SelectedIndex = 0;
            this.DoVatDate.SelectedDate = DateTime.Today;
            this.OpenOrNot.IsChecked = true;
            this.DoRun.IsEnabled = false;
        }

        //Excel 路径
        string filePath = null;

        //订单数组
        List<MonthlyStatement> ms = new List<MonthlyStatement>();

        //选中的订单数组
        List<MonthlyStatement> Cms = new List<MonthlyStatement>();

        //业务员集合
        List<string> Sales = new List<string>();

        //获取对账月份
        DateTime VDate = new DateTime();

        private void Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OpenExcel = new OpenFileDialog
            {
                Title = "选择从K3文件中导出的订单总表。", //对话框标题
                Filter = "*.xls文件,*.xlsx文件|*.xls;*.xlsx"//对话框格式筛选
            };

            OpenExcel.ShowDialog(); // 弹出对话框
            filePath = OpenExcel.FileName; //获取用户选择的Excel路径
            this.ExcelPath.Text = filePath; //路径显示到界面

            //创建json对象
            JavaScriptSerializer OrderSerializer = new JavaScriptSerializer() { MaxJsonLength = Int32.MaxValue };

            //获取Excel 转换为 DataTable，再转为的json
            string Json = Dt.Data2Json(Excel.Excel2Table(filePath, 0));

            //json转化为自定义数组
            ms = OrderSerializer.Deserialize<List<MonthlyStatement>>(Json);

            //业务员 数组集合所有业务员
            Sales.Add("全部业务员");

            //去重显示业务员到界面下拉框
            foreach (var item in ms)
            {
                if (!Sales.Contains(item.业务员))
                {
                    Sales.Add(item.业务员);
                }
                //从Excel表中获取实际对账时间
                VDate = DateTime.FromOADate(Convert.ToDouble(item.日期));
            }
            //界面显示业务员供用户选择
            this.SandC.ItemsSource = Sales;

            //界面显示默认事件
            this.VatDate.SelectedDate = VDate;

            //界面显示 发现的订单数
            this.ExcelPath.Text += "      -共发现有" + ms.Count + "笔订单数据。";
        }

        private void SandC_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //客户集合
            List<string> Customs = new List<string>();

            //根据用户选择的业务员刷新后的该业务员的客户集合
            foreach (var item in ms)
            {
                if (item.业务员 == this.SandC.SelectedItem.ToString())
                {
                    if (!Customs.Contains(item.购货单位))
                    {
                        Customs.Add(item.购货单位);
                    }
                }
                
                //用户选择全部业务员的时候展示所有客户
                if (this.SandC.SelectedItem.ToString() == "全部业务员")
                {
                    if (!Customs.Contains(item.购货单位))
                    {
                        Customs.Add(item.购货单位);
                    }
                }
            }

            //绑定刷新后的客户到界面显示
            this.CompanyList.ItemsSource = Customs;

            //排序
            this.CompanyList.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("", System.ComponentModel.ListSortDirection.Descending));
            this.Res.Text = "当前列表有 " + Customs.Count + " 家客户。";
        }

        [Obsolete]
        private void DoRun_Click(object sender, RoutedEventArgs e)
        {
            this.Res.Text = "正在进行处理，请稍后...";
            //选中的客户
            List<string> ChoosedCustoms = new List<string>();

            foreach (var item in CompanyList.SelectedItems)
            {
                ChoosedCustoms.Add(item.ToString());
            }

            foreach (var item in ChoosedCustoms)
            {
                foreach (var oder in ms)
                {
                    if (oder.购货单位 == item)
                    {
                        Cms.Add(oder);
                    }
                }
            }

            //数组给处理函数
            DoVat dv = new DoVat();
            dv.SplitVat(Cms, VatTime: (DateTime)VatDate.SelectedDate, doTime: (DateTime)DoVatDate.SelectedDate,
                contactPeole: ContactPeole.Text, Okpeole: OKPeople.Text);

            this.Res.Text = "已经处理完成全部任务。";

            if (OpenOrNot.IsChecked == true)
            {
                System.Diagnostics.Process.Start(@"D:\对账单\");
            }

            if (SaveOneFoder.IsChecked == true)
            {
                DirectoryInfo theFolder = new DirectoryInfo(@"D:\对账单\");
                DirectoryInfo[] dirInfo = theFolder.GetDirectories();
                //遍历文件夹
                foreach (DirectoryInfo NextFolder in dirInfo)
                {
                    FileInfo[] fileInfo = NextFolder.GetFiles();

                    //遍历文件
                    foreach (FileInfo NextFile in fileInfo)  
                    {
                        if (!Directory.Exists(@"D:\对账单\总的对账单\"))
                        {
                            Directory.CreateDirectory(@"D:\对账单\总的对账单\");
                        }
                        //复制所有文件
                        File.Copy(NextFile.FullName, @"D:\对账单\总的对账单\" + System.IO.Path.GetFileName(NextFile.FullName), true);
                    }
                }
            }

            //清空结合给第二次使用做准备
            ChoosedCustoms.Clear();
        }

        private void ChooseAll_Checked(object sender, RoutedEventArgs e)
        {
            CompanyList.SelectAll();
        }

        private void ChooseAll_Unchecked(object sender, RoutedEventArgs e)
        {
            CompanyList.UnselectAll();
        }

        private void Fbt_Click(object sender, RoutedEventArgs e)
        {
            //查找结果
            List<string> foundCompany = new List<string>();

            foreach (var item in ms)
            {
                if (item.购货单位.Contains(fountC.Text.ToString()))
                {
                    if (!foundCompany.Contains(item.购货单位))
                    {
                        foundCompany.Add(item.购货单位);
                    }
                }
            }
            this.Res.Text = "符合条件的 " + foundCompany.Count + " 家客户。";
            if (foundCompany.Count != 0)
            {
                CompanyList.ItemsSource = foundCompany;
                this.CompanyList.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("", System.ComponentModel.ListSortDirection.Descending));
            }

            else
            {
                MessageBox.Show("变更查找关键字再试。", "未能查询到包含输入关键词的公司名。");
            }
        }

        private void SaveOneFoder_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Pw.Password == "19901205")
            {
                //密码正确，启用运行按钮
                this.DoRun.IsEnabled = true;
            }
            else
            {
                //密码错误，禁用运行按钮并弹出提示
                this.DoRun.IsEnabled = false;
                MessageBoxResult mr = MessageBox.Show("密码不正确，禁止使用。点击 是 重新输入，点击 否 退出程序。","检测输入了错误的密码",MessageBoxButton.YesNo, MessageBoxImage.Warning);
                
                //用户选择性退出程序
                if (mr == MessageBoxResult.No)
                {
                    this.Close();
                }
                else
                {
                    return;
                }
            }
        }
    }
}
