using System;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Windows.Input;
using DataProcessing.Model;

namespace DataProcessing
{
    /// <summary>
    /// Interaction logic for Page1.xaml
    /// </summary>
    public partial class thietlapHeSo : Page
    {
        public static string startdatetime = "", enddatetime = "";
        public static string limit = "";
        public static string txtpath = "";
        public static int n = 0;
        public static int ncolor = 0;
        public static Boolean checkstop = false, findmax = true;
        Controller.ExcelController excelcontroller = new Controller.ExcelController();
        Controller.AlgorithmController tlhscontroller = new Controller.AlgorithmController();
        Controller.OutputController outcontroller = new Controller.OutputController();
        string color1, color2, color3, color4, color5;
        Boolean group2 = false, group3 = false, group4 = false, group5 = false;
        Boolean colorgroup0 = false, colorgroup1 = false, colorgroup2 = false, colorgroup3 = false, colorgroup4 = false, colorgroup5 = false;

        public thietlapHeSo()
        {
            InitializeComponent();
        }

        public void setProcess()
        {
            int i = 1;

            while (i < 107)
            {
                i++;

            }
        }

        public void comboBoxValue(ComboBox a, string[] array)
        {
            a.ItemsSource = array;
        }


        /// <summary>
        /// Mở đường dẫn đến file xls, xlsx và điền đường dẫn vào textbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public async void browseFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xls";

            openfile.Filter = "(.xls)|*.xls|(.xlsx)|*.xlsx";

            var browsefile = openfile.ShowDialog();

            if (browsefile == true)
            {
                searchbutton.Text = "ĐANG TINH CHỈNH ...";
                MessageBox.Show("Bạn vừa nhập đường dẫn: " + openfile.FileName);
                txtFilePath.Text = openfile.FileName;
                txtpath = openfile.FileName;
                excelcontroller.getColorAndDate(openfile.FileName);
                searchbutton.Text = "TIẾN HÀNH TÌM KIẾM";
                date1.ItemsSource = excelcontroller.fillDateTime();
                date2.ItemsSource = excelcontroller.fillDateTime();
                combo1.ItemsSource = excelcontroller.fillColorCombobox();
                combo2.ItemsSource = excelcontroller.fillColorCombobox();
                combo3.ItemsSource = excelcontroller.fillColorCombobox();
                combo4.ItemsSource = excelcontroller.fillColorCombobox();
                combo5.ItemsSource = excelcontroller.fillColorCombobox();
            }


        }


        /// <summary>
        /// Bắt đầu tìm kiếm
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public async void startSearch(object sender, RoutedEventArgs e)
        {
            startdatetime = date1.SelectedValue == null ? "" : date1.SelectedValue.ToString();
            enddatetime = date2.SelectedValue == null ? "" : date2.SelectedValue.ToString();
            limit = inputvalue.Text;
            Boolean blankcolor = false;
            Boolean checkinputcolor = false;


            if (ncolor != 0)
            {
                if (ncolor == 1)
                {
                    if (combo1.SelectedIndex == -1)
                    {
                        blankcolor = true;
                    }
                }
                if (ncolor == 2)
                {
                    if (combo1.SelectedIndex == -1 || combo2.SelectedIndex == -1)
                    {
                        blankcolor = true;
                    }
                    else
                    {
                        if (combo1.SelectedValue.ToString() == combo2.SelectedValue.ToString())
                        {
                            checkinputcolor = true;
                        }
                    }
                }
                else if (ncolor == 3)
                {
                    if (combo1.SelectedIndex == -1 || combo2.SelectedIndex == -1 || combo3.SelectedIndex == -1)
                    {
                        blankcolor = true;
                    }
                    else
                    {
                        if ((combo1.SelectedValue.ToString() == combo2.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo3.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo3.SelectedValue.ToString()))
                        {
                            checkinputcolor = true;
                        }
                    }

                }
                else if (ncolor == 4)
                {
                    if (combo1.SelectedIndex == -1 || combo2.SelectedIndex == -1 || combo3.SelectedIndex == -1
                        || combo4.SelectedIndex == -1)
                    {
                        blankcolor = true;
                    }
                    else
                    {
                        if ((combo1.SelectedValue.ToString() == combo2.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo3.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo4.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo3.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo4.SelectedValue.ToString())
                            || (combo3.SelectedValue.ToString() == combo4.SelectedValue.ToString()))
                        {
                            checkinputcolor = true;
                        }
                    }

                }
                else if (ncolor == 5)
                {
                    if (combo1.SelectedIndex == -1 || combo2.SelectedIndex == -1 || combo3.SelectedIndex == -1
                       || combo4.SelectedIndex == -1 || combo5.SelectedIndex == -1)
                    {
                        blankcolor = true;
                    }
                    else
                    {
                        if ((combo1.SelectedValue.ToString() == combo2.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo3.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo4.SelectedValue.ToString())
                            || (combo1.SelectedValue.ToString() == combo5.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo3.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo4.SelectedValue.ToString())
                            || (combo2.SelectedValue.ToString() == combo5.SelectedValue.ToString())
                            || (combo3.SelectedValue.ToString() == combo4.SelectedValue.ToString())
                            || (combo3.SelectedValue.ToString() == combo5.SelectedValue.ToString())
                            || (combo4.SelectedValue.ToString() == combo5.SelectedValue.ToString()))
                        {
                            checkinputcolor = true;
                        }
                    }

                }
            }


            if (txtFilePath.Text.Length == 0)
            {
                MessageBox.Show("Bạn chưa chọn đường dẫn!");
            }
            else if (startdatetime.Length == 0)
            {
                MessageBox.Show("Ngày bắt đầu không hợp lệ");
            }
            else if (enddatetime.Length == 0)
            {
                MessageBox.Show("Ngày kết thúc không hợp lệ");
            }
            else if (!group2 && !group3 && !group4 && !group5)
            {
                MessageBox.Show("Bạn chưa chọn nhóm màu");
            }
            else if (n > thietlaphesoModel.colcount)
            {
                MessageBox.Show("File có ít màu hơn nhóm màu bạn chọn tìm kiếm");
            }
            else if (blankcolor)
            {
                MessageBox.Show("Bạn chưa nhập đủ tên mã màu cần tìm kiếm");
            }
            else if (checkinputcolor == true)
            {
                MessageBox.Show("Bạn vừa chọn màu đã chọn trước đó");
            }
            else
            {
                searchbutton.Text = "ĐANG ĐỌC DỮ LIỆU ...";
                MessageBox.Show("Đang đọc dữ liệu từ: " + txtpath);
                excelcontroller.readExcel(txtpath);


                //tìm lớn nhất
                if (findmax)
                {
                    FindingStatus find = new FindingStatus();
                    this.NavigationService.Navigate(find);
                    if (limit == "")
                    {
                        tlhscontroller.readLimit(0);
                    }
                    else
                    {
                        int limitvalue = Int32.Parse(limit);
                        tlhscontroller.readLimit(limitvalue);
                    }

                    if (group2)
                    {
                        await Task.Run(new Action(tlhscontroller.processGroup));
                        await Task.Run(() => outcontroller.sortOutPut(2));
                        await Task.Run(() => outcontroller.countColorNumberMaxValue(2));
                        checkstop = true;
                    }
                    else if (group3)
                    {
                        await Task.Run(new Action(tlhscontroller.processGroup));
                        await Task.Run(() => outcontroller.sortOutPut(3));
                        await Task.Run(() => outcontroller.countColorNumberMaxValue(3));
                        checkstop = true;
                    }
                    else if (group4)
                    {
                        await Task.Run(new Action(tlhscontroller.processGroup));
                        await Task.Run(() => outcontroller.sortOutPut(4));
                        await Task.Run(() => outcontroller.countColorNumberMaxValue(4));
                        checkstop = true;
                    }
                    else if (group5)
                    {
                        await Task.Run(new Action(tlhscontroller.processGroup));
                        await Task.Run(() => outcontroller.sortOutPut(5));
                        await Task.Run(() => outcontroller.countColorNumberMaxValue(5));
                        checkstop = true;
                    }
                }
                else
                {
                    if (limit == "")
                    {
                        tlhscontroller.readLimit(0);
                    }
                    else
                    {
                        int limitvalue = Int32.Parse(limit);
                        tlhscontroller.readLimit(limitvalue);
                    }

                    if (group2) //in tất cả nhóm 2
                    {
                        if (ncolor == 2)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll2(ncolor, color1, color2));
                            checkstop = true;
                        }
                        else if (ncolor == 1)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll2(ncolor, color1, ""));
                            checkstop = true;
                        }
                        else if (ncolor == 0)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            await Task.Run(() => tlhscontroller.processGroupAll2(ncolor, "", ""));
                            checkstop = true;
                        }
                    }
                    else if (group3) // in tất cả nhóm 3
                    {
                        int timestart = Environment.TickCount;
                        if (ncolor == 3)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll3(ncolor, color1, color2, color3));
                            checkstop = true;
                        }

                        else if (ncolor == 2)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll3(ncolor, color1, color2, ""));
                            checkstop = true;
                        }

                        else if (ncolor == 1)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll3(ncolor, color1, "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 0)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            await Task.Run(() => tlhscontroller.processGroupAll3(ncolor, "", "", ""));
                            checkstop = true;
                        }
                    }
                    else if (group4) // in tất cả nhóm 4 màu
                    {
                        int timestart = Environment.TickCount;
                        if (ncolor == 4)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            color4 = combo4.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll4(ncolor, color1, color2, color3, color4));
                            checkstop = true;
                        }
                        else if (ncolor == 3)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll4(ncolor, color1, color2, color3, ""));
                            checkstop = true;
                        }
                        else if (ncolor == 2)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll4(ncolor, color1, color2, "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 1)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll4(ncolor, color1, "", "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 0)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            await Task.Run(() => tlhscontroller.processGroupAll4(ncolor, "", "", "", ""));
                            checkstop = true;
                        }
                    }
                    else if (group5) // in tất cả nhóm 5 màu
                    {
                        int timestart = Environment.TickCount;
                        if (ncolor == 5)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            color4 = combo4.SelectedValue.ToString();
                            color5 = combo5.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, color1, color2, color3, color4, color5));
                            checkstop = true;
                        }
                        else if (ncolor == 4)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            color4 = combo4.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, color1, color2, color3, color4, ""));
                            checkstop = true;
                        }
                        else if (ncolor == 3)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            color3 = combo3.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, color1, color2, color3, "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 2)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            color2 = combo2.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, color1, color2, "", "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 1)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            color1 = combo1.SelectedValue.ToString();
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, color1, "", "", "", ""));
                            checkstop = true;
                        }
                        else if (ncolor == 0)
                        {
                            FindingStatus find = new FindingStatus();
                            this.NavigationService.Navigate(find);
                            await Task.Run(() => tlhscontroller.processGroupAll5(ncolor, "", "", "", "", ""));
                            checkstop = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Lỗi");
                    }
                }
            }
        }

        // Radio button số lượng mã màu

        private void RadioButton2_Checked(object sender, RoutedEventArgs e)
        {
            if (!findmax)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            else
            {
                textcolornumber.Visibility = Visibility.Hidden;
                colornumber0.Visibility = Visibility.Hidden;
                colornumber1.Visibility = Visibility.Hidden;
                colornumber2.Visibility = Visibility.Hidden;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            n = 2;
            tlhscontroller.readN(2);
            group2 = true;
        }

        private void RadioButton2_Unchecked(object sender, RoutedEventArgs e)
        {
            group2 = false;
        }

        private void RadioButton3_Checked(object sender, RoutedEventArgs e)
        {
            if (!findmax)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            else
            {
                textcolornumber.Visibility = Visibility.Hidden;
                colornumber0.Visibility = Visibility.Hidden;
                colornumber1.Visibility = Visibility.Hidden;
                colornumber2.Visibility = Visibility.Hidden;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            n = 3;
            tlhscontroller.readN(3);
            group3 = true;
        }

        private void RadioButton3_Unchecked(object sender, RoutedEventArgs e)
        {
            group3 = false;
        }

        private void RadioButton4_Checked(object sender, RoutedEventArgs e)
        {
            if (!findmax)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
                colornumber4.Visibility = Visibility.Visible;
                colornumber5.Visibility = Visibility.Hidden;
            }
            else
            {
                textcolornumber.Visibility = Visibility.Hidden;
                colornumber0.Visibility = Visibility.Hidden;
                colornumber1.Visibility = Visibility.Hidden;
                colornumber2.Visibility = Visibility.Hidden;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            n = 4;
            tlhscontroller.readN(4);
            group4 = true;
        }



        private void RadioButton4_Unchecked(object sender, RoutedEventArgs e)
        {
            group4 = false;
        }

        private void RadioButton5_Checked(object sender, RoutedEventArgs e)
        {
            if (!findmax)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
                colornumber4.Visibility = Visibility.Visible;
                colornumber5.Visibility = Visibility.Visible;
            }
            else
            {
                textcolornumber.Visibility = Visibility.Hidden;
                colornumber0.Visibility = Visibility.Hidden;
                colornumber1.Visibility = Visibility.Hidden;
                colornumber2.Visibility = Visibility.Hidden;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
            n = 5;
            tlhscontroller.readN(5);
            group5 = true;
        }

        private void RadioButton5_Unchecked(object sender, RoutedEventArgs e)
        {
            group5 = false;
        }

        private void inputvalue_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            {
                if (!char.IsDigit(e.Text, e.Text.Length - 1))
                {
                    e.Handled = true;
                    MessageBox.Show("I only accept numbers, sorry. :(", "This textbox says...");
                }
            }
        }

        private void txtFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {

        }


        //radiobutton tìm max hoặc tất cả
        private void RadioButtonTop_Checked(object sender, RoutedEventArgs e)
        {
            findmax = true;
            if (group2 || group3 || group4 || group5)
            {
                textcolornumber.Visibility = Visibility.Hidden;
                colornumber0.Visibility = Visibility.Hidden;
                colornumber1.Visibility = Visibility.Hidden;
                colornumber2.Visibility = Visibility.Hidden;
                colornumber3.Visibility = Visibility.Hidden;
                colornumber4.Visibility = Visibility.Hidden;
                colornumber5.Visibility = Visibility.Hidden;
            }
        }

        private void RadioButtonTop_Unchecked(object sender, RoutedEventArgs e)
        {
            findmax = false;
        }

        private void inputvalue_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void RadioButtonAll_Checked(object sender, RoutedEventArgs e)
        {
            findmax = false;
            if (group2)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
            }
            else if (group3)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
            }
            else if (group4)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
                colornumber4.Visibility = Visibility.Visible;
            }
            else if (group5)
            {
                textcolornumber.Visibility = Visibility.Visible;
                colornumber0.Visibility = Visibility.Visible;
                colornumber0.IsChecked = true;
                colornumber1.Visibility = Visibility.Visible;
                colornumber2.Visibility = Visibility.Visible;
                colornumber3.Visibility = Visibility.Visible;
                colornumber4.Visibility = Visibility.Visible;
                colornumber5.Visibility = Visibility.Visible;
            }
        }

        private void RadioButtonAll_Unchecked(object sender, RoutedEventArgs e)
        {
            findmax = true;
        }

        // Radiobutton số lượng mã màu người dùng nhập
        private void ColorButton0_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 0;
            colorgroup0 = true;
            combo1.Visibility = Visibility.Hidden;
            combo2.Visibility = Visibility.Hidden;
            combo3.Visibility = Visibility.Hidden;
            combo4.Visibility = Visibility.Hidden;
            combo5.Visibility = Visibility.Hidden;
        }

        private void ColorButton0_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup0 = false;
        }

        private void ColorButton1_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 1;
            colorgroup1 = true;
            combo1.Visibility = Visibility.Visible;
            combo2.Visibility = Visibility.Hidden;
            combo3.Visibility = Visibility.Hidden;
            combo4.Visibility = Visibility.Hidden;
            combo5.Visibility = Visibility.Hidden;
        }

        private void ColorButton1_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup1 = false;
        }

        private void ColorButton2_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 2;
            colorgroup2 = true;
            combo1.Visibility = Visibility.Visible;
            combo2.Visibility = Visibility.Visible;
            combo3.Visibility = Visibility.Hidden;
            combo4.Visibility = Visibility.Hidden;
            combo5.Visibility = Visibility.Hidden;
        }

        private void ColorButton2_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup2 = false;
        }

        private void ColorButton3_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 3;
            colorgroup3 = true;
            combo1.Visibility = Visibility.Visible;
            combo2.Visibility = Visibility.Visible;
            combo3.Visibility = Visibility.Visible;
            combo4.Visibility = Visibility.Hidden;
            combo5.Visibility = Visibility.Hidden;
        }

        private void ColorButton3_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup3 = false;
        }

        private void ColorButton4_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 4;
            colorgroup4 = true;
            combo1.Visibility = Visibility.Visible;
            combo2.Visibility = Visibility.Visible;
            combo3.Visibility = Visibility.Visible;
            combo4.Visibility = Visibility.Visible;
            combo5.Visibility = Visibility.Hidden;
        }

        private void ColorButton4_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup4 = false;
        }

        private void ColorButton5_Checked(object sender, RoutedEventArgs e)
        {
            ncolor = 5;
            colorgroup5 = true;
            combo1.Visibility = Visibility.Visible;
            combo2.Visibility = Visibility.Visible;
            combo3.Visibility = Visibility.Visible;
            combo4.Visibility = Visibility.Visible;
            combo5.Visibility = Visibility.Visible;
        }

        private void ColorButton5_Unchecked(object sender, RoutedEventArgs e)
        {
            colorgroup5 = false;
        }

    }
}