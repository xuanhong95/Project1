using System;
using System.Windows.Controls;
using System.Windows.Forms;
using DataProcessing.Controller;
using System.Numerics;
using DataProcessing.Model;

namespace DataProcessing
{
    /// <summary>
    /// Interaction logic for FindingStatus.xaml
    /// </summary>
    public partial class FindingStatus : Page
    {
        MiddlewareController middle = new MiddlewareController();
        private Timer timer1 = new Timer();
        int gio = 0, phut = 0, giay = 0;
        int speed1 = 0, speed2 = 0;
        BigInteger sp = 1;
        BigInteger totalcolor = 0;
        BigInteger totalcolor_old = (BigInteger)Decimal.MaxValue;
        Decimal foundcolor, foundcolor_old, number1, convertBigInt;

        private void continueButton(object sender, System.Windows.RoutedEventArgs e)
        {
            thietlaphesoModel.colcount = 1;
            thietlaphesoModel.rowcount = 1;
            Array.Clear(thietlaphesoModel.color, 0, thietlaphesoModel.color.Length);
            Array.Clear(thietlaphesoModel.colordefault, 0, thietlaphesoModel.colordefault.Length);
            Array.Clear(thietlaphesoModel.datetime, 0, thietlaphesoModel.datetime.Length);
            Array.Clear(thietlaphesoModel.value, 0, thietlaphesoModel.value.Length);
            Array.Clear(thietlaphesoModel.zeroOne, 0, thietlaphesoModel.zeroOne.Length);
            Array.Clear(thietlaphesoModel.index, 0, thietlaphesoModel.index.Length);
            thietlaphesoModel.n = 0;
            thietlaphesoModel.limitvalue = 0;
            MiddlewareModel.foundedColor = 0;
            MiddlewareModel.foundedColor_MaxValue = 0;
            thietlapHeSo.checkstop = false;
            thietlapHeSo.findmax = true;
            MiddlewareController.tu = 1;
            MiddlewareController.mau = 1;
            ExcelController.duplicateindex.Clear();
            ExcelController.mergeduplicateindex.Clear();
            ExcelController.tempo.Clear();
            ExcelController.origin_color_index.Clear();
            thietlapHeSo.n = 0;
            thietlapHeSo.ncolor = 0;

            thietlapHeSo tlhs = new thietlapHeSo();
            this.NavigationService.Navigate(tlhs);
        }

        private void exitApp(object sender, System.Windows.RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        BigInteger es_gio = 0, es_phut = 0, es_giay = 0, es_giay_old = 0;
        public FindingStatus()
        {

            InitializeComponent();
            start.Text = thietlapHeSo.startdatetime;
            end.Text = thietlapHeSo.enddatetime;
            colorgroup.Text = thietlapHeSo.n.ToString();
            totalcolor = MiddlewareController.estimateTime(middle.getExcelCol(), thietlapHeSo.n);
            number1 = (decimal)totalcolor;
            totalcolor_old = totalcolor;
            if (thietlapHeSo.limit.ToString() == "")
            {
                limitvalue.Text = "0";
            }
            else
                limitvalue.Text = thietlapHeSo.limit.ToString();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 1000; // in miliseconds
            timer1.Start();
        }





        private void timer1_Tick(object sender, EventArgs e)
        {
            if (thietlapHeSo.checkstop)
            {
                timer1.Stop();
                if (thietlapHeSo.findmax == false)
                {
                    foundedcolor.Text = middle.getFoundedColorValue().ToString();
                }
                else
                {
                    foundedcolor.Text = middle.getColorNumberMaxValue().ToString();
                }
                processSpeed.Text = "0";
                if (giay < 1)
                {
                    processtime.Text = "0h 0m 0s";
                }
                estimate.Text = "0h 0m 0s";

                pbMyProgressBar.Value = 100;
                continueGrid.Visibility = System.Windows.Visibility.Visible;
                MessageBox.Show("Tìm kiếm kết thúc");
            }
            else
            {
                if ((giay + 1) == 60)
                {
                    ++phut;
                    giay = -1;
                }
                if ((phut + 1) == 60)
                {
                    ++gio;
                    phut = -1;
                }
                processtime.Text = gio.ToString() + "h " + phut.ToString() + "m " + (++giay).ToString() + "s";


                speed1 = middle.getFoundedColorValue() - speed2;
                speed2 += speed1;


                sp = speed1;
                totalcolor -= sp;
                if(sp == 0)
                {
                    sp = 1;
                }

                if (giay == 1 && phut == 0 && gio == 0)
                {
                    es_giay = totalcolor / sp;
                    es_giay_old = es_giay;
                    es_gio = es_giay / 3600;
                    es_giay %= 3600;
                    es_phut = es_giay / 60;
                    es_giay %= 60;
                    estimate.Text = es_gio.ToString() + "h " + es_phut.ToString() + "m " + es_giay.ToString() + "s";
                }
                else
                {
                    if (sp == 0)
                    {
                        sp = 1;
                    }
                    if (es_giay_old > (totalcolor / sp))
                    {
                        es_giay = totalcolor / sp;
                        es_giay_old = es_giay;
                        es_gio = es_giay / 3600;
                        es_giay %= 3600;
                        es_phut = es_giay / 60;
                        es_giay %= 60;
                        if (thietlapHeSo.findmax == false)
                        {
                            estimate.Text = es_gio.ToString() + "h " + es_phut.ToString() + "m " + es_giay.ToString() + "s";
                        }
                        else
                        {
                            estimate.Text = "Đang cập nhật ...";
                        }
                    }
                }
                
                if (thietlapHeSo.findmax == false)
                {
                    processSpeed.Text = speed1.ToString() + " nhóm màu/s";
                    foundedcolor.Text = middle.getFoundedColorValue().ToString();
                    foundcolor = (decimal)middle.getFoundedColorValue();
                    convertBigInt += (Decimal.Divide(foundcolor - foundcolor_old, number1) * 100);
                    foundcolor_old = foundcolor;
                    pbMyProgressBar.Value = (double)convertBigInt;
                }
                else
                {
                    processSpeed.Text = "Đang cập nhật ...";
                    foundedcolor.Text = "Đang cập nhật ...";
                }
            }

        }
    }
}
