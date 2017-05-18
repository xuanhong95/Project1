using System;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;
using DataProcessing.Model;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace DataProcessing.Controller
{
    class ExcelController
    {
        thietlaphesoModel model = new thietlaphesoModel();
        public static int ngaybatdau = 0, ngayketthuc = 0;
        public static int ncolor = 0;
        public static int colprogress = 0;
        public static List<int> duplicateindex = new List<int>();
        public static List<int> mergeduplicateindex = new List<int>();
        public static List<List<int>> tempo = new List<List<int>>();

        public static List<int> origin_color_index = new List<int>();
        public static int cotdautien = 0;


        public void setNColor(int numberinputcolor)
        {
            ncolor = numberinputcolor;
        }
        public string[] fillColorCombobox()
        {
            string[] array = model.getColorDefault();
            return array;
        }

        public string[] fillDateTime()
        {
            string[] array = model.getDateTime();
            return array;
        }



        public void getColorAndDate(String path)
        {
            Excel.Application excel;
            excel = new Excel.Application();
            Excel.Workbook WB = excel.Workbooks.Open(path);
            WB = excel.ActiveWorkbook;
            Excel.Worksheet WS;
            WS = WB.ActiveSheet;


            model.setColCount(WS.UsedRange.Columns.Count);
            model.setRowCount(WS.UsedRange.Rows.Count);
            String[] color = new string[model.getColCount() - 1];
            String[] colordefault = new string[model.getColCount() - 1];
            String[] datetime = new string[model.getRowCount() - 1];
            int[] index = new int[model.getColCount() - 1];
            for (int i = 0; i < model.getColCount() - 1; i++)
            {
                index[i] = i;
            }

            Dictionary<string, int> hashmap = new Dictionary<string, int>();

            Excel.Range colornumber = WS.get_Range((Excel.Range)WS.Cells[2][1], (Excel.Range)WS.Cells[model.getColCount()][1]);
            object mamau = colornumber.Value;
            //Lấy mã màu vào mảng
            int colorcount = 0;
            int t = 0;
            foreach (object objcolor in (Array)mamau)
            {

                string colorname = (string)objcolor;
                try
                {
                    hashmap.Add(colorname, t);
                    t++;
                    color[colorcount] = colorname;
                    colordefault[colorcount] = colorname;
                }
                catch (Exception e)
                {
                    View.Warning.Show("Màu " + colorname + " đã trùng");
                    X:
                    if (View.Warning.yourchoise == 1)
                    {
                        duplicateindex.Add(colorcount + duplicateindex.Count);
                        colorcount--;
                    }
                    else if (View.Warning.yourchoise == 2)
                    {
                        List<int> dcm = new List<int>();
                        int dict = hashmap[colorname];
                        origin_color_index.Add(dict);
                        cotdautien = dict;

                        dcm.Add(dict);
                        dcm.Add(colorcount + duplicateindex.Count);
                        tempo.Add(dcm);

                        mergeduplicateindex.Add(colorcount + duplicateindex.Count);
                        duplicateindex.Add(colorcount + duplicateindex.Count);

                        colorcount--;
                    }
                    else if (View.Warning.yourchoise == 3)
                    {
                            try
                            {
                                hashmap.Add(View.Warning.newnamecolor, 0);
                                color[colorcount] = View.Warning.newnamecolor;
                                colordefault[colorcount] = View.Warning.newnamecolor;
                            }
                            catch (Exception a)
                            {
                                View.Warning.Show("Màu " + View.Warning.newnamecolor + " đã tồn tại");
                                colorname = View.Warning.newnamecolor;
                                goto X;
                            }
                    }
                }
                colorcount++;
            }
            //Lấy mã ngày vào mảng
            int datecount = 0;
            for (int row = 2; row <= model.getRowCount(); row++)
            {
                string cell = (WS.Cells[1][row] as Excel.Range).Value == null ? "" : (WS.Cells[1][row] as Excel.Range).Value.ToString("M/dd/yyyy");
                datetime[datecount] = cell;
                datecount++;
            }
            model.setColor(color);
            model.setColorDefault(colordefault);
            model.setDateTime(datetime);
            model.setIndex(index);
            excel.Quit();

        }


        public void readExcel(String path)
        {

            Excel.Application excel;
            excel = new Excel.Application();
            Excel.Workbook WB = excel.Workbooks.Open(path);
            WB = excel.ActiveWorkbook;
            Excel.Worksheet WS;
            WS = WB.ActiveSheet;

            int start = Environment.TickCount;

            model.setColCount(WS.UsedRange.Columns.Count);
            model.setRowCount(WS.UsedRange.Rows.Count);
            String[] color = model.getColor();
            String[] colordefault = new string[model.getColCount() - 1];
            string[] datetime = model.getDateTime();
            int[] value = new int[model.getColCount() - 1];
            int[] index = model.getIndex();
            int[][] zeroOne = new int[model.getColCount() - 1][];
            for (int i = 0; i < model.getDateTime().Length; i++)
            //Khoanh vùng ngày bắt đầu và kết thúc
            {
                if (datetime[i] == thietlapHeSo.startdatetime)
                {
                    ngaybatdau = i + 2;
                }
                else if (datetime[i] == thietlapHeSo.enddatetime)
                {
                    ngayketthuc = i + 2;
                }
            }

            //Tạo mảng 2 chiều zeroOne
            for (int i = 0; i < model.getColCount() - 1; i++)
            {
                zeroOne[i] = new int[ngayketthuc - ngaybatdau + 1];
            }

            int check_change_first_time = 0; //Kiểm tra phát hiện đầu trùng trong vòng lặp
            int check_count_duplicate = 0; //Biến đếm số lượng màu trùng đã tìm để dịch mảng phía sau màu trùng đó
            int merge = 0;
            int delete = 0;
            int checktempo = 0;
            //Tính tổng tất cả các cột theo thời gian đã định
            for (int i = 2; i <= model.getColCount(); i++)
            {

                int temp = 0;
                Excel.Range b = WS.get_Range((Excel.Range)WS.Cells[i][ngaybatdau], (Excel.Range)WS.Cells[i][ngayketthuc]);
                object arr = b.Value;
                int j = 0;

                //Kiểm tra mã màu đã bị xóa hoặc gộp chưa
                foreach (int check_duplicate_index in duplicateindex)
                {
                    if (index[i - 2] == check_duplicate_index)
                    {
                        check_change_first_time += 1;
                        check_count_duplicate++;
                        delete++;
                        foreach (int check_merge in mergeduplicateindex)
                        {
                            if (index[i - 2] == check_merge)
                            {
                                merge++;
                                break;
                            }
                        }
                        break;
                    }
                }

                /// <!-- Kiểm tra nếu chưa có màu trùng -->
                if (check_count_duplicate == 0)
                {
                    foreach (object s in (Array)arr)
                    {

                        string tmp = s == null ? "" : "1";
                        if (tmp != "1")
                        {
                            zeroOne[i - 2][j] = 0;
                            j++;
                            continue;
                        }

                        else
                        {
                            zeroOne[i - 2][j] = 1;
                            j++;
                            temp += int.Parse(tmp);

                        }

                    }
                    value[i - 2] = temp;
                }
                //Nếu đã có màu trùng
                else if (check_count_duplicate != 0)
                {
                    //Nếu phát hiện màu trùng trong vòng lặp hiện tại
                    if (check_change_first_time == 1)
                    {
                        //Phát hiện màu trùng xử lý theo kiểu gộp
                        if (merge == 1)
                        {
                            check_change_first_time = 0;
                            merge = 0;

                            List<int> aaa = tempo[checktempo];
                            j = 0;
                            Console.WriteLine(aaa.Count);

                            int dodaingay = 0;
                            foreach (object s in (Array)arr)
                            {
                                string tmp = s == null ? "" : "1";
                                if (tmp != "1")
                                {
                                    zeroOne[aaa[1]][j] = 0;
                                    j++;
                                }
                                else
                                {
                                    zeroOne[aaa[1]][j] = 1;
                                    j++;
                                    temp += int.Parse(tmp);
                                }
                                dodaingay++;
                            }

                            for (int m = 0; m < dodaingay; m++)
                            {
                                if (zeroOne[aaa[1]][m] == 1)
                                {
                                    zeroOne[aaa[0]][m] = 1;
                                }
                            }
                            checktempo++;
                        }
                        //xử lý kiểu xóa
                        else if (delete == 1)
                        {
                            delete = 0;
                            check_change_first_time = 0;
                            continue;
                        }
                        else
                        {
                            check_change_first_time = 0;
                            foreach (object s in (Array)arr)
                            {

                                string tmp = s == null ? "" : "1";
                                if (tmp != "1")
                                {
                                    zeroOne[i - 2 - check_count_duplicate][j] = 0;
                                    j++;
                                }
                                else
                                {
                                    zeroOne[i - 2 - check_count_duplicate][j] = 1;
                                    j++;
                                    temp += int.Parse(tmp);
                                }
                            }
                            value[i - 2 - check_count_duplicate] = temp;
                        }
                    }
                    //Dịch mảng
                    else
                    {
                        foreach (object s in (Array)arr)
                        {

                            string tmp = s == null ? "" : "1";
                            if (tmp != "1")
                            {
                                zeroOne[i - 2 - check_count_duplicate][j] = 0;
                                j++;
                            }
                            else
                            {
                                zeroOne[i - 2 - check_count_duplicate][j] = 1;
                                j++;
                                temp += int.Parse(tmp);
                            }
                        }
                        value[i - 2 - check_count_duplicate] = temp;
                    }
                }
            }
            model.setValue(value);
            model.setZeroOne(zeroOne);
            excel.Quit();
        }

        /// <summary>
        /// Hàm sắp xếp mảng theo mã màu người dùng nhập vào
        /// </summary>
        public void readExcelSortByColor(int ncolor, string[] color, int[] value, int[] index, int[][] zeroOne, string color1, string color2, string color3, string color4, string color5)
        {
            string tmp4 = "";
            int tmp5, tmp6, tmpindex0;
            if (ncolor == 1) //Người dùng nhập sẵn 1 mã màu
            {
                for (int i = 0; i < model.getColCount() - 1 - duplicateindex.Count; i++)
                {
                    if (color[i] == color1)
                    {
                        tmp4 = color[i];
                        tmp5 = value[i];
                        tmpindex0 = index[i];
                        color[i] = color[0];
                        value[i] = value[0];
                        index[i] = index[0];
                        color[0] = tmp4;
                        value[0] = tmp5;
                        index[0] = tmpindex0;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp6 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[0][n];
                            zeroOne[0][n] = tmp6;
                        }
                        break;
                    }
                }
                int tmp1, tmp2, tmpindex;
                string tmp3 = "";
                for (int i = 1; i < value.Length - duplicateindex.Count; i++)
                {
                    for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp1 = value[i];
                            tmpindex = index[i];
                            tmp3 = color[i];
                            value[i] = value[j];
                            index[i] = index[j];
                            color[i] = color[j];
                            value[j] = tmp1;
                            index[j] = tmpindex;
                            color[j] = tmp3;
                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                            {
                                tmp2 = zeroOne[i][n];
                                zeroOne[i][n] = zeroOne[j][n];
                                zeroOne[j][n] = tmp2;
                            }
                        }
                    }
                }
            }
            else if (ncolor == 2) //Người dùng nhập sẵn 2 mã màu
            {
                for (int i = 0; i < model.getColCount() - 1 - duplicateindex.Count; i++)
                {
                    int checkcolor = 0;
                    if ((color[i] == color1 || color[i] == color2))
                    {
                        tmp4 = color[i];
                        tmp5 = value[i];
                        tmpindex0 = index[i];
                        color[i] = color[0];
                        value[i] = value[0];
                        index[i] = index[0];
                        color[0] = tmp4;
                        value[0] = tmp5;
                        index[0] = tmpindex0;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp6 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[0][n];
                            zeroOne[0][n] = tmp6;
                        }
                        checkcolor++;
                        if (checkcolor == 1)
                        {
                            for (int j = i + 1; j < model.getColCount() - 1 - duplicateindex.Count; j++)
                            {
                                if ((color[j] == color1 || color[j] == color2))
                                {
                                    tmp4 = color[j];
                                    tmp5 = value[j];
                                    tmpindex0 = index[j];
                                    color[j] = color[1];
                                    value[j] = value[1];
                                    index[j] = index[1];
                                    color[1] = tmp4;
                                    value[1] = tmp5;
                                    index[1] = tmpindex0;
                                    for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                    {
                                        tmp6 = zeroOne[j][n];
                                        zeroOne[j][n] = zeroOne[1][n];
                                        zeroOne[1][n] = tmp6;
                                    }
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                int tmp1, tmp2, tmpindex;
                string tmp3 = "";
                for (int i = 2; i < value.Length; i++)
                {
                    for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp1 = value[i];
                            tmpindex = index[i];
                            tmp3 = color[i];
                            value[i] = value[j];
                            index[i] = index[j];
                            color[i] = color[j];
                            value[j] = tmp1;
                            index[j] = tmpindex;
                            color[j] = tmp3;
                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                            {
                                tmp2 = zeroOne[i][n];
                                zeroOne[i][n] = zeroOne[j][n];
                                zeroOne[j][n] = tmp2;
                            }
                        }
                    }
                }
            }
            else if (ncolor == 3)
            {
                for (int i = 0; i < model.getColCount() - 1 - duplicateindex.Count; i++)
                {
                    int checkcolor = 0;
                    if ((color[i] == color1 || color[i] == color2 || color[i] == color3))
                    {
                        tmp4 = color[i];
                        tmp5 = value[i];
                        tmpindex0 = index[i];
                        color[i] = color[0];
                        value[i] = value[0];
                        index[i] = index[0];
                        color[0] = tmp4;
                        value[0] = tmp5;
                        index[0] = tmpindex0;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp6 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[0][n];
                            zeroOne[0][n] = tmp6;
                        }
                        checkcolor++;
                        if (checkcolor == 1)
                        {
                            for (int j = i + 1; j < model.getColCount() - 1 - duplicateindex.Count; j++)
                            {
                                if ((color[j] == color1 || color[j] == color2 || color[j] == color3))
                                {
                                    tmp4 = color[j];
                                    tmp5 = value[j];
                                    tmpindex0 = index[j];
                                    color[j] = color[1];
                                    value[j] = value[1];
                                    index[j] = index[1];
                                    color[1] = tmp4;
                                    value[1] = tmp5;
                                    index[1] = tmpindex0;
                                    for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                    {
                                        tmp6 = zeroOne[j][n];
                                        zeroOne[j][n] = zeroOne[1][n];
                                        zeroOne[1][n] = tmp6;
                                    }
                                    checkcolor++;
                                    if (checkcolor == 2)
                                    {
                                        for (int k = j + 1; k < model.getColCount() - 1 - duplicateindex.Count; k++)
                                        {
                                            if ((color[k] == color1 || color[k] == color2 || color[k] == color3))
                                            {
                                                tmp4 = color[k];
                                                tmp5 = value[k];
                                                tmpindex0 = index[k];
                                                color[k] = color[2];
                                                value[k] = value[2];
                                                index[k] = index[2];
                                                color[2] = tmp4;
                                                value[2] = tmp5;
                                                index[2] = tmpindex0;
                                                for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                {
                                                    tmp6 = zeroOne[k][n];
                                                    zeroOne[k][n] = zeroOne[2][n];
                                                    zeroOne[2][n] = tmp6;
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }

                int tmp1, tmp2, tmpindex;
                string tmp3 = "";
                for (int i = 3; i < value.Length - duplicateindex.Count; i++)
                {
                    for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp1 = value[i];
                            tmpindex = index[i];
                            tmp3 = color[i];
                            value[i] = value[j];
                            index[i] = index[j];
                            color[i] = color[j];
                            value[j] = tmp1;
                            index[j] = tmpindex;
                            color[j] = tmp3;
                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                            {
                                tmp2 = zeroOne[i][n];
                                zeroOne[i][n] = zeroOne[j][n];
                                zeroOne[j][n] = tmp2;
                            }
                        }
                    }
                }
            }
            else if (ncolor == 4)
            {
                for (int i = 0; i < model.getColCount() - 1 - duplicateindex.Count; i++)
                {
                    int checkcolor = 0;
                    if (color[i] == color1 || color[i] == color2 || color[i] == color3 || color[i] == color4)
                    {
                        tmp4 = color[i];
                        tmp5 = value[i];
                        tmpindex0 = index[i];
                        color[i] = color[0];
                        value[i] = value[0];
                        index[i] = index[0];
                        color[0] = tmp4;
                        value[0] = tmp5;
                        index[0] = tmpindex0;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp6 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[0][n];
                            zeroOne[0][n] = tmp6;
                        }
                        checkcolor++;
                        if (checkcolor == 1)
                        {
                            for (int j = i + 1; j < model.getColCount() - 1 - duplicateindex.Count; j++)
                            {
                                if (color[j] == color1 || color[j] == color2 || color[j] == color3 || color[j] == color4)
                                {
                                    tmp4 = color[j];
                                    tmp5 = value[j];
                                    tmpindex0 = index[j];
                                    color[j] = color[1];
                                    value[j] = value[1];
                                    index[j] = index[1];
                                    color[1] = tmp4;
                                    value[1] = tmp5;
                                    index[1] = tmpindex0;
                                    for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                    {
                                        tmp6 = zeroOne[j][n];
                                        zeroOne[j][n] = zeroOne[1][n];
                                        zeroOne[1][n] = tmp6;
                                    }
                                    checkcolor++;
                                    if (checkcolor == 2)
                                    {
                                        for (int k = j + 1; k < model.getColCount() - 1 - duplicateindex.Count; k++)
                                        {
                                            if (color[k] == color1 || color[k] == color2 || color[k] == color3 || color[k] == color4)
                                            {
                                                tmp4 = color[k];
                                                tmp5 = value[k];
                                                tmpindex0 = index[k];
                                                color[k] = color[2];
                                                value[k] = value[2];
                                                index[k] = index[2];
                                                color[2] = tmp4;
                                                value[2] = tmp5;
                                                index[2] = tmpindex0;
                                                for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                {
                                                    tmp6 = zeroOne[k][n];
                                                    zeroOne[k][n] = zeroOne[2][n];
                                                    zeroOne[2][n] = tmp6;
                                                }
                                                checkcolor++;
                                                if (checkcolor == 3)
                                                {
                                                    for (int l = k + 1; l < model.getColCount() - 1 - duplicateindex.Count; l++)
                                                    {
                                                        if (color[l] == color1 || color[l] == color2 || color[l] == color3 || color[l] == color4)
                                                        {
                                                            tmp4 = color[l];
                                                            tmp5 = value[l];
                                                            tmpindex0 = index[l];
                                                            color[l] = color[3];
                                                            value[l] = value[3];
                                                            index[l] = index[3];
                                                            color[3] = tmp4;
                                                            value[3] = tmp5;
                                                            index[3] = tmpindex0;
                                                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                            {
                                                                tmp6 = zeroOne[l][n];
                                                                zeroOne[l][n] = zeroOne[3][n];
                                                                zeroOne[3][n] = tmp6;
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }

                int tmp1, tmp2, tmpindex;
                string tmp3 = "";
                for (int i = 4; i < value.Length - duplicateindex.Count; i++)
                {
                    for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp1 = value[i];
                            tmpindex = index[i];
                            tmp3 = color[i];
                            value[i] = value[j];
                            index[i] = index[j];
                            color[i] = color[j];
                            value[j] = tmp1;
                            index[j] = tmpindex;
                            color[j] = tmp3;
                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                            {
                                tmp2 = zeroOne[i][n];
                                zeroOne[i][n] = zeroOne[j][n];
                                zeroOne[j][n] = tmp2;
                            }
                        }
                    }
                }
            }
            else if (ncolor == 5)
            {
                //sắp xếp mảng
                for (int i = 0; i < model.getColCount() - 1 - duplicateindex.Count; i++)
                {
                    int checkcolor = 0;
                    if (color[i] == color1 || color[i] == color2 || color[i] == color3 || color[i] == color4 || color[i] == color5)
                    {
                        tmp4 = color[i];
                        tmp5 = value[i];
                        tmpindex0 = index[i];
                        color[i] = color[0];
                        value[i] = value[0];
                        index[i] = index[0];
                        color[0] = tmp4;
                        value[0] = tmp5;
                        index[0] = tmpindex0;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp6 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[0][n];
                            zeroOne[0][n] = tmp6;
                        }
                        MessageBox.Show(color[0]);
                        checkcolor++;
                        if (checkcolor == 1)
                        {
                            for (int j = i + 1; j < model.getColCount() - 1 - duplicateindex.Count; j++)
                            {
                                if (color[j] == color1 || color[j] == color2 || color[j] == color3 || color[j] == color4 || color[j] == color5)
                                {
                                    tmp4 = color[j];
                                    tmp5 = value[j];
                                    tmpindex0 = index[j];
                                    color[j] = color[1];
                                    value[j] = value[1];
                                    index[j] = index[1];
                                    color[1] = tmp4;
                                    value[1] = tmp5;
                                    index[1] = tmpindex0;
                                    for (int n = 0; n < ngayketthuc - ngaybatdau + 1 - duplicateindex.Count; n++)
                                    {
                                        tmp6 = zeroOne[j][n];
                                        zeroOne[j][n] = zeroOne[1][n];
                                        zeroOne[1][n] = tmp6;
                                    }
                                    checkcolor++;
                                    if (checkcolor == 2)
                                    {
                                        for (int k = j + 1; k < model.getColCount() - 1 - duplicateindex.Count; k++)
                                        {
                                            if (color[k] == color1 || color[k] == color2 || color[k] == color3 || color[k] == color4 || color[k] == color5)
                                            {
                                                tmp4 = color[k];
                                                tmp5 = value[k];
                                                tmpindex0 = index[k];
                                                color[k] = color[2];
                                                value[k] = value[2];
                                                index[k] = index[2];
                                                color[2] = tmp4;
                                                value[2] = tmp5;
                                                index[2] = tmpindex0;
                                                for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                {
                                                    tmp6 = zeroOne[k][n];
                                                    zeroOne[k][n] = zeroOne[2][n];
                                                    zeroOne[2][n] = tmp6;
                                                }
                                                checkcolor++;
                                                if (checkcolor == 3)
                                                {
                                                    for (int l = k + 1; l < model.getColCount() - 1 - duplicateindex.Count; l++)
                                                    {
                                                        if (color[l] == color1 || color[l] == color2 || color[l] == color3 || color[l] == color4 || color[l] == color5)
                                                        {
                                                            tmp4 = color[l];
                                                            tmp5 = value[l];
                                                            tmpindex0 = index[l];
                                                            color[l] = color[3];
                                                            value[l] = value[3];
                                                            index[l] = index[3];
                                                            color[3] = tmp4;
                                                            value[3] = tmp5;
                                                            index[3] = tmpindex0;
                                                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                            {
                                                                tmp6 = zeroOne[l][n];
                                                                zeroOne[l][n] = zeroOne[3][n];
                                                                zeroOne[3][n] = tmp6;
                                                            }
                                                            checkcolor++;
                                                            if (checkcolor == 4)
                                                            {
                                                                for (int q = l + 1; q < model.getColCount() - 1 - duplicateindex.Count; q++)
                                                                {
                                                                    if (color[q] == color1 || color[q] == color2 || color[q] == color3 || color[q] == color4 || color[q] == color5)
                                                                    {
                                                                        tmp4 = color[q];
                                                                        tmp5 = value[q];
                                                                        tmpindex0 = index[q];
                                                                        color[q] = color[4];
                                                                        value[q] = value[4];
                                                                        index[q] = index[4];
                                                                        color[4] = tmp4;
                                                                        value[4] = tmp5;
                                                                        index[4] = tmpindex0;
                                                                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                                                                        {
                                                                            tmp6 = zeroOne[q][n];
                                                                            zeroOne[q][n] = zeroOne[4][n];
                                                                            zeroOne[4][n] = tmp6;
                                                                        }
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }

                int tmp1, tmp2, tmpindex;
                string tmp3 = "";
                for (int i = 5; i < value.Length - duplicateindex.Count; i++)
                {
                    for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp1 = value[i];
                            tmpindex = index[i];
                            tmp3 = color[i];
                            value[i] = value[j];
                            index[i] = index[j];
                            color[i] = color[j];
                            value[j] = tmp1;
                            index[j] = tmpindex;
                            color[j] = tmp3;
                            for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                            {
                                tmp2 = zeroOne[i][n];
                                zeroOne[i][n] = zeroOne[j][n];
                                zeroOne[j][n] = tmp2;
                            }
                        }
                    }
                }
            }
            model.setColor(color);
            model.setIndex(index);
            model.setValue(value);
            model.setZeroOne(zeroOne);
        }

        /// <summary>
        /// Hàm sắp xếp mảng theo tổng số ngày bán được của mã màu
        /// </summary>
        public void readExcelSortByValue(string[] color, int[] value, int[] index, int[][] zeroOne)
        {
            //sắp xếp mảng
            int tmp1, tmp2, tmpindex;
            string tmp3 = "";
            for (int i = 0; i < value.Length - duplicateindex.Count; i++)
            {
                for (int j = i + 1; j < value.Length - duplicateindex.Count; j++)
                {
                    if (value[i] < value[j])
                    {
                        tmp1 = value[i];
                        tmpindex = index[i];
                        tmp3 = color[i];
                        value[i] = value[j];
                        index[i] = index[j];
                        color[i] = color[j];
                        value[j] = tmp1;
                        index[j] = tmpindex;
                        color[j] = tmp3;
                        for (int n = 0; n < ngayketthuc - ngaybatdau + 1; n++)
                        {
                            tmp2 = zeroOne[i][n];
                            zeroOne[i][n] = zeroOne[j][n];
                            zeroOne[j][n] = tmp2;
                        }
                    }
                }
            }
            model.setIndex(index);
            model.setColor(color);
            model.setValue(value);
            model.setZeroOne(zeroOne);
        }
    }
}
