using System;
using System.Collections.Generic;
using System.Linq;
using DataProcessing.Model;

namespace DataProcessing.Controller
{
    public class AlgorithmController
    {
        /// <summary>
        /// Hàm xử lý nhóm 2 màu
        /// </summary>
        /// 

        bool canStop = true;
        string printOut = "";
        static int[] max;
        int limitedInputValue = 0; // nguong gioi han dau vao
        thietlaphesoModel model = new thietlaphesoModel();
        MiddlewareController middle = new MiddlewareController();
        ExcelController exc = new ExcelController();
        public void readN(int n)
        {
            model.setN(n);
        }

        public void readLimit(int limit)
        {
            model.setLimit(limit);
        }
        // tìm nhóm lớn nhất theo yêu cầu gviên
        public void processGroup()
        {
            exc.readExcelSortByValue(model.getColor(), model.getValue(), model.getIndex(), model.getZeroOne());
            int n = model.getN();
            string print = "";
            int limitedInputValue = model.getLimit();
            int currentValue1; // giá trị ở vòng 1
            int currentValue2; // giá trị ở vòng 2
            int currentValue3; // giá trị ở vòng 3
            int currentValue4; // giá trị ở vòng 4
            int currentValue5;
            int biggestValue = 0;
            int[] value = model.getValue();
            int[][] zeroOne = model.getZeroOne();
            int[] index = model.getIndex();
            int[] max = new int[model.getColCount() - 1];
            string[] color = model.getColor();
            List<int> savedRound2 = new List<int>(); // lưu các màu mang giá trị lớn nhất ở mốc màu thứ 2
            List<int> savedRound3 = new List<int>();
            List<int> savedRound4 = new List<int>();

            if (value[n - 1] == 0)
            {
                canStop = false;
            }

            for (int i = 0; i < model.getColCount() - n; i++)
            {
                // điều kiện dừng
                if (checkToBreak(n, biggestValue, value[i]) || (value[i + n - 1] == 0 && canStop))
                {
                    break;
                }
                print = "";
                biggestValue = 0;
                List<int> checkList1 = new List<int>(); // list so sánh theo ngày không bán được sau vòng 1
                currentValue1 = value[i];
                for (int j = 0; j < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; j++) // tạo list chứa những ngày không bán được của màu đầu tiên
                {
                    if (zeroOne[i][j] == 0) // tìm ngày không bán được để add vào list
                    {
                        checkList1.Add(j);
                    }
                }

                if (!checkList1.Any()) // màu đầu tiên full 1
                {
                    biggestValue = currentValue1;
                    if (biggestValue < limitedInputValue)
                    {
                        break;
                    }
                    for (int j = i + 1; j < model.getColCount() - n + 1; j++)
                    {
                        if (value[j] == 0 && (canStop || (value[n - 1] == 0 && value[1] != 0)))
                        {
                            break;
                        }

                        if (n == 2)
                        {
                            print += color[i] + " " + color[j] + ": " + biggestValue + Environment.NewLine;
                            continue;
                        }
                        else // n > 2
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2; q++)
                            {
                                if (value[q] == 0 && (canStop || (value[n - 1] == 0 && value[2] != 0)))
                                {
                                    break;
                                }

                                if (n == 3)
                                {
                                    middle.updateFoundedColor();
                                    print += color[i] + " " + color[j] + " " + color[q] + ": " + biggestValue + Environment.NewLine;
                                    continue;
                                }
                                else // n > 3
                                {
                                    for (int k = q + 1; k < model.getColCount() - n + 3; k++)
                                    {
                                        if (value[k] == 0 && (canStop || (value[n - 1] == 0 && value[3] != 0)))
                                        {
                                            break;
                                        }
                                        if (n == 4)
                                        {
                                            middle.updateFoundedColor();
                                            print += color[i] + " " + color[j] + " " + color[q] + "- " + color[k] + ": " + biggestValue + Environment.NewLine;
                                            continue;
                                        }
                                        else // n > 4
                                        {
                                            for (int l = k + 1; l < model.getColCount() - n + 4; l++)
                                            {
                                                if (value[l] == 0 && canStop)
                                                {
                                                    break;
                                                }
                                                middle.updateFoundedColor();
                                                print += color[i] + " " + color[j] + " " + color[q] + "- " + color[k] + " " + color[l] + ": " + biggestValue + Environment.NewLine;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    printOut += print;
                }
                else // màu đầu tiên không full 1
                {
                    for (int j = 0; j < model.getColCount() - n + 1; j++)
                    {
                        if (j == i)
                        {
                            continue;
                        }

                        if (checkToBreak(n, biggestValue, value[i] + value[j]) || (value[j] == 0 && (canStop || (value[n - 1] == 0 && value[1] != 0))))
                        {
                            break;
                        }

                        List<int> checkList2 = new List<int>(checkList1);

                        int currentCosts = 0; // trọng số cột hiện tại

                        foreach (int temp in checkList1) // đánh trọng số cho màu thứ 2
                        {
                            if (zeroOne[j][temp] == 1)
                            {
                                currentCosts++;
                                checkList2.Remove(temp);
                            }
                        }
                        currentValue2 = currentValue1 + currentCosts;

                        if (currentValue2 < biggestValue)
                        {
                            continue;
                        }

                        if (n == 2)
                        {
                            if (j > i || j < i && currentValue2 < max[j])
                            {
                                if (currentValue2 > biggestValue)
                                {
                                    biggestValue = currentValue2;
                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }
                                    if (index[i] < index[j])
                                    {

                                        print = color[i] + " " + color[j] + ": " + biggestValue + Environment.NewLine;
                                    }
                                    else
                                    {

                                        print = color[j] + " " + color[i] + ": " + biggestValue + Environment.NewLine;
                                    }

                                }
                                else if (currentValue2 == biggestValue)
                                {
                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }
                                    if (index[i] < index[j])
                                    {

                                        print += color[i] + " " + color[j] + ": " + biggestValue + Environment.NewLine;
                                    }
                                    else
                                    {

                                        print += color[j] + " " + color[i] + ": " + biggestValue + Environment.NewLine;
                                    }
                                }
                            }
                        }
                        else // n > 2, tìm nhóm 2 màu có OR lớn nhất để ghép
                        {
                            if (j > i || j < i && currentValue2 < max[j])
                            {
                                if (currentValue2 > biggestValue)
                                {
                                    biggestValue = currentValue2;
                                    savedRound2.Clear();
                                    savedRound2.Add(j);
                                }
                                else if (currentValue2 == biggestValue)
                                {
                                    savedRound2.Add(j);
                                }
                            }
                        }
                    }

                    if (n == 2)
                    {
                        middle.updateFoundedColor();
                        printOut += print;
                    }
                    else //(n > 2)
                    {
                        foreach (int j in savedRound2)
                        {
                            biggestValue = 0;

                            int currentCosts = 0; // trọng số cột hiện tại
                            List<int> checkList2 = new List<int>(checkList1);
                            foreach (int temp in checkList1) // đánh trọng số cho màu thứ 2
                            {
                                if (zeroOne[j][temp] == 1)
                                {
                                    currentCosts++;
                                    checkList2.Remove(temp);
                                }
                            }
                            currentValue2 = currentValue1 + currentCosts;

                            if (!checkList2.Any()) // 2 màu làm full 1 luôn
                            {
                                biggestValue = currentValue2;
                                if (biggestValue < limitedInputValue)
                                {
                                    break;
                                }

                                for (int q = 0; q < model.getColCount() - n + 2; q++)
                                {
                                    if (q == i || q == j)
                                    {
                                        continue;
                                    }

                                    if (value[q] == 0 && (canStop || (value[n - 1] == 0 && value[2] != 0)))
                                    {
                                        break;
                                    }

                                    if (n == 3)
                                    {
                                        String[] colorOut = new String[3];
                                        int[] colorOutIndex = new int[3];

                                        colorOut[0] = color[i];
                                        colorOut[1] = color[j];
                                        colorOut[2] = color[q];

                                        colorOutIndex[0] = index[i];
                                        colorOutIndex[1] = index[j];
                                        colorOutIndex[2] = index[q];

                                        for (int x = 0; x < 3; x++)
                                        {
                                            for (int y = x + 1; y < 3; y++)
                                            {
                                                if (colorOutIndex[x] > colorOutIndex[y])
                                                {
                                                    String temp;
                                                    temp = colorOut[x];
                                                    colorOut[x] = colorOut[y];
                                                    colorOut[y] = temp;

                                                    int tempInt;
                                                    tempInt = colorOutIndex[x];
                                                    colorOutIndex[x] = colorOutIndex[y];
                                                    colorOutIndex[y] = tempInt;
                                                }
                                            }
                                        }
                                        print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + ": " + biggestValue + Environment.NewLine;
                                        continue;
                                    }
                                    else // n > 3
                                    {
                                        for (int k = 0; k < model.getColCount() - n + 3; k++)
                                        {
                                            if (k == i || k == j || k == q)
                                            {
                                                continue;
                                            }

                                            if (value[k] == 0 && (canStop || (value[n - 1] == 0 && value[3] != 0)))
                                            {
                                                break;
                                            }

                                            if (n == 4)
                                            {
                                                String[] colorOut = new String[4];
                                                int[] colorOutIndex = new int[4];

                                                colorOut[0] = color[i];
                                                colorOut[1] = color[j];
                                                colorOut[2] = color[q];
                                                colorOut[3] = color[k];

                                                colorOutIndex[0] = index[i];
                                                colorOutIndex[1] = index[j];
                                                colorOutIndex[2] = index[q];
                                                colorOutIndex[3] = index[k];

                                                for (int x = 0; x < 4; x++)
                                                {
                                                    for (int y = x + 1; y < 4; y++)
                                                    {
                                                        if (colorOutIndex[x] > colorOutIndex[y])
                                                        {
                                                            String temp;
                                                            temp = colorOut[x];
                                                            colorOut[x] = colorOut[y];
                                                            colorOut[y] = temp;

                                                            int tempInt;
                                                            tempInt = colorOutIndex[x];
                                                            colorOutIndex[x] = colorOutIndex[y];
                                                            colorOutIndex[y] = tempInt;
                                                        }
                                                    }
                                                }
                                                print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;
                                                continue;
                                            }
                                            else // n > 4
                                            {
                                                for (int l = 0; l < model.getColCount() - n + 4; l++)
                                                {
                                                    if (l == i || l == j || l == q || l == k)
                                                    {
                                                        continue;
                                                    }

                                                    if (value[l] == 0 && canStop)
                                                    {
                                                        break;
                                                    }
                                                    String[] colorOut = new String[5];
                                                    int[] colorOutIndex = new int[5];

                                                    colorOut[0] = color[i];
                                                    colorOut[1] = color[j];
                                                    colorOut[2] = color[q];
                                                    colorOut[3] = color[k];
                                                    colorOut[4] = color[l];

                                                    colorOutIndex[0] = index[i];
                                                    colorOutIndex[1] = index[j];
                                                    colorOutIndex[2] = index[q];
                                                    colorOutIndex[3] = index[k];
                                                    colorOutIndex[4] = index[l];

                                                    for (int x = 0; x < 5; x++)
                                                    {
                                                        for (int y = x + 1; y < 5; y++)
                                                        {
                                                            if (colorOutIndex[x] > colorOutIndex[y])
                                                            {
                                                                String temp;
                                                                temp = colorOut[x];
                                                                colorOut[x] = colorOut[y];
                                                                colorOut[y] = temp;

                                                                int tempInt;
                                                                tempInt = colorOutIndex[x];
                                                                colorOutIndex[x] = colorOutIndex[y];
                                                                colorOutIndex[y] = tempInt;
                                                            }
                                                        }
                                                    }
                                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;
                                                }
                                            }
                                        }
                                    }
                                }
                                printOut += print;
                            }
                            else // 2 màu không full 1
                            {
                                for (int q = 0; q < model.getColCount() - n + 2; q++)
                                {
                                    if (q == i || q == j)
                                    {
                                        continue;
                                    }

                                    if (checkToBreak(n, biggestValue, value[i] + value[j] + value[q]) || (value[q] == 0 && (canStop || (value[n - 1] == 0 && value[2] != 0))))
                                    {
                                        break;
                                    }

                                    List<int> checkList3 = new List<int>(checkList2);
                                    currentCosts = 0;

                                    foreach (int temp in checkList2)
                                    {
                                        if (zeroOne[q][temp] == 1)
                                        {
                                            currentCosts++;
                                            checkList3.Remove(temp);
                                        }
                                    }
                                    currentValue3 = currentValue2 + currentCosts;

                                    if (currentValue3 < biggestValue)
                                    {
                                        continue;
                                    }

                                    if (n == 3)
                                    {
                                        if (currentValue3 > biggestValue)
                                        {
                                            biggestValue = currentValue3;
                                            if (biggestValue < limitedInputValue)
                                            {
                                                continue;
                                            }

                                            String[] colorOut = new String[3];
                                            int[] colorOutIndex = new int[3];

                                            colorOut[0] = color[i];
                                            colorOut[1] = color[j];
                                            colorOut[2] = color[q];

                                            colorOutIndex[0] = index[i];
                                            colorOutIndex[1] = index[j];
                                            colorOutIndex[2] = index[q];

                                            for (int x = 0; x < 3; x++)
                                            {
                                                for (int y = x + 1; y < 3; y++)
                                                {
                                                    if (colorOutIndex[x] > colorOutIndex[y])
                                                    {
                                                        String temp;
                                                        temp = colorOut[x];
                                                        colorOut[x] = colorOut[y];
                                                        colorOut[y] = temp;

                                                        int tempInt;
                                                        tempInt = colorOutIndex[x];
                                                        colorOutIndex[x] = colorOutIndex[y];
                                                        colorOutIndex[y] = tempInt;
                                                    }
                                                }
                                            }
                                            print = colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + ": " + biggestValue + Environment.NewLine;
                                        }
                                        else if (currentValue3 == biggestValue)
                                        {
                                            biggestValue = currentValue3;
                                            if (biggestValue < limitedInputValue)
                                            {
                                                continue;
                                            }
                                            String[] colorOut = new String[3];
                                            int[] colorOutIndex = new int[3];

                                            colorOut[0] = color[i];
                                            colorOut[1] = color[j];
                                            colorOut[2] = color[q];

                                            colorOutIndex[0] = index[i];
                                            colorOutIndex[1] = index[j];
                                            colorOutIndex[2] = index[q];

                                            for (int x = 0; x < 3; x++)
                                            {
                                                for (int y = x + 1; y < 3; y++)
                                                {
                                                    if (colorOutIndex[x] > colorOutIndex[y])
                                                    {
                                                        String temp;
                                                        temp = colorOut[x];
                                                        colorOut[x] = colorOut[y];
                                                        colorOut[y] = temp;

                                                        int tempInt;
                                                        tempInt = colorOutIndex[x];
                                                        colorOutIndex[x] = colorOutIndex[y];
                                                        colorOutIndex[y] = tempInt;
                                                    }
                                                }
                                            }
                                            print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + ": " + biggestValue + Environment.NewLine;
                                        }
                                    }
                                    else
                                    {
                                        if (currentValue3 > biggestValue)
                                        {
                                            biggestValue = currentValue3;
                                            savedRound3.Clear();
                                            savedRound3.Add(q);
                                        }
                                        else if (currentValue3 == biggestValue)
                                        {
                                            savedRound3.Add(q);
                                        }
                                    }
                                }
                            }

                            if (n == 3)
                            {
                                middle.updateFoundedColor();
                                printOut += print;
                            }
                            else //(n > 3)
                            {
                                foreach (int q in savedRound3)
                                {
                                    biggestValue = 0;

                                    List<int> checkList3 = new List<int>(checkList2);
                                    currentCosts = 0;

                                    foreach (int temp in checkList2) // đánh trọng số cho màu thứ 2
                                    {
                                        if (zeroOne[q][temp] == 1)
                                        {
                                            currentCosts++;
                                            checkList3.Remove(temp);
                                        }
                                    }
                                    currentValue3 = currentValue2 + currentCosts;

                                    if (!checkList3.Any()) // 3 mau lam full 1
                                    {
                                        biggestValue = currentValue3;
                                        if (biggestValue < limitedInputValue)
                                        {
                                            break;
                                        }

                                        for (int k = 0; k < model.getColCount() - n + 3; k++)
                                        {
                                            if (k == i || k == j || k == q)
                                            {
                                                continue;
                                            }

                                            if (value[k] == 0 && (canStop || (value[n - 1] == 0 && value[3] != 0)))
                                            {
                                                break;
                                            }

                                            if (n == 4)
                                            {
                                                String[] colorOut = new String[4];
                                                int[] colorOutIndex = new int[4];

                                                colorOut[0] = color[i];
                                                colorOut[1] = color[j];
                                                colorOut[2] = color[q];
                                                colorOut[3] = color[k];

                                                colorOutIndex[0] = index[i];
                                                colorOutIndex[1] = index[j];
                                                colorOutIndex[2] = index[q];
                                                colorOutIndex[3] = index[k];

                                                for (int x = 0; x < 4; x++)
                                                {
                                                    for (int y = x + 1; y < 4; y++)
                                                    {
                                                        if (colorOutIndex[x] > colorOutIndex[y])
                                                        {
                                                            String temp;
                                                            temp = colorOut[x];
                                                            colorOut[x] = colorOut[y];
                                                            colorOut[y] = temp;

                                                            int tempInt;
                                                            tempInt = colorOutIndex[x];
                                                            colorOutIndex[x] = colorOutIndex[y];
                                                            colorOutIndex[y] = tempInt;
                                                        }
                                                    }
                                                }
                                                print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;
                                                continue;
                                            }
                                            else // n > 4
                                            {
                                                for (int l = 0; l < model.getColCount() - n + 4; l++)
                                                {
                                                    if (l == i || l == j || l == q || l == k)
                                                    {
                                                        continue;
                                                    }

                                                    if (value[l] == 0 && canStop)
                                                    {
                                                        break;
                                                    }
                                                    String[] colorOut = new String[5];
                                                    int[] colorOutIndex = new int[5];

                                                    colorOut[0] = color[i];
                                                    colorOut[1] = color[j];
                                                    colorOut[2] = color[q];
                                                    colorOut[3] = color[k];
                                                    colorOut[4] = color[l];

                                                    colorOutIndex[0] = index[i];
                                                    colorOutIndex[1] = index[j];
                                                    colorOutIndex[2] = index[q];
                                                    colorOutIndex[3] = index[k];
                                                    colorOutIndex[4] = index[l];

                                                    for (int x = 0; x < 5; x++)
                                                    {
                                                        for (int y = x + 1; y < 5; y++)
                                                        {
                                                            if (colorOutIndex[x] > colorOutIndex[y])
                                                            {
                                                                String temp;
                                                                temp = colorOut[x];
                                                                colorOut[x] = colorOut[y];
                                                                colorOut[y] = temp;

                                                                int tempInt;
                                                                tempInt = colorOutIndex[x];
                                                                colorOutIndex[x] = colorOutIndex[y];
                                                                colorOutIndex[y] = tempInt;
                                                            }
                                                        }
                                                    }
                                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;
                                                }
                                            }

                                        }
                                        printOut += print;
                                    }
                                    else // 3 mau khong full 1
                                    {
                                        for (int k = 0; k < model.getColCount() - n + 3; k++)
                                        {
                                            if (k == i || k == j || k == q)
                                            {
                                                continue;
                                            }

                                            // điều kiện dừng
                                            if (checkToBreak(n, biggestValue, value[i] + value[j] + value[q] + value[k]) || (value[k] == 0 && (canStop || (value[n - 1] == 0 && value[3] != 0))))
                                            {
                                                break;
                                            }
                                            
                                            List<int> checkList4 = new List<int>(checkList3);
                                            currentCosts = 0;

                                            foreach (int temp in checkList3) // đánh trọng số cho màu thứ 2
                                            {
                                                if (zeroOne[k][temp] == 1)
                                                {
                                                    currentCosts++;
                                                    checkList4.Remove(temp);
                                                }
                                            }
                                            currentValue4 = currentValue3 + currentCosts;

                                            if (currentValue4 < biggestValue)
                                            {
                                                continue;
                                            }

                                            if (n == 4)
                                            {
                                                if (currentValue4 > biggestValue)
                                                {
                                                    biggestValue = currentValue4;
                                                    if (biggestValue < limitedInputValue)
                                                    {
                                                        continue;
                                                    }

                                                    String[] colorOut = new String[4];
                                                    int[] colorOutIndex = new int[4];

                                                    colorOut[0] = color[i];
                                                    colorOut[1] = color[j];
                                                    colorOut[2] = color[q];
                                                    colorOut[3] = color[k];

                                                    colorOutIndex[0] = index[i];
                                                    colorOutIndex[1] = index[j];
                                                    colorOutIndex[2] = index[q];
                                                    colorOutIndex[3] = index[k];

                                                    for (int x = 0; x < 4; x++)
                                                    {
                                                        for (int y = x + 1; y < 4; y++)
                                                        {
                                                            if (colorOutIndex[x] > colorOutIndex[y])
                                                            {
                                                                String temp;
                                                                temp = colorOut[x];
                                                                colorOut[x] = colorOut[y];
                                                                colorOut[y] = temp;

                                                                int tempInt;
                                                                tempInt = colorOutIndex[x];
                                                                colorOutIndex[x] = colorOutIndex[y];
                                                                colorOutIndex[y] = tempInt;
                                                            }
                                                        }
                                                    }
                                                    print = colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;

                                                }
                                                else if (currentValue4 == biggestValue)
                                                {
                                                    if (biggestValue < limitedInputValue)
                                                    {
                                                        continue;
                                                    }
                                                    String[] colorOut = new String[4];
                                                    int[] colorOutIndex = new int[4];

                                                    colorOut[0] = color[i];
                                                    colorOut[1] = color[j];
                                                    colorOut[2] = color[q];
                                                    colorOut[3] = color[k];

                                                    colorOutIndex[0] = index[i];
                                                    colorOutIndex[1] = index[j];
                                                    colorOutIndex[2] = index[q];
                                                    colorOutIndex[3] = index[k];

                                                    for (int x = 0; x < 4; x++)
                                                    {
                                                        for (int y = x + 1; y < 4; y++)
                                                        {
                                                            if (colorOutIndex[x] > colorOutIndex[y])
                                                            {
                                                                String temp;
                                                                temp = colorOut[x];
                                                                colorOut[x] = colorOut[y];
                                                                colorOut[y] = temp;

                                                                int tempInt;
                                                                tempInt = colorOutIndex[x];
                                                                colorOutIndex[x] = colorOutIndex[y];
                                                                colorOutIndex[y] = tempInt;
                                                            }
                                                        }
                                                    }
                                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;
                                                }
                                            }
                                            else
                                            {
                                                if (currentValue4 > biggestValue)
                                                {
                                                    biggestValue = currentValue4;
                                                    savedRound4.Clear();
                                                    savedRound4.Add(k);
                                                }
                                                else if (currentValue4 == biggestValue)
                                                {
                                                    savedRound4.Add(k);
                                                }
                                            }

                                        }
                                    }

                                    if (n == 4)
                                    {
                                        middle.updateFoundedColor();
                                        printOut += print;
                                    }
                                    else //(n > 4)
                                    {
                                        foreach (int k in savedRound4)
                                        {
                                            biggestValue = 0;

                                            List<int> checkList4 = new List<int>(checkList3);
                                            currentCosts = 0;

                                            foreach (int temp in checkList3) // đánh trọng số cho màu thứ 2
                                            {
                                                if (zeroOne[k][temp] == 1)
                                                {
                                                    currentCosts++;
                                                    checkList4.Remove(temp);
                                                }
                                            }
                                            currentValue4 = currentValue3 + currentCosts;

                                            if (!checkList4.Any()) // 4 mau full 1
                                            {
                                                biggestValue = currentValue4;
                                                if (biggestValue < limitedInputValue)
                                                {
                                                    continue;
                                                }

                                                for (int l = 0; l < model.getColCount() - n + 4; l++)
                                                {
                                                    if (l == i || l == j || l == q || l == k)
                                                    {
                                                        continue;
                                                    }

                                                    if (value[l] == 0 && canStop)
                                                    {
                                                        break;
                                                    }

                                                    String[] colorOut = new String[5];
                                                    int[] colorOutIndex = new int[5];

                                                    colorOut[0] = color[i];
                                                    colorOut[1] = color[j];
                                                    colorOut[2] = color[q];
                                                    colorOut[3] = color[k];
                                                    colorOut[4] = color[l];

                                                    colorOutIndex[0] = index[i];
                                                    colorOutIndex[1] = index[j];
                                                    colorOutIndex[2] = index[q];
                                                    colorOutIndex[3] = index[k];
                                                    colorOutIndex[4] = index[l];

                                                    for (int x = 0; x < 5; x++)
                                                    {
                                                        for (int y = x + 1; y < 5; y++)
                                                        {
                                                            if (colorOutIndex[x] > colorOutIndex[y])
                                                            {
                                                                String temp;
                                                                temp = colorOut[x];
                                                                colorOut[x] = colorOut[y];
                                                                colorOut[y] = temp;

                                                                int tempInt;
                                                                tempInt = colorOutIndex[x];
                                                                colorOutIndex[x] = colorOutIndex[y];
                                                                colorOutIndex[y] = tempInt;
                                                            }
                                                        }
                                                    }

                                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;
                                                }
                                            }
                                            else // 4 mau khong full 1
                                            {
                                                for (int l = 0; l < model.getColCount() - n + 4; l++)
                                                {
                                                    if (l == i || l == j || l == q || l == k)
                                                    {
                                                        continue;
                                                    }
                                                    
                                                    List<int> checkList5 = new List<int>(checkList4);
                                                    currentCosts = 0;

                                                    foreach (int temp in checkList4) // đánh trọng số cho màu thứ 2
                                                    {
                                                        if (zeroOne[l][temp] == 1)
                                                        {
                                                            currentCosts++;
                                                            checkList5.Remove(temp);
                                                        }
                                                    }
                                                    currentValue5 = currentValue4 + currentCosts;

                                                    if (currentValue4 + currentCosts > biggestValue)
                                                    {
                                                        biggestValue = currentValue4 + currentCosts;
                                                        if (biggestValue < limitedInputValue)
                                                        {
                                                            continue;
                                                        }

                                                        String[] colorOut = new String[5];
                                                        int[] colorOutIndex = new int[5];

                                                        colorOut[0] = color[i];
                                                        colorOut[1] = color[j];
                                                        colorOut[2] = color[q];
                                                        colorOut[3] = color[k];
                                                        colorOut[4] = color[l];

                                                        colorOutIndex[0] = index[i];
                                                        colorOutIndex[1] = index[j];
                                                        colorOutIndex[2] = index[q];
                                                        colorOutIndex[3] = index[k];
                                                        colorOutIndex[4] = index[l];

                                                        for (int x = 0; x < 5; x++)
                                                        {
                                                            for (int y = x + 1; y < 5; y++)
                                                            {
                                                                if (colorOutIndex[x] > colorOutIndex[y])
                                                                {
                                                                    String temp;
                                                                    temp = colorOut[x];
                                                                    colorOut[x] = colorOut[y];
                                                                    colorOut[y] = temp;

                                                                    int tempInt;
                                                                    tempInt = colorOutIndex[x];
                                                                    colorOutIndex[x] = colorOutIndex[y];
                                                                    colorOutIndex[y] = tempInt;
                                                                }
                                                            }
                                                        }
                                                        print = colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;
                                                    }
                                                    else if (currentValue4 + currentCosts == biggestValue)
                                                    {
                                                        if (biggestValue < limitedInputValue)
                                                        {
                                                            continue;
                                                        }
                                                        String[] colorOut = new String[5];
                                                        int[] colorOutIndex = new int[5];

                                                        colorOut[0] = color[i];
                                                        colorOut[1] = color[j];
                                                        colorOut[2] = color[q];
                                                        colorOut[3] = color[k];
                                                        colorOut[4] = color[l];

                                                        colorOutIndex[0] = index[i];
                                                        colorOutIndex[1] = index[j];
                                                        colorOutIndex[2] = index[q];
                                                        colorOutIndex[3] = index[k];
                                                        colorOutIndex[4] = index[l];

                                                        for (int x = 0; x < 5; x++)
                                                        {
                                                            for (int y = x + 1; y < 5; y++)
                                                            {
                                                                if (colorOutIndex[x] > colorOutIndex[y])
                                                                {
                                                                    String temp;
                                                                    temp = colorOut[x];
                                                                    colorOut[x] = colorOut[y];
                                                                    colorOut[y] = temp;

                                                                    int tempInt;
                                                                    tempInt = colorOutIndex[x];
                                                                    colorOutIndex[x] = colorOutIndex[y];
                                                                    colorOutIndex[y] = tempInt;
                                                                }
                                                            }
                                                        }
                                                        print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;
                                                    }
                                                }
                                            }
                                            middle.updateFoundedColor();
                                            printOut += print;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                max[i] = biggestValue;
                if (value[i + 1] < value[i])
                {
                    break;
                }
            }
            using (System.IO.StreamWriter writetext = new System.IO.StreamWriter(n + "-output.txt"))
            {
                writetext.WriteLine(printOut);
            }
        }

        public void processGroupAll2(int nColorChose, string color1, string color2) // print All n = 2
        {
            exc.readExcelSortByColor(nColorChose, model.getColor(), model.getValue(), model.getIndex(), model.getZeroOne(), color1, color2, "", "", "");
            int n = model.getN();
            int limitedInputValue = model.getLimit();
            string print = "";
            int biggestValue = 0; // giá trị lớn nhất khi gộp 2 cột
            int[] value = model.getValue();
            int[][] zeroOne = model.getZeroOne();
            int[] index = model.getIndex();
            string[] color = model.getColor();

            n = 2;


            if (nColorChose == 0) // truờng hợp mặc định: in bt
            {
                for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                {
                    print = "";
                    for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                    {
                        biggestValue = 0;
                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                        {
                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1)
                            {
                                biggestValue++;
                            }
                        }
                        if (biggestValue < limitedInputValue)
                        {
                            continue;
                        }
                        middle.updateFoundedColor();
                        print += color[i] + " " + color[j] + ": " + biggestValue + Environment.NewLine;
                    }
                    using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("2-outputall.txt", true))
                    {
                        writetext.Write(print);
                    }
                }
            }
            else
            {
                for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                {
                    print = "";
                    for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                    {
                        biggestValue = 0;
                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                        {
                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1)
                            {
                                biggestValue++;
                            }
                        }

                        if (biggestValue < limitedInputValue) // không đạt ngưỡng giới hạn
                        {
                            continue;
                        }

                        if (index[i] < index[j]) // sắp xếp đầu ra
                        {
                            print = color[i] + " " + color[j] + ": " + biggestValue + Environment.NewLine;
                        }
                        else
                        {
                            print = color[j] + " " + color[i] + ": " + biggestValue + Environment.NewLine;
                        }
                        middle.updateFoundedColor();
                        if (nColorChose == 2)
                        {
                            break;
                        }

                    }
                    using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("2-outputall.txt", true))
                    {
                        writetext.Write(print);
                    }

                    if (nColorChose >= 1)
                    {
                        break;
                    }
                }
            }
        }

        public void processGroupAll3(int nColorChose, string color1, string color2, string color3) // print All n = 3
        {
            thietlaphesoModel model = new thietlaphesoModel();
            exc.readExcelSortByColor(nColorChose, model.getColor(), model.getValue(), model.getIndex(), model.getZeroOne(), color1, color2, color3, "", "");
            int n = model.getN();
            int limitedInputValue = model.getLimit();
            string print = "";
            int biggestValue = 0; // giá trị lớn nhất khi gộp 2 cột
            int[] value = model.getValue();
            int[][] zeroOne = model.getZeroOne();
            int[] index = model.getIndex();
            string[] color = model.getColor();

            n = 3;

            if (nColorChose == 0) // truờng hợp mặc định: in bt
            {
                if (model.getColCount() < 500)
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        print = "";
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                biggestValue = 0;
                                for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                {
                                    if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1)
                                    {
                                        biggestValue++;
                                    }
                                }

                                if (biggestValue < limitedInputValue)
                                {
                                    continue;
                                }
                                middle.updateFoundedColor();
                                print += color[i] + " " + color[j] + " " + color[q] + ": " + biggestValue + Environment.NewLine;
                            }
                        }
                        using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("3-outputall.txt", true))
                        {
                            writetext.Write(print);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            print = "";
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                biggestValue = 0;
                                for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                {
                                    if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1)
                                    {
                                        biggestValue++;
                                    }
                                }

                                if (biggestValue < limitedInputValue)
                                {
                                    continue;
                                }
                                middle.updateFoundedColor();
                                print += color[i] + " " + color[j] + " " + color[q] + ": " + biggestValue + Environment.NewLine;
                            }
                            using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("3-outputall.txt", true))
                            {
                                writetext.Write(print);
                            }
                        }
                    }
                }

            }
            else// in theo các màu người dùng nhập
            {

                for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                {
                    print = "";
                    for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                    {
                        for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                        {
                            biggestValue = 0;
                            for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                            {
                                if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1)
                                {
                                    biggestValue++;
                                }
                            }

                            if (biggestValue < limitedInputValue)
                            {
                                continue;
                            }

                            String[] colorOut = new String[3];
                            int[] colorOutIndex = new int[3];

                            colorOut[0] = color[i];
                            colorOut[1] = color[j];
                            colorOut[2] = color[q];

                            colorOutIndex[0] = index[i];
                            colorOutIndex[1] = index[j];
                            colorOutIndex[2] = index[q];

                            for (int x = 0; x < 3; x++)
                            {
                                for (int y = x + 1; y < 3; y++)
                                {
                                    if (colorOutIndex[x] > colorOutIndex[y])
                                    {
                                        String temp;
                                        temp = colorOut[x];
                                        colorOut[x] = colorOut[y];
                                        colorOut[y] = temp;

                                        int tempInt;
                                        tempInt = colorOutIndex[x];
                                        colorOutIndex[x] = colorOutIndex[y];
                                        colorOutIndex[y] = tempInt;
                                    }
                                }
                            }
                            middle.updateFoundedColor();
                            print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + ": " + biggestValue + Environment.NewLine;

                            if (nColorChose >= 3)
                            {
                                break;
                            }
                        }
                        if (nColorChose >= 2)
                        {
                            break;
                        }
                    }
                    using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("3-outputall.txt", true))
                    {
                        writetext.Write(print);
                    }
                    if (nColorChose >= 1)
                    {
                        break;
                    }
                }
            }
        }

        public void processGroupAll4(int nColorChose, string color1, string color2, string color3, string color4) // print All n = 4
        {
            thietlaphesoModel model = new thietlaphesoModel();
            exc.readExcelSortByColor(nColorChose, model.getColor(), model.getValue(), model.getIndex(), model.getZeroOne(), color1, color2, color3, color4, "");
            int n = model.getN();
            int limitedInputValue = model.getLimit();
            string print = "";
            int biggestValue = 0; // giá trị lớn nhất khi gộp 2 cột
            int[] value = model.getValue();
            int[][] zeroOne = model.getZeroOne();
            int[] index = model.getIndex();
            string[] color = model.getColor();

            n = 4;

            if (nColorChose == 0) // in bt
            {
                if (model.getColCount() < 500)
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            print = "";
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    biggestValue = 0;
                                    for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                    {
                                        if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1)
                                        {
                                            biggestValue++;
                                        }
                                    }

                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }
                                    middle.updateFoundedColor();
                                    print += color[i] + " " + color[j] + " " + color[q] + " " + color[k] + ": " + biggestValue + Environment.NewLine;
                                }
                            }
                            using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("4-outputall.txt", true))
                            {
                                writetext.Write(print);
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                print = "";
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    biggestValue = 0;
                                    for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                    {
                                        if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1)
                                        {
                                            biggestValue++;
                                        }
                                    }

                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }
                                    middle.updateFoundedColor();
                                    print += color[i] + " " + color[j] + " " + color[q] + " " + color[k] + ": " + biggestValue + Environment.NewLine;
                                }
                                using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("4-outputall.txt", true))
                                {
                                    writetext.Write(print);
                                }
                            }
                        }
                    }
                }

            }
            else // in theo nhập vào của người dùng
            {
                if (model.getColCount() < 500)
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            print = "";
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    biggestValue = 0;
                                    for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                    {
                                        if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1)
                                        {
                                            biggestValue++;
                                        }
                                    }

                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }

                                    String[] colorOut = new String[4];
                                    int[] colorOutIndex = new int[4];

                                    colorOut[0] = color[i];
                                    colorOut[1] = color[j];
                                    colorOut[2] = color[q];
                                    colorOut[3] = color[k];

                                    colorOutIndex[0] = index[i];
                                    colorOutIndex[1] = index[j];
                                    colorOutIndex[2] = index[q];
                                    colorOutIndex[3] = index[k];

                                    for (int x = 0; x < 4; x++)
                                    {
                                        for (int y = x + 1; y < 4; y++)
                                        {
                                            if (colorOutIndex[x] > colorOutIndex[y])
                                            {
                                                String temp;
                                                temp = colorOut[x];
                                                colorOut[x] = colorOut[y];
                                                colorOut[y] = temp;

                                                int tempInt;
                                                tempInt = colorOutIndex[x];
                                                colorOutIndex[x] = colorOutIndex[y];
                                                colorOutIndex[y] = tempInt;
                                            }
                                        }
                                    }
                                    middle.updateFoundedColor();
                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;

                                    if (nColorChose >= 4)
                                    {
                                        break;
                                    }
                                }
                                if (nColorChose >= 3)
                                {
                                    break;
                                }
                            }
                            using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("4-outputall.txt", true))
                            {
                                writetext.Write(print);
                            }
                            if (nColorChose >= 2)
                            {
                                break;
                            }
                        }
                        if (nColorChose >= 1)
                        {
                            break;
                        }
                    }
                }

                else
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {

                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                print = "";
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    biggestValue = 0;
                                    for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                    {
                                        if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1)
                                        {
                                            biggestValue++;
                                        }
                                    }

                                    if (biggestValue < limitedInputValue)
                                    {
                                        continue;
                                    }

                                    String[] colorOut = new String[4];
                                    int[] colorOutIndex = new int[4];

                                    colorOut[0] = color[i];
                                    colorOut[1] = color[j];
                                    colorOut[2] = color[q];
                                    colorOut[3] = color[k];

                                    colorOutIndex[0] = index[i];
                                    colorOutIndex[1] = index[j];
                                    colorOutIndex[2] = index[q];
                                    colorOutIndex[3] = index[k];

                                    for (int x = 0; x < 4; x++)
                                    {
                                        for (int y = x + 1; y < 4; y++)
                                        {
                                            if (colorOutIndex[x] > colorOutIndex[y])
                                            {
                                                String temp;
                                                temp = colorOut[x];
                                                colorOut[x] = colorOut[y];
                                                colorOut[y] = temp;

                                                int tempInt;
                                                tempInt = colorOutIndex[x];
                                                colorOutIndex[x] = colorOutIndex[y];
                                                colorOutIndex[y] = tempInt;
                                            }
                                        }
                                    }
                                    middle.updateFoundedColor();
                                    print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + ": " + biggestValue + Environment.NewLine;

                                    if (nColorChose >= 4)
                                    {
                                        break;
                                    }
                                }
                                if (nColorChose >= 3)
                                {
                                    break;
                                }
                                using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("4-outputall.txt", true))
                                {
                                    writetext.Write(print);
                                }
                            }

                            if (nColorChose >= 2)
                            {
                                break;
                            }
                        }
                        if (nColorChose >= 1)
                        {
                            break;
                        }
                    }
                }
            }
        }

        public void processGroupAll5(int nColorChose, string color1, string color2, string color3, string color4, string color5) // print All n = 5
        {
            thietlaphesoModel model = new thietlaphesoModel();
            exc.readExcelSortByColor(nColorChose, model.getColor(), model.getValue(), model.getIndex(), model.getZeroOne(), color1, color2, color3, color4, color5);
            int n = model.getN();
            string print = "";
            int limitedInputValue = model.getLimit();
            int biggestValue = 0; // giá trị lớn nhất khi gộp 2 cột
            int[] value = model.getValue();
            int[][] zeroOne = model.getZeroOne();
            int[] index = model.getIndex();
            string[] color = model.getColor();

            n = 5;

            if (nColorChose == 0)
            {
                if (model.getColCount() < 500)
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                print = "";
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    for (int l = k + 1; l < model.getColCount() - n + 4 - ExcelController.duplicateindex.Count; l++)
                                    {
                                        biggestValue = 0;
                                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                        {
                                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1 || zeroOne[l][temp] == 1)
                                            {
                                                biggestValue++;
                                            }
                                        }

                                        if (biggestValue < limitedInputValue)
                                        {
                                            continue;
                                        }
                                        middle.updateFoundedColor();
                                        print += color[i] + " " + color[j] + " " + color[q] + " " + color[k] + " " + color[l] + ": " + biggestValue + Environment.NewLine;
                                    }
                                }
                                using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("5-outputall.txt", true))
                                {
                                    writetext.Write(print);
                                }

                            }

                        }
                    }
                }
                else
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    print = "";
                                    for (int l = k + 1; l < model.getColCount() - n + 4 - ExcelController.duplicateindex.Count; l++)
                                    {
                                        biggestValue = 0;
                                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                        {
                                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1 || zeroOne[l][temp] == 1)
                                            {
                                                biggestValue++;
                                            }
                                        }

                                        if (biggestValue < limitedInputValue)
                                        {
                                            continue;
                                        }
                                        middle.updateFoundedColor();
                                        print += color[i] + " " + color[j] + " " + color[q] + " " + color[k] + " " + color[l] + ": " + biggestValue + Environment.NewLine;
                                    }
                                    using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("5-outputall.txt", true))
                                    {
                                        writetext.Write(print);
                                    }
                                }
                            }
                        }
                    }
                }            
            }
            else //in theo nguoi dung nhap
            {
                if (model.getColCount() < 500)
                {
                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                print = "";
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    for (int l = k + 1; l < model.getColCount() - n + 4 - ExcelController.duplicateindex.Count; l++)
                                    {
                                        biggestValue = 0;
                                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                        {
                                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1 || zeroOne[l][temp] == 1)
                                            {
                                                biggestValue++;
                                            }
                                        }

                                        if (biggestValue < limitedInputValue)
                                        {
                                            continue;
                                        }

                                        String[] colorOut = new String[5];
                                        int[] colorOutIndex = new int[5];

                                        colorOut[0] = color[i];
                                        colorOut[1] = color[j];
                                        colorOut[2] = color[q];
                                        colorOut[3] = color[k];
                                        colorOut[4] = color[l];

                                        colorOutIndex[0] = index[i];
                                        colorOutIndex[1] = index[j];
                                        colorOutIndex[2] = index[q];
                                        colorOutIndex[3] = index[k];
                                        colorOutIndex[4] = index[l];

                                        for (int x = 0; x < 5; x++)
                                        {
                                            for (int y = x + 1; y < 5; y++)
                                            {
                                                if (colorOutIndex[x] > colorOutIndex[y])
                                                {
                                                    String temp;
                                                    temp = colorOut[x];
                                                    colorOut[x] = colorOut[y];
                                                    colorOut[y] = temp;

                                                    int tempInt;
                                                    tempInt = colorOutIndex[x];
                                                    colorOutIndex[x] = colorOutIndex[y];
                                                    colorOutIndex[y] = tempInt;
                                                }
                                            }
                                        }
                                        middle.updateFoundedColor();
                                        print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;

                                        if (nColorChose == 5)
                                        {
                                            break;
                                        }
                                    }

                                    if (nColorChose >= 4)
                                    {
                                        break;
                                    }
                                }
                                using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("5-outputall.txt", true))
                                {
                                    writetext.Write(print);
                                }

                                if (nColorChose >= 3)
                                {
                                    break;
                                }

                            }
                            if (nColorChose >= 2)

                            {
                                break;
                            }

                        }
                        if (nColorChose >= 1)
                        {
                            break;
                        }
                    }
                }
                else
                {

                    for (int i = 0; i < model.getColCount() - n - ExcelController.duplicateindex.Count; i++)
                    {
                        for (int j = i + 1; j < model.getColCount() - n + 1 - ExcelController.duplicateindex.Count; j++)
                        {
                            for (int q = j + 1; q < model.getColCount() - n + 2 - ExcelController.duplicateindex.Count; q++)
                            {
                                for (int k = q + 1; k < model.getColCount() - n + 3 - ExcelController.duplicateindex.Count; k++)
                                {
                                    print = "";
                                    for (int l = k + 1; l < model.getColCount() - n + 4 - ExcelController.duplicateindex.Count; l++)
                                    {
                                        biggestValue = 0;
                                        for (int temp = 0; temp < ExcelController.ngayketthuc - ExcelController.ngaybatdau + 1; temp++)
                                        {
                                            if (zeroOne[i][temp] == 1 || zeroOne[j][temp] == 1 || zeroOne[q][temp] == 1 || zeroOne[k][temp] == 1 || zeroOne[l][temp] == 1)
                                            {
                                                biggestValue++;
                                            }
                                        }

                                        if (biggestValue < limitedInputValue)
                                        {
                                            continue;
                                        }

                                        String[] colorOut = new String[5];
                                        int[] colorOutIndex = new int[5];

                                        colorOut[0] = color[i];
                                        colorOut[1] = color[j];
                                        colorOut[2] = color[q];
                                        colorOut[3] = color[k];
                                        colorOut[4] = color[l];

                                        colorOutIndex[0] = index[i];
                                        colorOutIndex[1] = index[j];
                                        colorOutIndex[2] = index[q];
                                        colorOutIndex[3] = index[k];
                                        colorOutIndex[4] = index[l];

                                        for (int x = 0; x < 5; x++)
                                        {
                                            for (int y = x + 1; y < 5; y++)
                                            {
                                                if (colorOutIndex[x] > colorOutIndex[y])
                                                {
                                                    String temp;
                                                    temp = colorOut[x];
                                                    colorOut[x] = colorOut[y];
                                                    colorOut[y] = temp;

                                                    int tempInt;
                                                    tempInt = colorOutIndex[x];
                                                    colorOutIndex[x] = colorOutIndex[y];
                                                    colorOutIndex[y] = tempInt;
                                                }
                                            }
                                        }
                                        middle.updateFoundedColor();
                                        print += colorOut[0] + " " + colorOut[1] + " " + colorOut[2] + " " + colorOut[3] + " " + colorOut[4] + ": " + biggestValue + Environment.NewLine;

                                        if (nColorChose == 5)
                                        {
                                            break;
                                        }
                                    }

                                    if (nColorChose >= 4)
                                    {
                                        break;
                                    }
                                    using (System.IO.StreamWriter writetext = new System.IO.StreamWriter("5-outputall.txt", true))
                                    {
                                        writetext.Write(print);
                                    }
                                }
                                if (nColorChose >= 3)
                                {
                                    break;
                                }
                            }
                            if (nColorChose >= 2)
                            {
                                break;
                            }
                        }
                        if (nColorChose >= 1)
                        {
                            break;
                        }
                    }
                }
            }        
        }


        // biggestValue: Giá trị lớn nhất
        // valueCol: Giá trị của cột được chọn làm mốc
        public bool checkToBreak(int n, int biggestValue, int valueCol)
        {
            if (biggestValue % n == 0 && valueCol < biggestValue / n)
            {
                return true;
            }
            if (biggestValue % n != 0 && valueCol == biggestValue / n)
            {
                return true;
            }
            return false;
        }
    }
}
