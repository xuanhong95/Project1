using System;

namespace DataProcessing.Model
{
    public class thietlaphesoModel
    {
        public static int colcount = 1;
        public static int rowcount = 1;
        public static String[] color = new String[colcount - 1];
        public static String[] colordefault = new String[colcount - 1];
        public static String[] datetime = new String[rowcount - 1];
        public static int[] value = new int[colcount - 1];
        public static int[][] zeroOne = new int[colcount - 1][];
        public static int[] index = new int[colcount - 1];
        public static int n = 0;
        public static int limitvalue = 0;


        /// <summary>
        /// Set giá trị n
        /// </summary>
        /// <param name="limit"></param>
        public void setLimit(int limit)
        {
            limitvalue = limit; 
        }
        /// <summary>
        /// Get giá trị limit value
        /// </summary>
        /// <returns></returns>
        public int getLimit()
        {
            return limitvalue;
        }

        /// <summary>
        /// Set giá trị n
        /// </summary>
        /// <param name="n1"></param>
        public void setN(int n1)
        {
            n = n1;
        }
        /// <summary>
        /// Get giá trị n
        /// </summary>
        /// <returns></returns>
        public int getN()
        {
            return n;
        }
        /// <summary>
        /// Set giá trị mảng cột
        /// </summary>
        /// <param name="colcount1"></param>
        public void setColCount(int colcount1)
        {
            colcount = colcount1;
        }
        /// <summary>
        /// Get giá trị mảng cột
        /// </summary>
        /// <returns></returns>
        public int getColCount()
        {
            return colcount;
        }
        /// <summary>
        /// Set giá trị mảng hàng
        /// </summary>
        /// <param name="colcount1"></param>
        public void setRowCount(int rowcount1)
        {
            rowcount = rowcount1;
        }
        /// <summary>
        /// Get giá trị mảng hàng
        /// </summary>
        /// <returns></returns>
        public int getRowCount()
        {
            return rowcount;
        }
        /// <summary>
        /// Set giá trị mảng tên màu ko sắp xếp
        /// </summary>
        /// <param name="colordefault"></param>
        public void setColorDefault(String[] color12)
        {
            colordefault = color12;
        }
        /// <summary>
        /// Get giá trị mảng tên màu ko sắp xếp
        /// </summary>
        /// <returns></returns>
        public String[] getColorDefault()
        {
            return colordefault;
        }

        /// <summary>
        /// Set giá trị mảng tên màu ko sắp xếp
        /// </summary>
        /// <param name="colordefault"></param>
        public void setDateTime(String[] date)
        {
            datetime = date;
        }
        /// <summary>
        /// Get giá trị mảng tên màu ko sắp xếp
        /// </summary>
        /// <returns></returns>
        public String[] getDateTime()
        {
            return datetime;
        }

        /// <summary>
        /// Set giá trị mảng tên màu
        /// </summary>
        /// <param name="color"></param>
        public void setColor(String[] color1)
        {
            color = color1;
        }
        /// <summary>
        /// Get giá trị mảng tên màu
        /// </summary>
        /// <returns></returns>
        public String[] getColor()
        {
            return color;
        }
        /// <summary>
        /// Set giá trị mảng tổng màu bán được theo cột
        /// </summary>
        /// <param name="value"></param>
        public void setValue(int[] value1)
        {
            value = value1;
        }
        /// <summary>
        /// Get giá trị tổng màu bán được theo cột
        /// </summary>
        /// <returns></returns>
        public int[] getValue()
        {
            return value;
        }
        /// <summary>
        /// Set giá trị mảng cell theo cột
        /// </summary>
        /// <param name="zeroOne"></param>
        public void setZeroOne(int[][] zeroOne1)
        {
            zeroOne = zeroOne1;
        }
        /// <summary>
        /// Get giá trị mảng cell theo cột
        /// </summary>
        /// <returns></returns>
        public int[][] getZeroOne()
        {
            return zeroOne;
        }
        /// <summary>
        /// Set giá trị mảng số thứ tự theo excel
        /// </summary>
        /// <param name="index"></param>
        public void setIndex(int[] index1)
        {
            index = index1;
        }
        /// <summary>
        /// Get giá trị mảng số thứ tự theo excel
        /// </summary>
        /// <returns></returns>
        public int[] getIndex()
        {
            return index;
        }
    }
}
