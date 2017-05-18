namespace DataProcessing.Model
{
    class MiddlewareModel
    {
        public static int foundedColor = 0;
        public static int foundedColor_MaxValue = 0;
        /// <summary>
        /// Set giá trị số màu đã tìm được
        /// </summary>
        public void setFoundedColor()
        {
            foundedColor += 1;
        }
        /// <summary>
        /// Lấy giá trị số màu đã tìm được
        /// </summary>
        /// <returns></returns>
        public int getFoundedColor()
        {
            return foundedColor;
        }

        /// <summary>
        /// Set giá trị số màu đã tìm được khi chọn tìm kiếm lớn nhất
        /// </summary>
        public void setFoundedColorMaxValue()
        {
            foundedColor_MaxValue += 1;
        }
        /// <summary>
        /// lấy giá trị số màu đã tìm được khi chọn tìm kiếm lớn nhất
        /// </summary>
        /// <returns></returns>
        public int getFoundedColorMaxValue()
        {
            return foundedColor_MaxValue;
        }

    }
}
