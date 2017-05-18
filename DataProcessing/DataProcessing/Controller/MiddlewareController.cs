using DataProcessing.Model;
using System.Numerics;

namespace DataProcessing.Controller
{
    class MiddlewareController
    {
        MiddlewareModel model = new MiddlewareModel();
        thietlaphesoModel tlhs = new thietlaphesoModel();
        public static BigInteger tu = 1, mau = 1;
        public void updateFoundedColor()
        {
            model.setFoundedColor();
        }

        public int getFoundedColorValue()
        {
            return model.getFoundedColor();
        }
        public int getColorNumberMaxValue()
        {
            return model.getFoundedColorMaxValue();
        }

        
        public static BigInteger estimateTime(int n, int k)
        {
            for (int i = k +1; i <= n; i++)
            {
                tu *= i;
            }
            for (int  i = 1; i <= (n-k); i++)
            {
                mau *= i;
            }

            return tu / mau;
        }

        public int getExcelCol()
        {
            return tlhs.getColCount() - 1;
        }


    }
}
