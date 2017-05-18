using System;
using System.Collections.Generic;
using System.IO;
using DataProcessing.Model;

namespace DataProcessing.Controller
{
    class OutputController
    {
        MiddlewareModel middle = new MiddlewareModel();
        public void sortOutPut(int n)
        {
            List<string> color = new List<string>();
            List<int> value = new List<int>();
            Dictionary<string, int> hashmap = new Dictionary<string, int>();

            //Đọc file txt output và xử lý sắp xếp
            using (StreamReader sr = new StreamReader(n + "-output.txt"))
            {
                string line;
                while ((line = sr.ReadLine()) != "" || (line = sr.ReadLine()) != null)
                {
                    string[] s = line.Split((":").ToCharArray());
                    s[1] = s[1].Trim();
                    color.Add(s[0]);
                    value.Add(Int32.Parse(s[1]));
                }
                int tmp;
                string tmpcolor;
                for (int i = 0; i < value.Count; i++)
                {
                    for (int j = i + 1; j < value.Count; j++)
                    {
                        if (value[i] < value[j])
                        {
                            tmp = value[i];
                            tmpcolor = color[i];
                            value[i] = value[j];
                            color[i] = color[j];
                            value[j] = tmp;
                            color[j] = tmpcolor;
                        }
                    }
                }
            }

            for (int i = 0; i < value.Count; i++)
            {
                try
                {
                    hashmap.Add(color[i], value[i]);
                }
                catch (Exception e)
                {
                    continue;
                }
            }

            //In output mới
            using (StreamWriter writetext = new StreamWriter(n + "-output.txt"))
            {
                foreach (KeyValuePair<string, int> pair in hashmap)
                {
                    writetext.WriteLine("{0}: {1}", pair.Key, pair.Value);
                }
            }
        }

        public void countColorNumberMaxValue(int n)
        {
            using (StreamReader sr2 = new StreamReader(n + "-output.txt"))
            {
                string line;
                while ((line = sr2.ReadLine()) != null)
                {
                    middle.setFoundedColorMaxValue();
                }
            }
        }
    }
}
