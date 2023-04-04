using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CeBianLan
{
    public static class DataTableExtensions
    {
        public static void ToCSV(this DataTable table, string filePath)
        {
            // 创建CSV文件并写入表头
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(filePath))
            {
                string header = string.Join(",", table.Columns);
                writer.WriteLine(header);

                // 写入行数据
                foreach (DataRow row in table.Rows)
                {
                    string line = string.Join(",", row.ItemArray);
                    writer.WriteLine(line);
                }
            }
        }
    }
}
