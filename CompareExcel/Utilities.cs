using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareExcel
{
    class Utilities
    {
        public static string WriteRows(DataTable iTbl, string iLabel)
        {
            string results = "";
            foreach (DataRow row in iTbl.Rows)
            {
                foreach (DataColumn column in row.Table.Columns)
                {
                    results  += row[column] + ",";
                }
                results  += "\n";
            }
            return results;
        }

    }
}
