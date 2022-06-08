using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Reflection;
using System.Text;

namespace CreateCsv
{
    public class Class1 : ControllerBase
    {
        public IActionResult GenerateCsv<T>(List<T> data)
        {
            var data1 = ExportListUsingEPPlus(data);
            try
            {
                return File(Encoding.ASCII.GetBytes(data1.ToString()), "text/csv", "SuntechUpload" + "-" + Guid.NewGuid() + ".csv");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                return null;
            }
        }

        public StringBuilder ExportListUsingEPPlus<T>(List<T> data)
        {
            var dtDataTable = ToDataTable(data);

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sb.Append(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sb.Append(",");
                }
            }
            sb.Append("\n");
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sb.Append(value);
                        }
                        else
                        {
                            sb.Append(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sb.Append(",");
                    }
                }
                sb.Append("\n");
            }
            return sb;
        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

    }
}