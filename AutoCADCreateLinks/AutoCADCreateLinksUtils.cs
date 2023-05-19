using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using SimioAPI;
using SimioAPI.Extensions;
using System.Text.RegularExpressions;
using System.Linq.Expressions;

namespace AutoCADCreateLinks
{
    internal static class AutoCADCreateLinksUtils
    {
        public static DataTable ConvertTableToDataTable(ITable table)
        {
            // get all column names
            List<string> colNames = new List<string>();
            // get property column names
            List<string> propColNames = new List<string>();

            foreach (var col in table.Columns)
            {
                colNames.Add(col.Name);
                propColNames.Add(col.Name);
            }
           
            // get state column names
            List<string> stateColNames = new List<string>();
            foreach (var stateCol in table.StateColumns)
            {
                colNames.Add(stateCol.Name);
                stateColNames.Add(stateCol.Name);
            }

            List<object[]> tableList = new List<object[]>();
            int rowNumber = 0;
            int colIdx = 0;
            // Get Row Data
            foreach (var row in table.Rows)
            {
                rowNumber++;
                colIdx = 0;
                List<object> thisRow = new List<object>();
                // get properties
                foreach (var array in propColNames)
                {
                    if (row.Properties[array.ToString()].Value != null)
                    {
                        thisRow.Add(row.Properties[array.ToString()].Value);
                    }
                    else
                    {
                        thisRow.Add("");
                    }
                    colIdx++;
                }
                // get states
                foreach (var array in stateColNames)
                {
                    if (table.StateRows.Count > 0 && table.StateRows[rowNumber - 1].StateValues[array.ToString()].PlanValue != null) thisRow.Add(table.StateRows[rowNumber - 1].StateValues[array.ToString()].PlanValue.ToString());
                    else thisRow.Add("");
                    colIdx++;
                }
                tableList.Add(thisRow.ToArray());
            }


            // New table.
            var dataTable = new DataTable();

            colIdx = 0;
            foreach (var col in colNames)
            {
                dataTable.Columns.Add(col);
                colIdx++;
            }
            
            // Add rows.
            foreach (var array in tableList)
            {
                dataTable.Rows.Add(array);
            }

            return dataTable;
        }
    }
}