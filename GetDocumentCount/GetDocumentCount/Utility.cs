using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDocumentCount
{
    class Utility
    {

        public DataTable CreateDTFilterType()
        {
            DataTable dt = new DataTable();
            DataColumn FilterType = new DataColumn("FilterType");
            FilterType.DataType = typeof(string);
            dt.Columns.Add(FilterType);
            return dt;
        }

        public DataTable LoadFilterType()
        {
            DataTable dt = CreateDTFilterType();
            dt.Rows.Add("Ext");
            dt.Rows.Add("counterParty");
            dt.Rows.Add("masterAgreementNo");
            return dt;
        }
    }
}
