using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;

namespace FBLT_Tax.Utility
{
    public struct deSerializedFilter
    {
        public String query;
        public List<Object> param;
    }
    public class ColumnHelper
    {
        public static bool isValidColumn(String dataIndx)
        {
            //if (Regex.IsMatch(dataIndx, "^[a-z,A-Z]*$_"))
            //{
            //    return true;
            //}
            //else
            //{
            //    return false;
            //}
            return true;
        }
    }
    public class FilterHelper
    {
        struct Filter
        {
            public String dataIndx;
            public String condition;
            public String value;
            public String value2;
            public String dataType;
        }
        //map to json object posted by client
        struct FilterObj
        {
            public String mode;
            public List<Filter> data;
        }
        public static deSerializedFilter deSerializeFilter(String pq_filter)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();

            FilterObj filterObj = js.Deserialize<FilterObj>(pq_filter);
            String mode = filterObj.mode;
            List<Filter> filters = filterObj.data;

            List<String> fc = new List<String>();

            List<object> param = new List<object>();

            foreach (Filter filter in filters)
            {
                String dataIndx = filter.dataIndx;
                if (ColumnHelper.isValidColumn(dataIndx) == false)
                {
                    throw new Exception("Invalid column name");
                }
                String text = filter.value;
                String toValue = filter.value2;
                String condition = filter.condition;
                String dataType = filter.dataType;
                if (dataType == "date" && condition == "between")
                {
                    fc.Add("CONVERT(datetime," + dataIndx + ")" + " BETWEEN @" + dataIndx + " AND @" + dataIndx + "2");
                    param.Add(new SqlParameter(dataIndx, text));
                    param.Add(new SqlParameter(dataIndx + "2", toValue));
                }
                else if (condition == "contain")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text + "%"));
                }
                else if (condition == "notcontain")
                {
                    fc.Add(dataIndx + " not like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text + "%"));
                }
                else if (condition == "begin")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text + "%"));
                }
                else if (condition == "end")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text));
                }
                else if (condition == "equal")
                {
                    fc.Add(dataIndx + "=@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "notequal")
                {
                    fc.Add(dataIndx + "!=@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "empty")
                {
                    fc.Add("isnull(" + dataIndx + ",'')=''");
                    //param.Add(new SqlParameter(dataIndx, text + "%"));
                }
                else if (condition == "notempty")
                {
                    fc.Add("isnull(" + dataIndx + ",'')!=''");
                }
                else if (condition == "less")
                {
                    fc.Add(dataIndx + "<@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "great")
                {
                    fc.Add(dataIndx + ">@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
            }
            String query = "";
            if (filters.Count > 0 && fc.Count > 0)
            {
                query = " and " + String.Join(" " + mode + " ", fc.ToArray());
            }

            deSerializedFilter ds = new deSerializedFilter();
            ds.query = query;
            ds.param = param;
            return ds;
        }
        public static deSerializedFilter deSerializeFilter2(String pq_filter)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();

            FilterObj filterObj = js.Deserialize<FilterObj>(pq_filter);
            String mode = filterObj.mode;
            List<Filter> filters = filterObj.data;

            List<String> fc = new List<String>();

            List<object> param = new List<object>();

            foreach (Filter filter in filters)
            {
                String dataIndx = filter.dataIndx;
                if (ColumnHelper.isValidColumn(dataIndx) == false)
                {
                    throw new Exception("Invalid column name");
                }
                String text = filter.value;
                String toValue = filter.value2;
                String condition = filter.condition;
                String dataType = filter.dataType;
                if (dataType == "date" && condition == "between")
                {
                    fc.Add("CONVERT(datetime," + dataIndx + ")" + " BETWEEN @" + dataIndx + " AND @" + dataIndx + "2");
                    param.Add(new SqlParameter(dataIndx, text.Split('/')[2] + "-" + text.Split('/')[0] + "-" + text.Split('/')[1]));
                    param.Add(new SqlParameter(dataIndx + "2", toValue.Split('/')[2] + "-" + toValue.Split('/')[0] + "-" + toValue.Split('/')[1]));
                }//gte
                else if (dataType == "date" && condition == "gte")
                {
                    fc.Add("CONVERT(datetime," + dataIndx + ")" + " >= @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text.Split('/')[2] + "-" + text.Split('/')[0] + "-" + text.Split('/')[1]));
                }
                else if (dataType == "integer" && condition == "between")
                {

                    fc.Add("CONVERT(int," + dataIndx + ")" + " BETWEEN @" + dataIndx + " AND @" + dataIndx + "2");
                    param.Add(new SqlParameter(dataIndx, text));
                    param.Add(new SqlParameter(dataIndx + "2", toValue));

                }
                //gte
                else if (dataType == "integer" && condition == "gte")
                {
                    if (toValue == "")
                    {
                        fc.Add("CONVERT(int," + dataIndx + ")" + " >= @" + dataIndx);
                        param.Add(new SqlParameter(dataIndx, text));

                    }

                }
                //lte
                else if (dataType == "integer" && condition == "lte")
                {
                    if (toValue != "")
                    {
                        fc.Add("CONVERT(int," + dataIndx + ")" + " <= @" + dataIndx);
                        param.Add(new SqlParameter(dataIndx, text));

                    }

                }

                else if (dataType == "date" && condition == "=")
                {
                    fc.Add("CONVERT(datetime," + dataIndx + ")" + " = @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "contain")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text + "%"));
                }
                else if (condition == "notcontain")
                {
                    fc.Add(dataIndx + " not like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text + "%"));
                }
                else if (condition == "begin")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text + "%"));
                }
                else if (condition == "end")
                {
                    fc.Add(dataIndx + " like @" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, "%" + text));
                }
                else if (condition == "equal")
                {
                    fc.Add(dataIndx + "=@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "notequal")
                {
                    fc.Add(dataIndx + "!=@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "empty")
                {
                    fc.Add("isnull(" + dataIndx + ",'')=''");
                    //param.Add(new SqlParameter(dataIndx, text + "%"));
                }
                else if (condition == "notempty")
                {
                    fc.Add("isnull(" + dataIndx + ",'')!=''");
                }
                else if (condition == "less")
                {
                    fc.Add(dataIndx + "<@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
                else if (condition == "great")
                {
                    fc.Add(dataIndx + ">@" + dataIndx);
                    param.Add(new SqlParameter(dataIndx, text));
                }
            }
            String query = "";
            if (filters.Count > 0 && fc.Count > 0)
            {
                query = " and " + String.Join(" " + mode + " ", fc.ToArray());
            }

            deSerializedFilter ds = new deSerializedFilter();
            ds.query = query;
            ds.param = param;
            return ds;
        }
        //create in a static class
        static public string GetValObjDy(object obj, string propertyName)
        {
            return obj.GetType().GetProperty(propertyName).GetValue(obj, null).ToString();
        }
    }
}