using FBLT_Tax.Models;
using FBLT_Tax.Utility;
using Dapper;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace FBLT_Tax.Controllers
{
    public class ReportController : Controller
    {
        Helper _h = new Helper();

        //static string jsonData = System.IO.File.ReadAllText(@"C:\Users\thpratha\Documents\Project\jsconfig.json");
        static string jsonData = System.IO.File.ReadAllText(@"D:\ConnStrConfig\jsconfig.json");
        //RootConStr a = JsonConvert.DeserializeObject<IEnumerable<IDictionary<string, object>>>(jsonData);
        RootConStr[] jsConfigs = Newtonsoft.Json.JsonConvert.DeserializeObject<RootConStr[]>(jsonData);
        #region Report By Lease_Contract_For_Tax
        public ActionResult GetReport_By_Contract_For_Tax(String Period, String pq_filter, int pq_curPage, int pq_rPP)
        {
            String filterQuery = "";
            List<object> filterParam = new List<object>();
            if (pq_filter != null && pq_filter.Length > 0)
            {
                deSerializedFilter dsf = FilterHelper.deSerializeFilter2(pq_filter);
                filterQuery = dsf.query;
                filterParam = dsf.param;
            }
            #region FILTER
            //var filterCustomer_Id = "";
            //var filterCustomer_Name = "";
            //var filterCustomer_Name_Eng = "";
            //var filterBC = "";
            //var filterCLS_Team = "";
            //var filterAverage_Payment_Day = "";
            //var filterCredit_Rating = "";
            //var filterCredit_Limit = "";
            //var filterBilling_Type = "";
            //var filterCollection_Type = "";
            //var filterBilling_Placement_Date = "";
            //var filterCollection_Payment_Date = "";
            //var filterGroup_Customer = "";
            //var filterBusiness_Type = "";

            ////CUSTOMER_ID
            //var CUSTOMER_ID = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUST_ID").ToList();
            //if (CUSTOMER_ID.Count > 0)
            //    filterCustomer_Id = FilterHelper.GetValObjDy(CUSTOMER_ID[0], "Value");

            ////CUSTOMER_NAME
            //var CUSTOMER_NAME = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUSTOMER_NAME").ToList();
            //if (CUSTOMER_NAME.Count > 0)
            //    filterCustomer_Name = FilterHelper.GetValObjDy(CUSTOMER_NAME[0], "Value");

            ////Customer_Name_Eng
            //var Customer_Name_Eng = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUSTOMER_NAME_ENG").ToList();
            //if (Customer_Name_Eng.Count > 0)
            //    filterCustomer_Name_Eng = FilterHelper.GetValObjDy(Customer_Name_Eng[0], "Value");

            ////CLS_TEAM
            //var CLS_TEAM = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CLS_TEAM").ToList();
            //if (CLS_TEAM.Count > 0)
            //    filterCLS_Team = FilterHelper.GetValObjDy(CLS_TEAM[0], "Value");

            ////BC
            //var BC = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BC").ToList();
            //if (BC.Count > 0)
            //    filterBC = FilterHelper.GetValObjDy(BC[0], "Value");

            ////Average_Payment_Day
            //var Average_Payment_Day = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "AVERAGE_PAYMENT_DAY").ToList();
            //if (Average_Payment_Day.Count > 0)
            //    filterAverage_Payment_Day = FilterHelper.GetValObjDy(Average_Payment_Day[0], "Value");

            ////Credit_Rating
            //var Credit_Rating = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CREDIT_RATING").ToList();
            //if (Credit_Rating.Count > 0)
            //    filterCredit_Rating = FilterHelper.GetValObjDy(Credit_Rating[0], "Value");

            ////Credit_Limit
            //var Credit_Limit = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CREDIT_LIMIT").ToList();
            //if (Credit_Limit.Count > 0)
            //    filterCredit_Limit = FilterHelper.GetValObjDy(Credit_Limit[0], "Value");

            ////Billing_Type
            //var Billing_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BILLING_TYPE").ToList();
            //if (Billing_Type.Count > 0)
            //    filterBilling_Type = FilterHelper.GetValObjDy(Billing_Type[0], "Value");

            ////Collection_Type
            //var Collection_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "COLLECTION_TYPE").ToList();
            //if (Collection_Type.Count > 0)
            //    filterCollection_Type = FilterHelper.GetValObjDy(Collection_Type[0], "Value");

            ////Billing_Placement_Date
            //var Billing_Placement_Date = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BILLING_PLACEMENT_DATE").ToList();
            //if (Billing_Placement_Date.Count > 0)
            //    filterBilling_Placement_Date = FilterHelper.GetValObjDy(Billing_Placement_Date[0], "Value");

            ////Collection_Payment_Date
            //var Collection_Payment_Date = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "COLLECTION_PAYMENT_DATE").ToList();
            //if (Collection_Payment_Date.Count > 0)
            //    filterCollection_Payment_Date = FilterHelper.GetValObjDy(Collection_Payment_Date[0], "Value");

            ////Group_Customer	
            //var Group_Customer = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "GROUP_CUSTOMER").ToList();
            //if (Group_Customer.Count > 0)
            //    filterGroup_Customer = FilterHelper.GetValObjDy(Group_Customer[0], "Value");

            ////Business_Type
            //var Business_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BUSINESS_TYPE").ToList();
            //if (Business_Type.Count > 0)
            //    filterBusiness_Type = FilterHelper.GetValObjDy(Business_Type[0], "Value");



            #endregion
            using (var conn = new SqlConnection(jsConfigs[0].ConnStr))
            {
                conn.Open();
                var p = new DynamicParameters();
                p.Add("@Period", Period);
                //p.Add("@Customer_Id", filterCustomer_Id);
                //p.Add("@Customer_Name", filterCustomer_Name);
                //p.Add("@Customer_Name_Eng", filterCustomer_Name_Eng);
                //p.Add("@BC", filterBC);
                //p.Add("@CLS_Team", filterCLS_Team);
                //p.Add("@Average_Payment_Day", filterAverage_Payment_Day);
                //p.Add("@Credit_Rating", filterCredit_Rating);
                //p.Add("@Credit_Limit", filterCredit_Limit);
                //p.Add("@Billing_Type", filterBilling_Type);
                //p.Add("@Collection_Type", filterCollection_Type);
                //p.Add("@Billing_Placement_Date", filterBilling_Placement_Date);
                //p.Add("@Collection_Payment_Date", filterCollection_Payment_Date);
                //p.Add("@Group_Customer", filterGroup_Customer);
                //p.Add("@Business_Type", filterBusiness_Type);

                var data = conn.Query<Report_By_Contract_For_Tax>("[sp_FBLT_Tax_Report_Lease_Contract_For_Tax]", p, commandType: CommandType.StoredProcedure).ToList();

                int total_Records = data.Count();

                int skip = (pq_rPP * (pq_curPage - 1));
                if (skip >= total_Records)
                {
                    pq_curPage = (int)Math.Ceiling(((double)total_Records) / pq_rPP);
                    skip = (pq_rPP * (pq_curPage - 1));
                }

                var custResultFilter = (from custRow in data
                                        select custRow).Skip(skip).Take(pq_rPP);
                StringBuilder sb = new StringBuilder(@"{""totalRecords"":" + total_Records + @",""curPage"":" + pq_curPage + @",""data"":");
                JavaScriptSerializer js = new JavaScriptSerializer();
                js.MaxJsonLength = int.MaxValue;
                String json = js.Serialize(custResultFilter);
                sb.Append(json);
                sb.Append("}");
                return this.Content(sb.ToString(), "text/text");
            }
        }

        public ActionResult DownloadExcelReport_By_Contract_For_Tax(
            String FileName
            , String Period
            //, String Customer_Id
            //, String Customer_Name
            //, String Customer_Name_Eng
            //, String BC
            //, String CLS_Team
            //, String Average_Payment_Day
            //, String Credit_Rating
            //, String Credit_Limit
            //, String Billing_Type
            //, String Collection_Type
            //, String Billing_Placement_Date
            //, String Collection_Payment_Date
            //, String Group_Customer
            //, String Business_Type
            )
        {
            #region FILTER
            //var filterCustomer_Id = "";
            //var filterCustomer_Name = "";
            //var filterCustomer_Name_Eng = "";
            //var filterBC = "";
            //var filterCLS_Team = "";
            //var filterAverage_Payment_Day = "";
            //var filterCredit_Rating = "";
            //var filterCredit_Limit = "";
            //var filterBilling_Type = "";
            //var filterCollection_Type = "";
            //var filterBilling_Placement_Date = "";
            //var filterCollection_Payment_Date = "";
            //var filterGroup_Customer = "";
            //var filterBusiness_Type = "";


            //if (!string.IsNullOrEmpty(Customer_Id))
            //    filterCustomer_Id = '%' + Customer_Id + '%';
            //if (!string.IsNullOrEmpty(Customer_Name))
            //    filterCustomer_Name = '%' + Customer_Name + '%';
            //if (!string.IsNullOrEmpty(Customer_Name_Eng))
            //    filterCustomer_Name_Eng = '%' + Customer_Name_Eng + '%';
            //if (!string.IsNullOrEmpty(BC))
            //    filterBC = '%' + BC + '%';
            //if (!string.IsNullOrEmpty(CLS_Team))
            //    filterCLS_Team = '%' + CLS_Team + '%';
            //if (!string.IsNullOrEmpty(Average_Payment_Day))
            //    filterAverage_Payment_Day = '%' + Average_Payment_Day + '%';
            //if (!string.IsNullOrEmpty(Credit_Rating))
            //    filterCredit_Rating = '%' + Credit_Rating + '%';
            //if (!string.IsNullOrEmpty(Credit_Limit))
            //    filterCredit_Limit = '%' + Credit_Limit + '%';
            //if (!string.IsNullOrEmpty(Billing_Type))
            //    filterBilling_Type = '%' + Billing_Type + '%';
            //if (!string.IsNullOrEmpty(Collection_Type))
            //    filterCollection_Type = '%' + Collection_Type + '%';
            //if (!string.IsNullOrEmpty(Billing_Placement_Date))
            //    filterBilling_Placement_Date = '%' + Billing_Placement_Date + '%';
            //if (!string.IsNullOrEmpty(Collection_Payment_Date))
            //    filterCollection_Payment_Date = '%' + Collection_Payment_Date + '%';
            //if (!string.IsNullOrEmpty(Group_Customer))
            //    filterGroup_Customer = '%' + Group_Customer + '%';
            //if (!string.IsNullOrEmpty(Business_Type))
            //    filterBusiness_Type = '%' + Business_Type + '%';
            #endregion


            using (var conn = new SqlConnection(jsConfigs[0].ConnStr))
            {
                conn.Open();
                var p = new DynamicParameters();
                p.Add("@Period", Period);
                //p.Add("@Customer_Id", filterCustomer_Id);
                //p.Add("@Customer_Name", filterCustomer_Name);
                //p.Add("@Customer_Name_Eng", filterCustomer_Name_Eng);
                //p.Add("@BC", filterBC);
                //p.Add("@CLS_Team", filterCLS_Team);
                //p.Add("@Average_Payment_Day", filterAverage_Payment_Day);
                //p.Add("@Credit_Rating", filterCredit_Rating);
                //p.Add("@Credit_Limit", filterCredit_Limit);
                //p.Add("@Billing_Type", filterBilling_Type);
                //p.Add("@Collection_Type", filterCollection_Type);
                //p.Add("@Billing_Placement_Date", filterBilling_Placement_Date);
                //p.Add("@Collection_Payment_Date", filterCollection_Payment_Date);
                //p.Add("@Group_Customer", filterGroup_Customer);
                //p.Add("@Business_Type", filterBusiness_Type);

                var data = conn.Query<Report_By_Contract_For_Tax>("[sp_FBLT_Tax_Report_Lease_Contract_For_Tax]", p, commandType: CommandType.StoredProcedure).ToList();


                string json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(json);

                //Filter Column for export into excel "Customer_Code", "Customer_Name", "Case_ID", "Bank_Info", "Account_No", "Location", "Data_Date", "CLS_Team", "BC", "Status_Identify", "Status_Matching", "Status_Apply", "Status_Payment_Advice"
                DataView dvTemp = dt.DefaultView;
                DataTable dtTemp = new DataTable();
                if (data.Count() == 0)
                {
                    dtTemp = new DataTable();
                }
                else
                {
                    dtTemp = dvTemp.ToTable(false,
                     "Customer_Id"
                    ,"Contract_No"
                    ,"Reference"
                    ,"Lease_Type"
                    ,"Master_Contract"
                    ,"Contract_Name"
                    ,"Customer_Group"
                    ,"Start_Date"
                    ,"First_Payment_Due_Date"
                    , "Finish_Date"
                    ,"End_Date"
                    ,"Lease_Duration"
                    ,"Manual_Interest_Rate"
                    ,"Deposit_Amt"
                    ,"Down_Payment_Amt"
                    ,"Tradein_Amt"
                    ,"Funded_Amt"
                    ,"Invoice_Amt"
                    ,"Monthly_Installment"
                    ,"Residual_Value"
                    ,"Model"
                    ,"Serial_No"
                    ,"Ref_Invoice"
                    ,"Quantity"
                    ,"SR_Code"
                    ,"SR_Name"
                    ,"Mkt_Seg"
                    ,"Collector_Name"
                    ,"Status"
                    ,"PCR_Status"
                    ,"Booked_Date"
                    ,"Total_Quantity"
                    ,"Customer_Category"
                    ,"Billing_Condition"
                    ,"Payment_Condition"
                    ,"Rec_Contact_Date"
                    ,"Non_Serial_Desc"
                    ,"VAT_Type"
                    ,"WH_Tax"
                    ,"Start_Installment_Remark"
                    ,"Remark_General"
                    ,"RV_FXTH"
                    ,"Template_Form"
                    ,"Credit_Evaluate_Approved"
                    ,"chk_date"
                    ,"Installment_Billed_Per_Contract"
                    ,"Amount"
                    ,"Depreciation"
                    ,"Period_Of_Accumulated"
                    ,"AccumDepre"
                    ,"NBV"
                    ,"NBV_Only_Expired_Cancelled"
                    ,"Gain_Loss"
                    ,"Invoice_Values_BB"
                    ,"Total_BC"
                    ,"Invoice_Values_BD"
                    ,"Total_BE"
                    ,"Invoice_Values_BF"
                    ,"Total_BG"
                    ,"sum_Total" 
                    ,"End_Date_useful_life"
                    , "Useful_life"
                    , "EndDate_for_Cal"
                    , "Utilized_useful_life"
                    , "Accum_Useful_life");
                    dtTemp.Columns["Customer_Id"].ColumnName = "Cust. Id";
                    dtTemp.Columns["Contract_No"].ColumnName = "Contract No.";
                    dtTemp.Columns["Reference"].ColumnName = "Reference No";
                    dtTemp.Columns["Lease_Type"].ColumnName = "Contract Type";
                    dtTemp.Columns["Master_Contract"].ColumnName = "Master Contract No.";
                    dtTemp.Columns["Contract_Name"].ColumnName = "Contract Name";
                    dtTemp.Columns["Customer_Group"].ColumnName = "Customer Group";
                    dtTemp.Columns["Start_Date"].ColumnName = "Start Date";
                    dtTemp.Columns["First_Payment_Due_Date"].ColumnName = "First Payment Due Date";
                    dtTemp.Columns["Finish_Date"].ColumnName = "Finish Date";
                    dtTemp.Columns["End_Date"].ColumnName = "End Date(Expired)";
                    dtTemp.Columns["Lease_Duration"].ColumnName = "Lease Duration ";
                    dtTemp.Columns["Manual_Interest_Rate"].ColumnName = "Manual Interest Rate";
                    dtTemp.Columns["Deposit_Amt"].ColumnName = "Deposit Amt";
                    dtTemp.Columns["Down_Payment_Amt"].ColumnName = "Down Payment Amt";
                    dtTemp.Columns["Tradein_Amt"].ColumnName = "Tradein Amt";
                    dtTemp.Columns["Funded_Amt"].ColumnName = "Funded Amt";
                    dtTemp.Columns["Monthly_Installment"].ColumnName = "Monthly Installment";
                    dtTemp.Columns["Residual_Value"].ColumnName = "Residual Value";
                    dtTemp.Columns["Model"].ColumnName = "Model";
                    dtTemp.Columns["Serial_No"].ColumnName = "Serial";
                    dtTemp.Columns["Ref_Invoice"].ColumnName = "Ref Inv No.";
                    dtTemp.Columns["SR_Code"].ColumnName = "SR:Code";
                    dtTemp.Columns["SR_Name"].ColumnName = "SR:Name";
                    dtTemp.Columns["Mkt_Seg"].ColumnName = "Mkt. Seg";
                    dtTemp.Columns["Collector_Name"].ColumnName = "Collector Name";
                    dtTemp.Columns["Status"].ColumnName = "Status";
                    dtTemp.Columns["PCR_Status"].ColumnName = "PCR Status";
                    dtTemp.Columns["Booked_Date"].ColumnName = "Booked Date";
                    dtTemp.Columns["Total_Quantity"].ColumnName = "Total Quantity";
                    dtTemp.Columns["Customer_Category"].ColumnName = "Customer Category";
                    dtTemp.Columns["Billing_Condition"].ColumnName = "Billing Condition";
                    dtTemp.Columns["Payment_Condition"].ColumnName = "Payment Condition ";
                    dtTemp.Columns["Rec_Contact_Date"].ColumnName = "Rec.contact Date";
                    dtTemp.Columns["Non_Serial_Desc"].ColumnName = "Non Serial Desc";
                    dtTemp.Columns["VAT_Type"].ColumnName = "VAT Type";
                    dtTemp.Columns["WH_Tax"].ColumnName = "W/H Tax(%)";
                    dtTemp.Columns["Start_Installment_Remark"].ColumnName = "Start Installment remark";
                    dtTemp.Columns["Remark_General"].ColumnName = "Remark General";
                    dtTemp.Columns["RV_FXTH"].ColumnName = "R.V. to FXTH";
                    dtTemp.Columns["Template_Form"].ColumnName = "Template form";
                    dtTemp.Columns["Credit_Evaluate_Approved"].ColumnName = "Credit Evaluate Approved";

                    dtTemp.Columns["chk_date"].ColumnName = "*";
                    dtTemp.Columns["Installment_Billed_Per_Contract"].ColumnName = "Installment billed per contract";
                    dtTemp.Columns["Amount"].ColumnName = "Amount";
                    dtTemp.Columns["Depreciation"].ColumnName = "Depreciation";
                    dtTemp.Columns["Period_Of_Accumulated"].ColumnName = "Period of Accumulated";
                    dtTemp.Columns["AccumDepre"].ColumnName = "AccumDepre";
                    dtTemp.Columns["NBV"].ColumnName = "NBV";

                    dtTemp.Columns["NBV_Only_Expired_Cancelled"].ColumnName = "NBV(only expired, cancelled)";
                    dtTemp.Columns["Gain_Loss"].ColumnName = "-Gain/Loss (expired,Cancelled)";
                    dtTemp.Columns["Invoice_Values_BB"].ColumnName = "Tax inv. (Start 31,32,33,54,55,56)";
                    dtTemp.Columns["Total_BC"].ColumnName = "Amount(Exclude description 'RV')(Start 31,32,33,54,55,56 ไม่เอา RV)";
                    dtTemp.Columns["Invoice_Values_BD"].ColumnName = "Tax inv.(Start Inv.44,45,46,54,55,56)";
                    dtTemp.Columns["Total_BE"].ColumnName = "Residual Value (Other COG)(Only 'RV' Inv.44,45,46,54,55,56)";
                    dtTemp.Columns["Invoice_Values_BF"].ColumnName = "Credit note No.(Start CN no.37,38,21)";
                    dtTemp.Columns["Total_BG"].ColumnName = "CM/DM/INV.RV(Start CN no.37,38,21)";
                    dtTemp.Columns["sum_Total"].ColumnName = "Ref.L3";

                    dtTemp.Columns["End_Date_useful_life"].ColumnName = "End_Date_useful_life";
                    dtTemp.Columns["Useful_life"].ColumnName = "Useful_life";
                    dtTemp.Columns["EndDate_for_Cal"].ColumnName = "EndDate_for_Cal";
                    dtTemp.Columns["Utilized_useful_life"].ColumnName = "Utilized_useful_life";
                    dtTemp.Columns["Accum_Useful_life"].ColumnName = "Accum_Useful_life";
                }
                string fileResultFullPath;
                fileResultFullPath = CreateExportExcel(dtTemp);

                var resultByt = System.IO.File.ReadAllBytes(fileResultFullPath);
                System.IO.File.Delete(fileResultFullPath);
                return File(resultByt, "application/xlsx", FileName + ".xlsx");
            }
        }
        #endregion


        #region Report By Lease_Contract_Late_Invoice
        public ActionResult GetReport_By_Contract_Late_Invoice(String Period, String pq_filter, int pq_curPage, int pq_rPP)
        {
            String filterQuery = "";
            List<object> filterParam = new List<object>();
            if (pq_filter != null && pq_filter.Length > 0)
            {
                deSerializedFilter dsf = FilterHelper.deSerializeFilter2(pq_filter);
                filterQuery = dsf.query;
                filterParam = dsf.param;
            }
            #region FILTER
            //var filterCustomer_Id = "";
            //var filterCustomer_Name = "";
            //var filterCustomer_Name_Eng = "";
            //var filterBC = "";
            //var filterCLS_Team = "";
            //var filterAverage_Payment_Day = "";
            //var filterCredit_Rating = "";
            //var filterCredit_Limit = "";
            //var filterBilling_Type = "";
            //var filterCollection_Type = "";
            //var filterBilling_Placement_Date = "";
            //var filterCollection_Payment_Date = "";
            //var filterGroup_Customer = "";
            //var filterBusiness_Type = "";

            ////CUSTOMER_ID
            //var CUSTOMER_ID = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUST_ID").ToList();
            //if (CUSTOMER_ID.Count > 0)
            //    filterCustomer_Id = FilterHelper.GetValObjDy(CUSTOMER_ID[0], "Value");

            ////CUSTOMER_NAME
            //var CUSTOMER_NAME = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUSTOMER_NAME").ToList();
            //if (CUSTOMER_NAME.Count > 0)
            //    filterCustomer_Name = FilterHelper.GetValObjDy(CUSTOMER_NAME[0], "Value");

            ////Customer_Name_Eng
            //var Customer_Name_Eng = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CUSTOMER_NAME_ENG").ToList();
            //if (Customer_Name_Eng.Count > 0)
            //    filterCustomer_Name_Eng = FilterHelper.GetValObjDy(Customer_Name_Eng[0], "Value");

            ////CLS_TEAM
            //var CLS_TEAM = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CLS_TEAM").ToList();
            //if (CLS_TEAM.Count > 0)
            //    filterCLS_Team = FilterHelper.GetValObjDy(CLS_TEAM[0], "Value");

            ////BC
            //var BC = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BC").ToList();
            //if (BC.Count > 0)
            //    filterBC = FilterHelper.GetValObjDy(BC[0], "Value");

            ////Average_Payment_Day
            //var Average_Payment_Day = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "AVERAGE_PAYMENT_DAY").ToList();
            //if (Average_Payment_Day.Count > 0)
            //    filterAverage_Payment_Day = FilterHelper.GetValObjDy(Average_Payment_Day[0], "Value");

            ////Credit_Rating
            //var Credit_Rating = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CREDIT_RATING").ToList();
            //if (Credit_Rating.Count > 0)
            //    filterCredit_Rating = FilterHelper.GetValObjDy(Credit_Rating[0], "Value");

            ////Credit_Limit
            //var Credit_Limit = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "CREDIT_LIMIT").ToList();
            //if (Credit_Limit.Count > 0)
            //    filterCredit_Limit = FilterHelper.GetValObjDy(Credit_Limit[0], "Value");

            ////Billing_Type
            //var Billing_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BILLING_TYPE").ToList();
            //if (Billing_Type.Count > 0)
            //    filterBilling_Type = FilterHelper.GetValObjDy(Billing_Type[0], "Value");

            ////Collection_Type
            //var Collection_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "COLLECTION_TYPE").ToList();
            //if (Collection_Type.Count > 0)
            //    filterCollection_Type = FilterHelper.GetValObjDy(Collection_Type[0], "Value");

            ////Billing_Placement_Date
            //var Billing_Placement_Date = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BILLING_PLACEMENT_DATE").ToList();
            //if (Billing_Placement_Date.Count > 0)
            //    filterBilling_Placement_Date = FilterHelper.GetValObjDy(Billing_Placement_Date[0], "Value");

            ////Collection_Payment_Date
            //var Collection_Payment_Date = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "COLLECTION_PAYMENT_DATE").ToList();
            //if (Collection_Payment_Date.Count > 0)
            //    filterCollection_Payment_Date = FilterHelper.GetValObjDy(Collection_Payment_Date[0], "Value");

            ////Group_Customer	
            //var Group_Customer = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "GROUP_CUSTOMER").ToList();
            //if (Group_Customer.Count > 0)
            //    filterGroup_Customer = FilterHelper.GetValObjDy(Group_Customer[0], "Value");

            ////Business_Type
            //var Business_Type = filterParam.Where(p => FilterHelper.GetValObjDy(p, "ParameterName").ToUpper() == "BUSINESS_TYPE").ToList();
            //if (Business_Type.Count > 0)
            //    filterBusiness_Type = FilterHelper.GetValObjDy(Business_Type[0], "Value");



            #endregion
            using (var conn = new SqlConnection(jsConfigs[0].ConnStr))
            {
                conn.Open();
                var p = new DynamicParameters();
                p.Add("@Period", Period);
                //p.Add("@Customer_Id", filterCustomer_Id);
                //p.Add("@Customer_Name", filterCustomer_Name);
                //p.Add("@Customer_Name_Eng", filterCustomer_Name_Eng);
                //p.Add("@BC", filterBC);
                //p.Add("@CLS_Team", filterCLS_Team);
                //p.Add("@Average_Payment_Day", filterAverage_Payment_Day);
                //p.Add("@Credit_Rating", filterCredit_Rating);
                //p.Add("@Credit_Limit", filterCredit_Limit);
                //p.Add("@Billing_Type", filterBilling_Type);
                //p.Add("@Collection_Type", filterCollection_Type);
                //p.Add("@Billing_Placement_Date", filterBilling_Placement_Date);
                //p.Add("@Collection_Payment_Date", filterCollection_Payment_Date);
                //p.Add("@Group_Customer", filterGroup_Customer);
                //p.Add("@Business_Type", filterBusiness_Type);

                var data = conn.Query<Report_By_Contract_For_Tax>("[sp_FBLT_Tax_Report_Lease_Contract_Late_Invoice]", p, commandType: CommandType.StoredProcedure).ToList();

                int total_Records = data.Count();

                int skip = (pq_rPP * (pq_curPage - 1));
                if (skip >= total_Records)
                {
                    pq_curPage = (int)Math.Ceiling(((double)total_Records) / pq_rPP);
                    skip = (pq_rPP * (pq_curPage - 1));
                }

                var custResultFilter = (from custRow in data
                                        select custRow).Skip(skip).Take(pq_rPP);
                StringBuilder sb = new StringBuilder(@"{""totalRecords"":" + total_Records + @",""curPage"":" + pq_curPage + @",""data"":");
                JavaScriptSerializer js = new JavaScriptSerializer();
                js.MaxJsonLength = int.MaxValue;
                String json = js.Serialize(custResultFilter);
                sb.Append(json);
                sb.Append("}");
                return this.Content(sb.ToString(), "text/text");
            }
        }

        public ActionResult DownloadExcelReport_By_Contract_Late_Invoice(
            String FileName
            , String Period
            //, String Customer_Id
            //, String Customer_Name
            //, String Customer_Name_Eng
            //, String BC
            //, String CLS_Team
            //, String Average_Payment_Day
            //, String Credit_Rating
            //, String Credit_Limit
            //, String Billing_Type
            //, String Collection_Type
            //, String Billing_Placement_Date
            //, String Collection_Payment_Date
            //, String Group_Customer
            //, String Business_Type
            )
        {
            #region FILTER
            //var filterCustomer_Id = "";
            //var filterCustomer_Name = "";
            //var filterCustomer_Name_Eng = "";
            //var filterBC = "";
            //var filterCLS_Team = "";
            //var filterAverage_Payment_Day = "";
            //var filterCredit_Rating = "";
            //var filterCredit_Limit = "";
            //var filterBilling_Type = "";
            //var filterCollection_Type = "";
            //var filterBilling_Placement_Date = "";
            //var filterCollection_Payment_Date = "";
            //var filterGroup_Customer = "";
            //var filterBusiness_Type = "";


            //if (!string.IsNullOrEmpty(Customer_Id))
            //    filterCustomer_Id = '%' + Customer_Id + '%';
            //if (!string.IsNullOrEmpty(Customer_Name))
            //    filterCustomer_Name = '%' + Customer_Name + '%';
            //if (!string.IsNullOrEmpty(Customer_Name_Eng))
            //    filterCustomer_Name_Eng = '%' + Customer_Name_Eng + '%';
            //if (!string.IsNullOrEmpty(BC))
            //    filterBC = '%' + BC + '%';
            //if (!string.IsNullOrEmpty(CLS_Team))
            //    filterCLS_Team = '%' + CLS_Team + '%';
            //if (!string.IsNullOrEmpty(Average_Payment_Day))
            //    filterAverage_Payment_Day = '%' + Average_Payment_Day + '%';
            //if (!string.IsNullOrEmpty(Credit_Rating))
            //    filterCredit_Rating = '%' + Credit_Rating + '%';
            //if (!string.IsNullOrEmpty(Credit_Limit))
            //    filterCredit_Limit = '%' + Credit_Limit + '%';
            //if (!string.IsNullOrEmpty(Billing_Type))
            //    filterBilling_Type = '%' + Billing_Type + '%';
            //if (!string.IsNullOrEmpty(Collection_Type))
            //    filterCollection_Type = '%' + Collection_Type + '%';
            //if (!string.IsNullOrEmpty(Billing_Placement_Date))
            //    filterBilling_Placement_Date = '%' + Billing_Placement_Date + '%';
            //if (!string.IsNullOrEmpty(Collection_Payment_Date))
            //    filterCollection_Payment_Date = '%' + Collection_Payment_Date + '%';
            //if (!string.IsNullOrEmpty(Group_Customer))
            //    filterGroup_Customer = '%' + Group_Customer + '%';
            //if (!string.IsNullOrEmpty(Business_Type))
            //    filterBusiness_Type = '%' + Business_Type + '%';
            #endregion


            using (var conn = new SqlConnection(jsConfigs[0].ConnStr))
            {
                conn.Open();
                var p = new DynamicParameters();
                p.Add("@Period", Period);
                //p.Add("@Customer_Id", filterCustomer_Id);
                //p.Add("@Customer_Name", filterCustomer_Name);
                //p.Add("@Customer_Name_Eng", filterCustomer_Name_Eng);
                //p.Add("@BC", filterBC);
                //p.Add("@CLS_Team", filterCLS_Team);
                //p.Add("@Average_Payment_Day", filterAverage_Payment_Day);
                //p.Add("@Credit_Rating", filterCredit_Rating);
                //p.Add("@Credit_Limit", filterCredit_Limit);
                //p.Add("@Billing_Type", filterBilling_Type);
                //p.Add("@Collection_Type", filterCollection_Type);
                //p.Add("@Billing_Placement_Date", filterBilling_Placement_Date);
                //p.Add("@Collection_Payment_Date", filterCollection_Payment_Date);
                //p.Add("@Group_Customer", filterGroup_Customer);
                //p.Add("@Business_Type", filterBusiness_Type);

                var data = conn.Query<Report_By_Contract_For_Tax>("[sp_FBLT_Tax_Report_Lease_Contract_Late_Invoice]", p, commandType: CommandType.StoredProcedure).ToList();


                string json = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(json);

                //Filter Column for export into excel "Customer_Code", "Customer_Name", "Case_ID", "Bank_Info", "Account_No", "Location", "Data_Date", "CLS_Team", "BC", "Status_Identify", "Status_Matching", "Status_Apply", "Status_Payment_Advice"
                DataView dvTemp = dt.DefaultView;
                DataTable dtTemp = new DataTable();
                if (data.Count() == 0)
                {
                    dtTemp = new DataTable();
                }
                else
                {
                    dtTemp = dvTemp.ToTable(false,
                     "Customer_Id"
                    , "Contract_No"
                    , "Reference"
                    //, "Lease_Type"
                    //, "Master_Contract"
                    , "Contract_Name"
                    , "Customer_Group"
                    , "Start_Date"
                    //, "First_Payment_Due_Date"
                    , "Finish_Date"
                    , "End_Date"
                    , "Lease_Duration"
                    , "Manual_Interest_Rate"
                    , "Deposit_Amt"
                    , "Down_Payment_Amt"
                    , "Tradein_Amt"
                    , "Funded_Amt"
                    , "Invoice_Amt"
                    , "Monthly_Installment"
                    , "Residual_Value"
                    //, "Invoice_Number"
                    , "Invoice_Date"
                    //, "Amount_AR508"
                    , "Status"
                    , "Invoice_Values_BB"
                    , "Total_BC"
                    , "Invoice_Values_BD"
                    , "Total_BE"
                    , "Invoice_Values_BF"
                    , "Total_BG");
                    dtTemp.Columns["Customer_Id"].ColumnName = "Cust. Id";
                    dtTemp.Columns["Contract_No"].ColumnName = "Contract No.";
                    dtTemp.Columns["Reference"].ColumnName = "Reference No";
                    dtTemp.Columns["Contract_Name"].ColumnName = "Contract Name";
                    dtTemp.Columns["Customer_Group"].ColumnName = "Customer Group";
                    dtTemp.Columns["Start_Date"].ColumnName = "Start Date";
                    dtTemp.Columns["Finish_Date"].ColumnName = "Finish Date";
                    dtTemp.Columns["End_Date"].ColumnName = "End Date(Expired)";
                    dtTemp.Columns["Lease_Duration"].ColumnName = "Lease Duration";
                    dtTemp.Columns["Manual_Interest_Rate"].ColumnName = "Manual Interest Rate";
                    dtTemp.Columns["Deposit_Amt"].ColumnName = "Deposit Amt";
                    dtTemp.Columns["Down_Payment_Amt"].ColumnName = "Down Payment Amt";
                    dtTemp.Columns["Tradein_Amt"].ColumnName = "Tradein Amt";
                    dtTemp.Columns["Funded_Amt"].ColumnName = "Funded Amt";
                    dtTemp.Columns["Monthly_Installment"].ColumnName = "Monthly Installment";
                    dtTemp.Columns["Residual_Value"].ColumnName = "Residual Value";
                    //dtTemp.Columns["Invoice_Number"].ColumnName = "Invoice Number";
                    dtTemp.Columns["Invoice_Date"].ColumnName = "Invoice Date";
                    //dtTemp.Columns["Amount_AR508"].ColumnName = "Amount";
                    dtTemp.Columns["Status"].ColumnName = "Status";
                    dtTemp.Columns["Invoice_Values_BB"].ColumnName = "Tax inv. (Start 31,32,33,54,55,56)";
                    dtTemp.Columns["Total_BC"].ColumnName = "Amount(Exclude description 'RV')(Start 31,32,33,54,55,56 ไม่เอา RV)";
                    dtTemp.Columns["Invoice_Values_BD"].ColumnName = "Tax inv.(Start Inv.44,45,46,54,55,56)";
                    dtTemp.Columns["Total_BE"].ColumnName = "Residual Value (Other COG)(Only 'RV' Inv.44,45,46,54,55,56)";
                    dtTemp.Columns["Invoice_Values_BF"].ColumnName = "Credit note No.(Start CN no.37,38,21)";
                    dtTemp.Columns["Total_BG"].ColumnName = "CM/DM/INV.RV(Start CN no.37,38,21)";
                }
                string fileResultFullPath;
                fileResultFullPath = CreateExportExcel(dtTemp);

                var resultByt = System.IO.File.ReadAllBytes(fileResultFullPath);
                System.IO.File.Delete(fileResultFullPath);
                return File(resultByt, "application/xlsx", FileName + ".xlsx");
            }
        }
        #endregion
        #region Get DDL
        public ActionResult GetDDLOptions(String Module, String Filter1, String Filter2)
        {
            //StringBuilder whereSql = new StringBuilder();
            //whereSql.Append(" AND OptionsText IS NOT NULL ");
            //StringBuilder orderSql = new StringBuilder();
            //orderSql.Append(" Order by Data_Date DESC ");
            using (var conn = new SqlConnection(jsConfigs[0].ConnStr))
            {
                conn.Open();
                var p = new DynamicParameters();
                p.Add("@Module", Module);
                p.Add("@ValueFilter1", Filter1);
                p.Add("@ValueFilter2", Filter2);
                var data = conn.Query<DDLOptionsEntities>("[sp_FBLT_Tax_GetDDLList]", p, commandType: CommandType.StoredProcedure).ToList();

                int total_Records = data.Count();

                StringBuilder sb = new StringBuilder();
                JavaScriptSerializer js = new JavaScriptSerializer();
                js.MaxJsonLength = int.MaxValue;
                String json = js.Serialize(data);
                sb.Append(json);

                return this.Content(sb.ToString(), "text/text");
            }
        }
        #endregion
        #region CreateExportExcel
        private String CreateExportExcel(DataTable dt)
        {

            String fileResultFullPath = "";
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "RP4";
                excelPackage.Workbook.Properties.Title = "Result Export File";
                excelPackage.Workbook.Properties.Subject = "Export for AR Report Data file";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                int colNumber = 1;
                //Header 
                foreach (DataColumn dc in dt.Columns)
                {
                    worksheet.Cells[1, colNumber].Value = dc.ColumnName;

                    int rowNumber = 2;
                    //Detail
                    foreach (DataRow dr in dt.Rows)
                    {
                        worksheet.Cells[rowNumber, colNumber].Value = dr[colNumber - 1].ToString();
                        rowNumber = rowNumber + 1;
                    }

                    colNumber = colNumber + 1;
                }

                //Save your file
                String genDateTime = DateTime.Now.ToString("yyyyMMddHHmmssFFF");
                String fileNameSaveExcel = String.Format("{0}_{1}.xlsx", "exportTemp", genDateTime);
                fileResultFullPath = Server.MapPath(String.Format("../Export/{0}", fileNameSaveExcel));
                FileInfo fi = new FileInfo(fileResultFullPath);
                excelPackage.SaveAs(fi);
            }

            return fileResultFullPath;

        }
        #endregion
    }
}
