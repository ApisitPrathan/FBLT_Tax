using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FBLT_Tax.Models
{
    public class Report_By_Contract_For_Tax
	{
        public string Customer_Id { get; set; }
        public string Contract_No { get; set; }
        public string Reference { get; set; }
        public string Lease_Type { get; set; }
		public string Master_Contract { get; set; }
		public string Contract_Name { get; set; }
		public string Customer_Group { get; set; }
		public string Start_Date { get; set; }
		public string First_Payment_Due_Date { get; set; }
		public string Finish_Date { get; set; }
		public string End_Date { get; set; }
		public string Lease_Duration { get; set; }
		public string Manual_Interest_Rate { get; set; }
		public string Deposit_Amt { get; set; }
		public string Down_Payment_Amt { get; set; }
		public string Tradein_Amt { get; set; }
		public string Funded_Amt { get; set; }
		public string Invoice_Amt { get; set; }
		public string Monthly_Installment { get; set; }
		public string Residual_Value { get; set; }
		public string Model { get; set; }
		public string Serial_No { get; set; }
		public string Ref_Invoice { get; set; }
		public string Quantity { get; set; }
		public string SR_Code { get; set; }
		public string SR_Name { get; set; }
		public string Mkt_Seg { get; set; }
		public string Collector_Name { get; set; }
		public string Status { get; set; }
		public string PCR_Status { get; set; }
		public string Booked_Date { get; set; }
		public string Total_Quantity { get; set; }
		public string Customer_Category { get; set; }
		public string Billing_Condition { get; set; }
		public string Payment_Condition { get; set; }
		public string Rec_Contact_Date { get; set; }
		public string Non_Serial_Desc { get; set; }
		public string VAT_Type { get; set; }
		public string WH_Tax { get; set; }
		public string Start_Installment_Remark { get; set; }
		public string Remark_General { get; set; }
		public string RV_FXTH { get; set; }
		public string Template_Form { get; set; }
		public string Credit_Evaluate_Approved { get; set; }
		public string chk_date { get; set; }
		public string Installment_Billed_Per_Contract { get; set; }
		public string Amount { get; set; }
		public string Depreciation { get; set; }
		public string Period_Of_Accumulated { get; set; }
		public string AccumDepre { get; set; }
		public string NBV { get; set; }
		public string NBV_Only_Expired_Cancelled { get; set; }
		public string Gain_Loss { get; set; }
		public string Invoice_Values_BB { get; set; }
		public string Total_BC { get; set; }
		public string Invoice_Values_BD { get; set; }
		public string Total_BE { get; set; }
		public string Invoice_Values_BF { get; set; }
		public string Total_BG { get; set; }
		public string sum_Total { get; set; }
		public string Invoice_Number { get; set; }
		public string Invoice_Date { get; set; }
		public string Amount_AR508 { get; set; }

		public string End_Date_useful_life { get; set; }
		public string Useful_life { get; set; }
		public string EndDate_for_Cal { get; set; }
		public string Utilized_useful_life { get; set; }
		public string Accum_Useful_life { get; set; }

	}

	public class DDLOptionsEntities
	{
		public string OptionsText { get; set; }
		public string OptionsValue { get; set; }
	}
}