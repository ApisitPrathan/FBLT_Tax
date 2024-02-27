using System;
using System.Linq;
using System.Web;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace FBLT_Tax.Utility
{
	public class genExcelByTemplate
	{
		public string genByDataTable(string TemplateFileName, DataTable dt, int rowStart, int colStart, string sUserName, string sRef_ID, int refRow_pCustID, int refCol_pCustID, string pCustID, int refRow_pCustName, int refCol_pCustName, string pCustName, int refRow_pCustNameENG, int refCol_pCustNameENG, string pCustNameENG)
		{
			Boolean bErr = false;
			CultureInfo provider = CultureInfo.InvariantCulture;
			string sFileName = "";
			//TemplateFileName = @"C:\Users\thpratha\Documents\Project\Customer Payment Performance\Document\Template_Report_Summary.xlsx";
			sFileName = TemplateFileName.Replace(".xlsx", "_" + sUserName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssFFF") + ".xlsx").Replace("Excel Template\\", "");

			try
			{
				//System.IO.File.Copy(TemplateFileName, sFileName, true);
				System.IO.File.Copy(TemplateFileName, sFileName, true);
				System.IO.FileInfo newFile = new System.IO.FileInfo(sFileName);

				OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
				OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile);
				OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets["Main"];

				//set border
				//var modelRows = exportQuery.Count() + 1;
				//string modelRange = "D1:F" + modelRows.ToString();
				//var modelTable = worksheet.Cells[modelRange];

				int rowNumber = rowStart;
				int colNumber = colStart;

				if (pCustID != "")
				{
					//DateTime date = DateTime.ParseExact(pDate, "d/M/yyyy", CultureInfo.InvariantCulture);
					ws.Cells[refRow_pCustID, refCol_pCustID].Value = pCustID;
					//font
					ws.Cells[refRow_pCustID, refCol_pCustID].Style.Font.Name = "CordiaUPC";
					ws.Cells[refRow_pCustID, refCol_pCustID].Style.Font.Size = 12;
				}
				if (pCustName != "")
				{
					ws.Cells[refRow_pCustName, refCol_pCustName].Value = pCustName;
					//font
					ws.Cells[refRow_pCustName, refCol_pCustName].Style.Font.Name = "CordiaUPC";
					ws.Cells[refRow_pCustName, refCol_pCustName].Style.Font.Size = 12;
				}
				if (pCustNameENG != "")
				{
					ws.Cells[refRow_pCustNameENG, refCol_pCustNameENG].Value = pCustNameENG;
					//font
					ws.Cells[refRow_pCustNameENG, refCol_pCustNameENG].Style.Font.Name = "CordiaUPC";
					ws.Cells[refRow_pCustNameENG, refCol_pCustNameENG].Style.Font.Size = 12;
				}
				foreach (DataColumn dc in dt.Columns)
				{
					rowNumber = rowStart;
					foreach (DataRow dr in dt.Rows)
					{
						string type = dc.DataType.Name.ToLower().ToString();

						switch (type)
						{
							case "string":
								if (dc.ColumnName.IndexOf("Date") > -1)
								{
									try
									{
										if (dr[dc.ColumnName].ToString() != "")
											ws.Cells[rowNumber, colNumber].Value = DateTime.Parse(dr[dc.ColumnName].ToString()); break;
									}
									catch
									{
										ws.Cells[rowNumber, colNumber].Value = dr[dc.ColumnName].ToString();
									}
								}
								else if (dc.ColumnName == "InstallationTargetMonth")
								{
									try
									{
										if (dr[dc.ColumnName].ToString() != "")
											ws.Cells[rowNumber, colNumber].Value = DateTime.ParseExact("01/" + dr[dc.ColumnName].ToString(), "dd/MM/yyyy", provider); break;
									}
									catch
									{
										ws.Cells[rowNumber, colNumber].Value = dr[dc.ColumnName].ToString();
									}
								}
								else
									ws.Cells[rowNumber, colNumber].Value = dr[dc.ColumnName].ToString();

								break;
							case "double": if (dr[dc.ColumnName].ToString() != "") ws.Cells[rowNumber, colNumber].Value = double.Parse(dr[dc.ColumnName].ToString()); break;
							case "int64": if (dr[dc.ColumnName].ToString() != "") ws.Cells[rowNumber, colNumber].Value = int.Parse(dr[dc.ColumnName].ToString()); break;
							case "datetime": if (dr[dc.ColumnName].ToString() != "") ws.Cells[rowNumber, colNumber].Value = DateTime.Parse(dr[dc.ColumnName].ToString()); break;
							default:
								ws.Cells[rowNumber, colNumber].Value = dr[dc.ColumnName].ToString(); break;
						}

						// Set border
						string modelRange = "A5:G" + rowNumber.ToString();
						var modelTable = ws.Cells[modelRange];
						// Assign borders
						modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
						modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
						modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
						modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
						//font
						modelTable.Style.Font.Name = "CordiaUPC";
						modelTable.Style.Font.Size = 12;
						//if (colNumber == 3)
						//{//set Customer name disable wrap text
						//	ws.Cells[rowNumber, 3].Style.WrapText = false;
						//	ws.Cells[rowNumber, 3].Style.Numberformat.Format = "General";
						//	//modelTable.Style.Numberformat.Format = "General";
						//}
						//else if (colNumber == 5)
						//{
						//	ws.Cells[rowNumber, 5].Style.WrapText = true;
						//	ws.Cells[rowNumber, 5].Style.Numberformat.Format = "0";
						//	//modelTable.Style.Numberformat.Format = "0";
						//}
						//else if (colNumber == 11)
						//{
						//	ws.Cells[rowNumber, 11].Style.WrapText = true;
						//	ws.Cells[rowNumber, 11].Style.Numberformat.Format = "0";
						//}
						//else
						//{
						//	ws.Cells[rowNumber, colNumber].Style.WrapText = true;
						//	ws.Cells[rowNumber, colNumber].Style.Numberformat.Format = "General";
						//	//modelTable.Style.Numberformat.Format = "General";
						//}
						//modelTable.Style.WrapText = true;
						//set IgnoredErrors convert text
						var ie = ws.IgnoredErrors.Add(ws.Cells[modelRange]);
						ie.NumberStoredAsText = true;


						rowNumber++;
					}
					colNumber++;
				}
				System.IO.FileInfo newFile1 = new System.IO.FileInfo(sFileName);
				pck.SaveAs(newFile1);
			}
			catch (Exception ex)
			{
				bErr = true;

			}

			if (bErr == true)
				return "";
			return sFileName;
		}
	}
}