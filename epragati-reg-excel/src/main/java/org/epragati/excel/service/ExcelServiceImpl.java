package org.epragati.excel.service;

import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * @author krishnarjun.pampana
 *
 */
public class ExcelServiceImpl implements ExcelService {

	private static final Logger logger = LoggerFactory.getLogger(ExcelServiceImpl.class);

	@Override
	public void setHeaders(List<String> headers, String key) {

		switch (key) {
		case "Payments":

			headers.add("Date");
			headers.add("Application Number");
			headers.add("Transaction Type");
			headers.add("Payment Mode");
			headers.add("Office Code");
			headers.add("Application Fee");
			headers.add("Service Fee");
			headers.add("Test Fee");
			headers.add("Postal Fee");
			headers.add("Card Fee");
			headers.add("Late Fee");
			break;
		case "regPayments":
			headers.add("transaction Date");
			headers.add("office Code");
			headers.add("referance Number");
			headers.add("transaction Number");
			headers.add("serviceId");
			headers.add("trApplication Fee");
			headers.add("trService Fee");
			headers.add("prApplication Fee");
			headers.add("prService Fee");
			headers.add("prPostal Fee");
			headers.add("prCard Fee");
			headers.add("hpaApplication Fee");
			headers.add("tax Type");
			headers.add("tax Amount");
			headers.add("cess Fee");
			headers.add("bid Amount");
			headers.add("fcApplication Fee");
			headers.add("fcService Fee");
			headers.add("hsrp Fee");
			break;
		case "hsrpDetails":
			headers.add("Auth_refNo");
			headers.add("RTO CODE");
			headers.add("RTO NAME");
			headers.add("affixationCenterCode");
			headers.add("transactionNo");
			headers.add("transactionDate");
			headers.add("authorizationDate");
			headers.add("engineNo");
			headers.add("chassisNo");
			headers.add("prNumber");
			headers.add("ownerName");
			headers.add("ownerAddress");
			headers.add("owneremailid");
			headers.add("ownerPinCode");
			headers.add("mobileNo");
			headers.add("vehicleType");
			headers.add("transType");
			headers.add("vehicleClassType");
			headers.add("mfrsName");
			headers.add("modelName");
			headers.add("hsrpFee");
			headers.add("oldNewFlag");
			headers.add("govtVehicleTag");
			headers.add("timeStamp");
			headers.add("trNumber");
			headers.add("dealerName");
			headers.add("dealerMail");
			headers.add("dealerRtoCode");
			headers.add("regDate");
			headers.add("message");
			break;

		case "EmsReport":
			headers.add("Serial Number");
			headers.add("Application Number");
			headers.add("PR Number");
			headers.add("Office Code");
			headers.add("UserName");
			headers.add("EMS Number");
			headers.add("Posted Date");
			headers.add("Dispatched By");
			headers.add("Mobile Number");
			headers.add("PinCode");
			headers.add("Remark");
			break;
		case "paymentReport":
			headers.add("transaction Date");
			headers.add("transaction Number");
			headers.add("gateWayType");
			headers.add("officeCode");
			/*
			 * headers.add("Service Fee"); headers.add("Application Fee");
			 * headers.add("Card Fee"); headers.add("Postal Fee"); headers.add("Life Tax");
			 * headers.add("Cess Fee"); headers.add("Quarterly Tax");
			 * headers.add("Halfyearly Tax"); headers.add("Yearly Tax");
			 * headers.add("Fitness Service Fee"); headers.add("Fitness Fee");
			 * headers.add("Authorization"); headers.add("penalty");
			 * headers.add("penalty Arrears"); headers.add("Tax Arrears");
			 * headers.add("Tax service fee"); headers.add("Fc Late Fee");
			 * headers.add("Green Tax Fee"); headers.add("Late Fee");
			 */
			break;
		case "districtReport":
			headers.add("District Name");
			headers.add("total");
			break;
		case "AutoApprovalReport":
			headers.add("S.No");
			headers.add("Application Number");
			headers.add("PR Number / TR Number");
			headers.add("Service  Type");
			headers.add("File Pending From");
			headers.add("Approved Date");
			headers.add("Application Pending At ");
			headers.add("Status");
			break;
		case "ShowCauseReport":
			headers.add("S.No");
			headers.add("Class of Vehicle");
			headers.add("Non Payment Count");
			headers.add("No of Show Cause Issued Under Section 55");
			headers.add("No of Show Cause Issued For Non Payment for More than 5 Quarters");
			headers.add("No of Show Cause Issued under Rule 12 A");
			headers.add("No of Show Cause Issued under Rule 6");
			headers.add("No of Show Cause Issued under Section 7");
			headers.add("No of Show Cause Issued");
			headers.add("No of Registrations Cancelled");
			headers.add("Vehicles Which Paid Tax After Issue of Show Cause");
			headers.add("Total Amount Collected");
			break;
		case "VcrReport":
			headers.add("S.No");
			headers.add("Reg/TR/Chasis No");
			headers.add("VCR Number");
			headers.add("Class of Vehicle");
			headers.add("Challan No.");
			headers.add("Booked Date");
			headers.add("Challan Date");
			headers.add("Action Taken");
			headers.add("Paid Date");
			headers.add("MVI Name");
			headers.add("Receipt Number");
			headers.add("VCR Status");
			headers.add("Compound Fee");
			headers.add("Service Fee");
			break;
		case "VcrReportListMVI":
			headers.add("S.No");
			headers.add("MVI Name");
			headers.add("Total VCR Count");
			headers.add("Total Paid CompoundFee");
			headers.add("Total UnPaid CompoundFee");
			headers.add("Total tax");
			headers.add("Total Penalty");
			headers.add("Total Tax Arrears");
			headers.add("Total Penalty Arrears");
			headers.add("Total");
			break;
		case "NonPaymentDistrictWiseList":
			headers.add("S.No");
			headers.add("District Names");
			headers.add("Non Payment Count");
			break;
		case "NonPaymentOfficeWiseList":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Office Code");
			headers.add("Count");
			break;
		case "NonPaymentCovMandalWiseList":
			headers.add("S.No");
			headers.add("Mandal Name");
			headers.add("Count");
			break;
		case "RoadSafetyVcrDistrictList":	
			headers.add("S.No");
			headers.add("District");
			headers.add("Carrying extra persons in goods vehicle(per Head)");
			headers.add("Driving at Excessive Speed");
			headers.add("No Reflectors");
			headers.add("Non Wearing of helmets");
			headers.add("Non wearing of Seat belts");
			headers.add("Over load of goods vehicles");
			headers.add("Over loading Passengers");
			headers.add("Vehicle plying in Wrong Direction");
			headers.add("Total");
			break;	
		case "EodCountList":	
			headers.add("S.No");
			headers.add("Service Type");
			headers.add("Total Count");
			headers.add("Approve Count");
			headers.add("Reject Count");
			break;	
		case "EodDataList":	
			headers.add("S.No");
			headers.add("Application Number");
			headers.add("PR No");
			headers.add("Class of Vehicle");
			headers.add("Service Type");
			headers.add("Created Date");
			headers.add("Action Date");
			headers.add("Status");
			headers.add("IP Address");
			break;
		case "EodDistRoleCountList":	
			headers.add("S.No");
			headers.add("User Name");
			headers.add("Total Count");
			headers.add("Approve Count");
			headers.add("Reject Count");
		    break;
		case "EodDistCountList":	
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Role");
			headers.add("Total Count");
			headers.add("Approve Count");
			headers.add("Reject Count");
			break;
		case "ContractCarriagePermitsCount":	
			headers.add("S.No");
			headers.add("Permit Type");
			headers.add("Total");
			break;
		case "ContractCarriagePermits":
			headers.add("S.No");
			headers.add("Registration Number");
			headers.add("Owner Name");
			headers.add("Class of Vehicle");
			headers.add("Maker Name");
			headers.add("Chassis Number");
			headers.add("Engine Number");
			headers.add("Tax Paid Date");
			headers.add("FC Valid Upto");
			headers.add("Permit Issued date");
			headers.add("Permit Valid Upto");
			headers.add("Permit Type");
			headers.add("Tax Valid Upto");
			headers.add("Seating Capacity");
			headers.add("Present Address");
			break;
		case "InvoiceDetailsReport":
			headers.add("S.No");
			headers.add("Dealer Name");
			headers.add("TR Number");
			headers.add("TR date");
			headers.add("Class of vehicle");
			headers.add("Maker Name");
			headers.add("Maker Class");
			headers.add("Invoice Date");
			headers.add("Tax Type");
			headers.add("Invoice Amount");
			headers.add("Tax Amount");
			headers.add("Total");
			break;
		case "NonPaymentVehicleWiseList":
			headers.add("S.No");
			headers.add("Vehicle Class");
			headers.add("Non Payment Count");
			break;
		case "NonPaymentDetailsWiseList":
			headers.add("S.No");
			headers.add("Registration Number");
			headers.add("Vehicle Class");
			headers.add("Tax Validity");
			headers.add("Owner Name");
			headers.add("Owner Address");
			headers.add("Mandal");
			headers.add("Finance Name");
			headers.add("Finance Address");
			headers.add("FC Validity");
			headers.add("permit Validity");
			headers.add("GVW");
			headers.add("Mobile Number");
			break;
		case "VcrPaymentReportList":
			headers.add("S.No");
			headers.add("District Name");
			headers.add("Count");
			break;
		case "VcrPaymentReportListOfficeData":
			headers.add("SlNo");
			headers.add("Office Name");
			headers.add("Count");
			break;
		case "RcSuspensionReportsList":
			headers.add("S.No");
			headers.add("District Name");
			headers.add("Action Status");
			headers.add("count");
			break;
		case "RcSuspensionReportsListOfficeData":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Action Status");
			headers.add("count");
			break;
		case "RcSuspensionReportsListUserData":
			headers.add("S.No");
			headers.add("PR Number");
			headers.add("Owner Name");
			headers.add("Reg. Validity");
			headers.add("COV");
			headers.add("Action Status");
			headers.add("Suspended From");
			headers.add("Suspended To");
			headers.add("Suspend By");
			headers.add("Suspension Reason");
			headers.add("Ref Number");
			headers.add("Ref Date");
			headers.add("Revoked By");
			headers.add("Revoked Date");
			break;
		case "E-BiddingReportsList":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("No.of Bidders");
			headers.add("Registration Amount");
			headers.add("Bid Amount");
			headers.add("Service Fee");
			headers.add("Total Collected");
			break;
		case "PermitReports":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Permit Description");
			headers.add("Count");
			break;
		case "PermitReportsData":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Permit Description");
			headers.add("Class of Vehicle");
			headers.add("Permit Number");
			headers.add("PR Number");
			headers.add("Valid From");
			headers.add("Valid To");
			break;
		case "VehicleStrengthReport":
			headers.add("S.No");
			headers.add("District Names");
			headers.add("Count");
			break;
		case "VehicleStrengthReportOfficeData":
			headers.add("S.No");
			headers.add("Office Name");
			headers.add("Transport");
			headers.add("NonTransport");
			headers.add("Total");
			break;
		case "VehicleStrengthReportTransportData":
			headers.add("S.No");
			headers.add("Class of Vehicle");
			headers.add("Count");
			break;
		case "OffenceEnforcementReports":
			headers.add("S.No");
			headers.add("Offence");
			headers.add("Offence Count");
			headers.add("Amount");
			break;
		case "RoadSafetyMviCount":
			headers.add("S.No");
			headers.add("District Name");
			headers.add("Carrying extra persons in goods vehicle(per Head)");
			headers.add("Driving at Excessive Speed");
			headers.add("No Reflectors");
			headers.add("Non Wearing of helmets");
			headers.add("Non wearing of Seat belts");
			headers.add("Over load of goods vehicles");
			headers.add("Over loading Passengers");
			headers.add("Vehicle plying in Wrong Direction");
			headers.add("Total");
			break;
		case "RoadSafetyVcrCount":
			headers.add("S.No");
			headers.add("VCR Number");
			headers.add("Offence");
			headers.add("Booked");
			break;
		case "PaymentCheckPostReportOfficeCount":
			headers.add("S.No");
			headers.add("Office");
			headers.add("No.of Vcr's");
			headers.add("Paid Cf");
			headers.add("UnPaid Cf");
			headers.add("No.of Permit");
			headers.add("Permit Fee");
			headers.add("Permit Tax");
			headers.add("Vol.Tax Count");
			headers.add("Vol.Tax Amount");
			headers.add("TOTAL");
			break;
		case "PaymentCheckPostReportMviCount":
			headers.add("S.No");
			headers.add("Mvi Name");
			headers.add("No.of Vcr's");
			headers.add("Paid Cf");
			headers.add("UnPaid Cf");
			headers.add("No.of Permit");
			headers.add("Permit Fee");
			headers.add("Permit Tax");
			headers.add("Vol.Tax Count");
			headers.add("Vol.Tax Amount");
			headers.add("TOTAL");
			break;
		case "FitnessCertIssue":
			headers.add("S.No");
			headers.add("Vehicle Number");
			headers.add("Class Of Vehicle");
			headers.add("Valid From");
			headers.add("Valid Upto");
			headers.add("UserName");
			headers.add("Type");
			break;
		case "CovWiseVcrCount":
			headers.add("S.No");
			headers.add("COV");
			headers.add("COV Description");
			headers.add("VCR Count");
			break;
		case "CovWiseVcrMviCount":
			headers.add("S.No");
			headers.add("MVI");
			headers.add("COV Description");
			headers.add("VCR Count");
			break;
		case "OffenceWiseVcrCount":
			headers.add("S.No");
			headers.add("Offence Description");
			headers.add("Offence Count");
			break;
		case "OffenceWiseVcrMviCount":
			headers.add("S.No");
			headers.add("MVI");
			headers.add("Offence Name");
			headers.add("Offence Count");
			break;
		case "EvcrReportDetail":
			headers.add("S.No");
			headers.add("Name");
			headers.add("Count");
			break;
		case "EvcrReportDetailData":
			headers.add("S.No");
			headers.add("Registration Number");
			headers.add("Office Name");
			headers.add("Owner Name");
			headers.add("Date of VCR");
			headers.add("Class of Vehicle");
			headers.add("Status");
			headers.add("Amount");
		default:
			break;
		}

	}

	private void renderHeaders(List<String> headers, XSSFWorkbook wb, /* String sheetName, */ XSSFSheet sheet) {

		XSSFRow primaryRow = sheet.createRow(0);

		CellStyle style = wb.createCellStyle();// Create style
		XSSFFont font = wb.createFont();// Create font
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);// Make font bold
		font.setColor(HSSFColor.ORANGE.index);
		style.setFont(font);// set it to bold
		style.setAlignment(CellStyle.ALIGN_LEFT);

		for (int c = 0; c <= headers.size() - 1; c++) {
			XSSFCell cell = primaryRow.createCell(c);
			primaryRow.getCell(c).setCellStyle(style);
			cell.setCellValue(headers.get(c).toUpperCase());
		}

	}

	@Override
	public XSSFWorkbook renderData(List<List<CellProps>> result, List<String> headers, String fileName,
			String sheetName) {

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName);
		renderHeaders(headers, wb, /* sheetName, */sheet);

		CellStyle style = wb.createCellStyle();// Create style
		XSSFFont font = wb.createFont();// Create font
		font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);// Make font bold
		font.setColor(HSSFColor.BLACK.index);
		style.setFont(font);// set it to bold

		Integer rowIndex = 1;
		Integer headerSize = headers.size();

		for (List<CellProps> cellProp : result) {

			XSSFRow row = sheet.createRow(rowIndex);

			int cellIndex = 0;

			for (CellProps cellV : cellProp) {

				XSSFCell cell = row.createCell(cellIndex);
				row.getCell(cellIndex).setCellStyle(style);
				if (cellV.getFieldValue() != null)
					cell.setCellValue(cellV.getFieldValue().toString());
				try {
					if (cellV.getFieldValue() != null) {
						double number = Double.parseDouble(cellV.getFieldValue());
						cell.setCellValue(number);
					}
				} catch (NumberFormatException ne) {
					logger.debug("Exception [{}]", ne);
					logger.error("Exception [{}]", ne.getMessage());
				}
				cellIndex++;
				if (cellIndex == headerSize) {
					cellIndex = 0;
					row = sheet.createRow(++rowIndex);
				}
			}

		}
		for (int i = 0; i < headers.size(); i++) {
			sheet.autoSizeColumn(i);
		}
		return wb;
	}

}
