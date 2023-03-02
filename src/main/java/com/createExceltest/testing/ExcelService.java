package com.createExceltest.testing;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;



@Service
public class ExcelService {

	

	public String getExcelFormat(List<Excel> excel) {
		try{
			String timeStamp = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
		//String timeStamp = new SimpleDateFormat("yyyy.MM.ddHH:mm:ss").format(new Date());
			//concanting the fileName and timeStamp to generate unique number
		String filename = new String ("D:\\demo\\CPB_CUST_FORMAT_"+timeStamp+".xlsx"); 
		System.out.println("filenme: "+filename);
		//creating an instance of HSSFWorkbook class  
		HSSFWorkbook workbook = new HSSFWorkbook();  
		//invoking creatSheet() method and passing the name of the sheet to be created   
		HSSFSheet sheet = workbook.createSheet("Sample");   
		//creating the 0th row using the createRow() method  
		HSSFRow rowhead = sheet.createRow((short)0); 

		
				
		        rowhead.createCell(0).setCellValue("Company Code");
		        rowhead.createCell(1).setCellValue("Sales Organization");
		        rowhead.createCell(2).setCellValue("Distribution Channel");
		        rowhead.createCell(3).setCellValue("DIVISION");
		        rowhead.createCell(4).setCellValue("Customer Account Group");
				rowhead.createCell(5).setCellValue("Name 1");  
				rowhead.createCell(6).setCellValue("Search Term 1");  
				rowhead.createCell(7).setCellValue("c o name");
				rowhead.createCell(8).setCellValue("Street 2");
				rowhead.createCell(9).setCellValue("Street 3");
				rowhead.createCell(10).setCellValue("Street 1");
				rowhead.createCell(11).setCellValue("District");
				rowhead.createCell(12).setCellValue("City postal code");
				rowhead.createCell(13).setCellValue("City");
				rowhead.createCell(14).setCellValue("Country Key");
				rowhead.createCell(15).setCellValue("Region (State, Province, County)");
				rowhead.createCell(16).setCellValue("Time Zone");
				rowhead.createCell(17).setCellValue("Regional structure grouping");
				rowhead.createCell(18).setCellValue("Language Key");
				rowhead.createCell(19).setCellValue("First telephone no.: dialling code+number");
				rowhead.createCell(20).setCellValue("First Cell Phone Number: Dialing Code + Number");
				rowhead.createCell(21).setCellValue("First fax no.: dialling code+number");
				rowhead.createCell(22).setCellValue("E-Mail Address");
				rowhead.createCell(23).setCellValue("PAN No");
				rowhead.createCell(24).setCellValue("GST Reg. No.");
				rowhead.createCell(25).setCellValue("VAT Registration Number");
				rowhead.createCell(26).setCellValue("Agri Lic No.");
				rowhead.createCell(27).setCellValue("Pestic Lic No.");
				rowhead.createCell(28).setCellValue("Reconciliation Account in General Ledger");
				rowhead.createCell(29).setCellValue("Key for sorting according to assignment numbers");
				rowhead.createCell(30).setCellValue("Planning group");
				rowhead.createCell(31).setCellValue("Payment Term");
				rowhead.createCell(32).setCellValue("Indicator: Record Payment History ?");
				rowhead.createCell(33).setCellValue("Indicator for periodic account statements");
				rowhead.createCell(34).setCellValue("Reg Zone");
				rowhead.createCell(35).setCellValue("Order probability of the item");
				rowhead.createCell(36).setCellValue("Sales office");
				rowhead.createCell(37).setCellValue("Manager");
				rowhead.createCell(38).setCellValue("Customer group");
				rowhead.createCell(39).setCellValue("Currency");
				rowhead.createCell(40).setCellValue("Price group (customer)");
				rowhead.createCell(41).setCellValue("Pricing procedure assigned to this customer");
				rowhead.createCell(42).setCellValue("Customer Statistics Group");
				rowhead.createCell(43).setCellValue("Delivery Priority");
				rowhead.createCell(44).setCellValue("Order Combination Indicator");
				rowhead.createCell(45).setCellValue("Shipping conditions");
				rowhead.createCell(46).setCellValue("Delivering Plant (Own or External)");
				rowhead.createCell(47).setCellValue("Maximum Number of Partial Deliveries Allowed Per Item");
				rowhead.createCell(48).setCellValue("Indicator: Customer Is Rebate-Relevant");
				rowhead.createCell(49).setCellValue("Incoterms (Part 1)");
				rowhead.createCell(50).setCellValue("Incoterms (Part 2)");
				rowhead.createCell(51).setCellValue("Terms of Payment Key");
				rowhead.createCell(52).setCellValue("Customer payment guarantee procedure");
				rowhead.createCell(53).setCellValue("Credit control area");
				rowhead.createCell(54).setCellValue("Account assignment group for this customer");
				rowhead.createCell(55).setCellValue("Tax classification for customer");
				rowhead.createCell(56).setCellValue("Tax classification for customer");
				rowhead.createCell(57).setCellValue("Tax classification for customer");
				rowhead.createCell(58).setCellValue("Tax classification for customer");
				rowhead.createCell(59).setCellValue("Tax classification for customer");
				rowhead.createCell(60).setCellValue("Tax classification for customer");
				rowhead.createCell(61).setCellValue("Tax classification for customer");
				rowhead.createCell(62).setCellValue("Tax classification for customer");
				rowhead.createCell(63).setCellValue("Tax classification for customer");
				rowhead.createCell(64).setCellValue("ASM S.EXE");
				rowhead.createCell(65).setCellValue("SALES OFFICER SENIOR S.O ESALES OFFICER SENIOR S.O EXECUTIVE");
				rowhead.createCell(66).setCellValue("DISRTICT");
				rowhead.createCell(67).setCellValue("Issue Date");
				rowhead.createCell(68).setCellValue("Valid From Date");
				rowhead.createCell(69).setCellValue("Issue Date");
				rowhead.createCell(70).setCellValue("Valid From Date");
				rowhead.createCell(71).setCellValue("Valid To Date");
				rowhead.createCell(72).setCellValue("Issue Date");
				rowhead.createCell(73).setCellValue("Valid From Date");
				rowhead.createCell(74).setCellValue("Valid To Date");
	
				
				
				
				//creating the 1st row  
				//HSSFRow row = sheet.createRow((short)1);  
				//inserting data in the first row  
							
				int dataRowIndex = 1;

				for (Excel excel1 : excel) {
					HSSFRow dataRow = sheet.createRow(dataRowIndex);
					dataRow.createCell(0).setCellValue(excel1.getCompany_code());
					dataRow.createCell(1).setCellValue(excel1.getSales_Organization());
					dataRow.createCell(2).setCellValue(excel1.getDistribution_Channel());
					dataRow.createCell(3).setCellValue(excel1.getDivision());
					dataRow.createCell(4).setCellValue(excel1.getCustomer_Account_Group());
					dataRow.createCell(5).setCellValue(excel1.getName1());
					dataRow.createCell(6).setCellValue(excel1.getSearchTerm1());
					dataRow.createCell(7).setCellValue(excel1.getC_o_name());
					dataRow.createCell(8).setCellValue(excel1.getStreet2());
					dataRow.createCell(9).setCellValue(excel1.getStreet3());
					dataRow.createCell(10).setCellValue(excel1.getStreet1());
					dataRow.createCell(11).setCellValue(excel1.getDistrict());
					dataRow.createCell(12).setCellValue(excel1.getCity_postal_code());
					dataRow.createCell(13).setCellValue(excel1.getCity());
					dataRow.createCell(14).setCellValue(excel1.getCountryKey());
					dataRow.createCell(15).setCellValue(excel1.getRegion());
					dataRow.createCell(16).setCellValue(excel1.getTimeZone());
					dataRow.createCell(17).setCellValue(excel1.getRegionalStructure());
					dataRow.createCell(18).setCellValue(excel1.getLanguageKey());
					dataRow.createCell(19).setCellValue(excel1.getFirstTelephone());
					dataRow.createCell(20).setCellValue(excel1.getFirstCellPhone());
					dataRow.createCell(21).setCellValue(excel1.getFirstFax());
					dataRow.createCell(22).setCellValue(excel1.getEmail());
					dataRow.createCell(23).setCellValue(excel1.getPanNo());
					dataRow.createCell(24).setCellValue(excel1.getGstReg());
					dataRow.createCell(25).setCellValue(excel1.getVatRegistrationNo());
					dataRow.createCell(26).setCellValue(excel1.getAgriLicNo());
					dataRow.createCell(27).setCellValue(excel1.getPest());
					dataRow.createCell(28).setCellValue(excel1.getReconciliation());
					dataRow.createCell(29).setCellValue(excel1.getKeyForSorting());
					dataRow.createCell(30).setCellValue(excel1.getPlanningGroup());
					dataRow.createCell(31).setCellValue(excel1.getPaymentTerm());
					dataRow.createCell(32).setCellValue(excel1.getIndicatorRecord());
					dataRow.createCell(33).setCellValue(excel1.getIndicatorForPeriod());
					dataRow.createCell(34).setCellValue(excel1.getRegZone());
					dataRow.createCell(35).setCellValue(excel1.getOrderProbability());
					dataRow.createCell(36).setCellValue(excel1.getSalesOffice());
					dataRow.createCell(37).setCellValue(excel1.getManager());
					dataRow.createCell(38).setCellValue(excel1.getCustomer_group());
					dataRow.createCell(39).setCellValue(excel1.getCurrency());
					dataRow.createCell(40).setCellValue(excel1.getPriceGroup());
					dataRow.createCell(41).setCellValue(excel1.getPricing_procedure());
					dataRow.createCell(42).setCellValue(excel1.getCustomer_statistics_group());
					dataRow.createCell(43).setCellValue(excel1.getDelivery_priority());
					dataRow.createCell(44).setCellValue(excel1.getOrder_combination_indicator());
					dataRow.createCell(45).setCellValue(excel1.getShipping_condition());
					dataRow.createCell(46).setCellValue(excel1.getDelivering_plan());
					dataRow.createCell(47).setCellValue(excel1.getMaximum_number());
					dataRow.createCell(48).setCellValue(excel1.getIndicator_customer_is_rebate_relevant());
					dataRow.createCell(49).setCellValue(excel1.getIncoterms1());
					dataRow.createCell(50).setCellValue(excel1.getIncoterms2());
					dataRow.createCell(51).setCellValue(excel1.getTerms_of_payment_key());
					dataRow.createCell(52).setCellValue(excel1.getCustomer_payment_guarantee_procedure());
					dataRow.createCell(53).setCellValue(excel1.getCredit_control_area());
					dataRow.createCell(54).setCellValue(excel1.getAccount_assignment_group());
					dataRow.createCell(55).setCellValue(excel1.getTax_classification_cust1());
					dataRow.createCell(56).setCellValue(excel1.getTax_classification_cust2());
					dataRow.createCell(57).setCellValue(excel1.getTax_classification_cust3());
					dataRow.createCell(58).setCellValue(excel1.getTax_classification_cust4());
					dataRow.createCell(59).setCellValue(excel1.getTax_classification_cust5());
					dataRow.createCell(60).setCellValue(excel1.getTax_classification_cust6());
					dataRow.createCell(61).setCellValue(excel1.getTax_classification_cust7());
					dataRow.createCell(62).setCellValue(excel1.getTax_classification_cust8());
					dataRow.createCell(63).setCellValue(excel1.getTax_classification_cust9());
					dataRow.createCell(64).setCellValue(excel1.getAsmS_Exe());
					dataRow.createCell(65).setCellValue(excel1.getSalesOfficerSenior());
					dataRow.createCell(66).setCellValue(excel1.getDistrict2());
					dataRow.createCell(67).setCellValue(excel1.getIssueDate1());
					dataRow.createCell(68).setCellValue(excel1.getValidFromdate1());
					dataRow.createCell(69).setCellValue(excel1.getIssuedate2());
					dataRow.createCell(70).setCellValue(excel1.getValidFromdate2());
					dataRow.createCell(71).setCellValue(excel1.getValidTodate1());
					dataRow.createCell(72).setCellValue(excel1.getIssueDate3());
					dataRow.createCell(73).setCellValue(excel1.getValidFromdate4());
					dataRow.createCell(74).setCellValue(excel1.getValidTodate2());
		
					dataRowIndex++;
				}
				
				FileOutputStream fileOut = new FileOutputStream(filename);  
				workbook.write(fileOut);  
				//closing the Stream  
				fileOut.close();  
				//closing the workbook  
				workbook.close();  
				System.out.println("Excel file has been generated successfully.");  
	}   
	catch (Exception e)   
	{  
	e.printStackTrace();  
	}  	
		
		return "successfull";
	}
	
	}

