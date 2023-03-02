package com.createExceltest.testing;



import javax.persistence.Entity;
import javax.persistence.Id;

import lombok.Data;
@Data
@Entity
public class Excel {

	@Id
	private String company_code;
	private String sales_Organization;
	private String distribution_Channel;
	private String division;
	private String customer_Account_Group;
    private String name1;
	private String searchTerm1;
	private String c_o_name;
	private String street2;
	private String street3;
	private String street1;
	private String district;
	private String city_postal_code;
	private String city;
	private String countryKey;
	private String region;
	private String timeZone;
	private String regionalStructure;
	private String languageKey;
	private String firstTelephone;
	private String firstCellPhone;
	private String firstFax;
	private String email;
	private String panNo;
	private String gstReg;
	private String vatRegistrationNo;
	private String agriLicNo;
	private String pest;
	private String reconciliation;
	private String keyForSorting;
	private String planningGroup;
	private String paymentTerm;
	private String indicatorRecord;
	private String indicatorForPeriod;
	private String regZone;
	private String orderProbability ;
	private String salesOffice;
	private String manager;
	private String customer_group;
	private String currency;
	private String priceGroup;
	private String pricing_procedure;
	private String customer_statistics_group;
	private String delivery_priority;
	private String order_combination_indicator;
	private String shipping_condition;
	private String delivering_plan;
	private String maximum_number;
	private String indicator_customer_is_rebate_relevant;
	private String incoterms1 ;
	private String incoterms2 ;
	private String terms_of_payment_key;
	private String customer_payment_guarantee_procedure;
	private String credit_control_area;
	private String account_assignment_group;
	private String tax_classification_cust1;
	private String tax_classification_cust2;
	private String tax_classification_cust3;
	private String tax_classification_cust4;
	private String tax_classification_cust5;
	private String tax_classification_cust6;
	private String tax_classification_cust7;
	private String tax_classification_cust8;
	private String tax_classification_cust9;
	private String asmS_Exe;
	private String salesOfficerSenior;
	private String district2;
	private String issueDate1;
	private String validFromdate1;
	private String issuedate2;
	private String validFromdate2;
	private String validTodate1;
	private String issueDate3;
	private String validFromdate4;
	private String validTodate2;		
}
