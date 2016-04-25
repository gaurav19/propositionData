package common;

public class ReadExcel {

	String pyName;
	String businessGroup;
	String strategy;
	String eligibility;
	
	public ReadExcel(){
		pyName = "";
		businessGroup = "";
		strategy = "";
		eligibility = "";
	}
	
	public String getPyName() {
		return pyName;
	}
	public void setPyName(String pyName) {
		this.pyName = pyName;
	}
	public String getBusinessGroup() {
		return businessGroup;
	}
	public void setBusinessGroup(String businessGroup) {
		this.businessGroup = businessGroup;
	}
	public String getStrategy() {
		return strategy;
	}
	public void setStrategy(String strategy) {
		this.strategy = strategy;
	}
	public String getEligibility() {
		return eligibility;
	}
	public void setEligibility(String eligibility) {
		this.eligibility = eligibility;
	}
}
