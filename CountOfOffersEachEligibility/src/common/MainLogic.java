package common;

import java.util.ArrayList;

public class MainLogic {

	public void selectOffers(){
		WriteObject wObj;
		ArrayList<WriteObject> arWO = new ArrayList<WriteObject>();
		
		ReadWriteExcel rdExl = new ReadWriteExcel();

		ArrayList<ReadExcel> rdExcl = rdExl.readExcel();
		int counter = rdExcl.size();
		
		for(int i=1; i<600; i++){
			String indc = "_" + i + "_";
			int cnt = 0;
			String elig = "", details = "";
			
			for(int j=0; j<counter; j++){
				elig = rdExcl.get(j).getEligibility();
				if(elig.contains(indc)){
					cnt += 1; 
					details += rdExcl.get(j).getPyName() + " -> " + rdExcl.get(j).getBusinessGroup() + " -> " + rdExcl.get(j).getStrategy() + "\n";
				}
			}
			
			if(cnt > 0){
				wObj = new WriteObject();
				wObj.setRule(indc);
				wObj.setPyName_ProdGrp_Strgy(details);
				arWO.add(wObj);
			}
		}
		System.out.println("Rule finding is complete...\nGenerating excel report now..");
		rdExl.writeExcel(arWO);
		
		System.out.println("Excel Report has been generated..");
	}
	
	public static void main(String[] args){
		MainLogic mLo = new MainLogic();
		mLo.selectOffers();
	}
}
