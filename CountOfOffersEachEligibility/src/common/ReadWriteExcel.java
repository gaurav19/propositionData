package common;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadWriteExcel {
	
	public WritableCellFormat setStyle(boolean hdrOrBody){
		int x = (hdrOrBody ? 11 : 9);
		WritableFont cellFont = new WritableFont(WritableFont.ARIAL, x);
		try {
			if(hdrOrBody)
				cellFont.setBoldStyle(WritableFont.BOLD);
		} catch (WriteException e1) {
			e1.printStackTrace();
		}
		
        WritableCellFormat fCellstatus = new WritableCellFormat(cellFont);
        try {
			fCellstatus.setWrap(true);
			fCellstatus.setAlignment(hdrOrBody ? jxl.format.Alignment.CENTRE : jxl.format.Alignment.LEFT);
		    fCellstatus.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN, jxl.format.Colour.BLACK);
		    if(hdrOrBody)
		    	fCellstatus.setBackground(jxl.format.Colour.ORANGE);
				
		} catch (WriteException e) {
			e.printStackTrace();
		}
        return fCellstatus;
	}
	
	public void writeExcel(ArrayList<WriteObject> arWO){
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
		String stringdate  = dateFormat.format(new Date());
		
		String fileName = "WriteOutput-" + stringdate + ".xls";
		fileName.replace(' ', '-');
		try {
			String path = System.getProperty("user.dir") + File.separator + fileName;
            File exlFile = new File(path);
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);
 
            WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);
            
            
           
            
            Label ruleNum = new Label(0, 0, "Rule#", setStyle(true));
            Label details = new Label(1, 0, "Details", setStyle(true));

            //Add the created Cells to the sheet
            writableSheet.addCell(ruleNum);
            writableSheet.addCell(details);
            
            for(int i=0; i<arWO.size(); i++){
            	Label pyRule = new Label(0, i+1, arWO.get(i).getRule(), setStyle(false));
                Label pyDetails = new Label(1, i+1, arWO.get(i).getPyName_ProdGrp_Strgy(), setStyle(false));
                
                writableSheet.addCell(pyRule);
                writableSheet.addCell(pyDetails);
            }
            
            writableSheet.setColumnView(0, 15);
            writableSheet.setColumnView(1, 130);
            
            //Write and close the workbook
            writableWorkbook.write();
            writableWorkbook.close();
 
        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }

	}
	
	public ArrayList<ReadExcel> readExcel(){
            
		String path = System.getProperty("user.dir") + File.separator + "ExcelRead.xls";
        Workbook wrk1;

        ArrayList<ReadExcel> rdExl = new ArrayList<ReadExcel>();
        ReadExcel rde;
        
		try {
			wrk1 = Workbook.getWorkbook(new File(path));
			Sheet sheet1 = wrk1.getSheet(0);
	        int noOfRecords = sheet1.getRows();
	         
	        for(int i=1; i<noOfRecords; i++){
	            Cell colArow = sheet1.getCell(0, i);
	            Cell colBrow = sheet1.getCell(1, i);
	            Cell colCrow = sheet1.getCell(2, i);
	            Cell colDrow = sheet1.getCell(3, i);
	        
	            String pyNme = colArow.getContents();
	            String busnGrp = colBrow.getContents();
	            String strgy = colCrow.getContents();
	            String elig = colDrow.getContents();
	            
	            rde = new ReadExcel();
	            rde.setPyName(pyNme);
	            rde.setBusinessGroup(busnGrp);
	            rde.setStrategy(strgy);
	            rde.setEligibility(elig);
	            
	            rdExl.add(rde);
	        }
	        
		} catch (BiffException | IOException e) {
			e.printStackTrace();
		}
		
		return rdExl;
	}
	
}
