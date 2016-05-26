import java.io.*;
import java.util.LinkedHashMap;
import jxl.*;
import jxl.write.*;
import java.lang.reflect.*;
/*import ExcelDemoTest;*/
public class ExcelController {
static Workbook readWorkBook;

static WritableWorkbook sendWorkBook; 
static WritableSheet wrtSheet;
static WritableWorkbook wrtWorkBook;

static LinkedHashMap<?, ?> lookUpMap = new LinkedHashMap<Object, Object>(); 
	public static void main(String[] args) throws WriteException, IOException {
	try{
		String writeCellVal = "";
		String file = "E:\\PROJECTS\\test2.xls";
		String wrtFile = "E:\\PROJECTS\\test2.xls";	
		int ctlrNavCount = 1;
		ExcelLibrary obj = new ExcelLibrary();
		Method method;
		readWorkBook =  Workbook.getWorkbook(new File(file));		
		wrtWorkBook = Workbook.createWorkbook(new File(wrtFile),readWorkBook);
		wrtWorkBook.write();
		wrtWorkBook.close();
		int tcCount = readWorkBook.getSheet("Main").getRows();  
		for (int tcIteration = 1; tcIteration < tcCount; tcIteration++) {
		String colExecute = readWorkBook.getSheet("Main").getCell(0,tcIteration).getContents();
		String curMainTcid=readWorkBook.getSheet("Main").getCell(1,tcIteration).getContents(); //Get 1st Test Case id from Main sheet
		if(colExecute.equalsIgnoreCase("YES")){		
		//String prevStepid=readWorkBook.getSheet("TestSuite").getCell(0,1).getContents(); //Get 1st Test Case id from Test Suite
		 int stepCount = readWorkBook.getSheet("TestSuite").getRows();  
		  for (int tsIteration = 1; tsIteration < stepCount; tsIteration++) {	
			  String sheetRef = "";
			  String fncName="";
			   String curStepid = readWorkBook.getSheet("TestSuite").getCell(0,tsIteration).getContents();
			   String colTcName= readWorkBook.getSheet("TestSuite").getCell(2,tsIteration).getContents();    
			   if(curMainTcid.equalsIgnoreCase(curStepid)){ 
				    sheetRef = colTcName;
				    fncName = readWorkBook.getSheet("TestSuite").getCell(3,tsIteration).getContents();
				    ExcelLibrary.getExcelData(sheetRef,curStepid,readWorkBook);
				    if(sheetRef.equalsIgnoreCase("Navigation") && curMainTcid.equalsIgnoreCase(curStepid)){
				    	ExcelLibrary.navCount(ctlrNavCount);
				    	ctlrNavCount++;
				    }				    
				    method = obj.getClass().getMethod(fncName, new Class<?>[0]);				    
			        method.invoke(obj);
			        String orderNo = String.valueOf(Math.random());
				    writeCellVal = ExcelLibrary.add("TC ",orderNo);
				    if(fncName.equalsIgnoreCase("OMS_Order_Entry")){
				    	System.out.println("Writing in OMS " +orderNo);
				    	ExcelLibrary.dict.remove("OrderNumber");
				    ExcelLibrary.putExcelData(sheetRef, curStepid, readWorkBook, writeCellVal, wrtFile);
				    }
			   }/*else{
				   	ctlrNavCount = 1;
				   	prevStepid = curStepid; //If prev tcid and cur tcid dont match then set curr tc  id as prev tc id
				   	sheetRef = colTcName;
				   	fncName = readWorkBook.getSheet("TestSuite").getCell(3,tsIteration).getContents();
				    ExcelLibrary.getExcelData(sheetRef,curStepid,readWorkBook);
				    if(sheetRef.equalsIgnoreCase("Navigation") && prevStepid.equalsIgnoreCase(curStepid)){
				    	ExcelLibrary.navCount(ctlrNavCount);
				    	ctlrNavCount++;
				    }
				    method = obj.getClass().getMethod(fncName, new Class<?>[0]);
			        method.invoke(obj);
				    ExcelLibrary.getExcelData(sheetRef,curStepid,readWorkBook);
				    writeCellVal = ExcelLibrary.add("TC"," Order2");
				    ExcelLibrary.putExcelData(sheetRef, curStepid, readWorkBook, writeCellVal, wrtFile);
			   }// End colTcname If-Else    
*/			
			  readWorkBook.close();
			  readWorkBook = Workbook.getWorkbook(new File(wrtFile));
			  }//End For Test-Suite
		  	}//End ColExecute 
			System.out.println("End of test case " +tcIteration);
		  }//End For Main
		  readWorkBook.close();
		  }catch(Exception e){
		  readWorkBook.close();
		  e.printStackTrace();
		  } 
	}	
}
	
	
