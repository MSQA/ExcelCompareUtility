import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang.StringEscapeUtils;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;

public class ExcelLibrary {
	static Workbook lookupWorkBook;
	static LinkedHashMap dict = new LinkedHashMap();//new Hashtable();
	static String lookupFile = "E:\\PROJECTS\\lookup.xls"; 
	static List<Map<String, String>> list = new ArrayList<Map<String, String>>();
	static Map<String, String> map = new HashMap<String, String>();
	static int navcnt = 0;
	
	
	public static int navCount(int navCount){
		navcnt = navCount;		
		return navcnt;
	}
	
	
	
	public static void Navigation(){		
	//****************Read Look Up File******************//
	try {
	lookupWorkBook = Workbook.getWorkbook(new File(lookupFile));
	int rowCount = lookupWorkBook.getSheet("LookUp").getRows();
	String objKey, objVal;
	int sheetRefColcnt = lookupWorkBook.getSheet("LookUp").getColumns();//2
	int sheetRefRowcnt = lookupWorkBook.getSheet("LookUp").getRows();//6
		for (int itrSheetRow = 1; itrSheetRow < sheetRefRowcnt ; itrSheetRow++) {
			String sheetColid = lookupWorkBook.getSheet("LookUp").getCell(0,itrSheetRow).getContents();
				for (int itrSheetCol = 0; itrSheetCol < sheetRefColcnt-1; itrSheetCol++) {
					objKey = lookupWorkBook.getSheet("LookUp").getCell(itrSheetCol,itrSheetRow).getContents();
					objVal = StringEscapeUtils.unescapeJava(lookupWorkBook.getSheet("LookUp").getCell(itrSheetCol+1,itrSheetRow).getContents());					
					if(objVal!=""){
					/*System.out.println(objKey);
					System.out.println(objVal);*/
					map.put(objKey, objVal);
					list.add(map);
					}
				}		
			}
		System.out.println("I am in ExcelLibrary Function- Navigation!!");
	//****************End Read Look Up File******************//
	//****************Get value from dictionary********************//
		Set set = dict.entrySet();
		Iterator i = set.iterator();
	    while(i.hasNext()) {
	         Map.Entry me = (Map.Entry)i.next();
	         String navVal = me.getKey().toString();//Key from the dict
	         if(navVal.contains("Navigation")){
		         String[] splitDictKey = navVal.split("[Navigation]"); 
		         //System.out.println(splitDictKey[splitDictKey.length-1]);
		         int index = Integer.parseInt(splitDictKey[splitDictKey.length-1]);
		         if(index == navcnt){
		        	//List<List<String>> navValue = new ArrayList<List<String>>(dict.values());
		     		Set<String>  keyset = map.keySet();		
		     		String[] temp = dict.get(navVal).toString().split("[|]");	//split navigation value
		     		for(int m=0;m<temp.length;m++){			
		     			String test = temp[m].toString();		  
		     			Iterator itr = keyset.iterator();
		     			while(itr.hasNext()){
		     				 String keyset1 = (String) itr.next();
		     				 if(test.equals(keyset1)){
		     				 //System.out.println("Key matched");
		     				 String nav1 = map.get(keyset1);	 
		     				 //System.out.println("Navi1 has value ="+nav1);
		     				 //web.element(nav1).click();
		     				 } 
		     			 }
		     		}//End for
		         }//End if- index
	         }
	    }// While to get value from dict and count  
	  
	    
		
	//****************End Get value from dictionary********************//	
	//****************Compare values******************//	
	//****************End Compare******************//	
		}catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}
	
	
	
	public static void getExcelData(String sheetRef, String colTcid,Workbook readWorkBook){
		String objKey, objVal;
		int sheetRefColcnt = readWorkBook.getSheet(sheetRef).getColumns();
		int sheetRefRowcnt = readWorkBook.getSheet(sheetRef).getRows();
		for (int itrSheetRow = 1; itrSheetRow < sheetRefRowcnt ; itrSheetRow++) {
			String sheetColid = readWorkBook.getSheet(sheetRef).getCell(0,itrSheetRow).getContents();
			if(sheetColid.equalsIgnoreCase(colTcid)){
				for (int itrSheetCol = 1; itrSheetCol < sheetRefColcnt; itrSheetCol++) {
					objKey = readWorkBook.getSheet(sheetRef).getCell(itrSheetCol,0).getContents();
					objVal = readWorkBook.getSheet(sheetRef).getCell(itrSheetCol,itrSheetRow).getContents();					
					if(objVal!=""){
					/*System.out.println(objKey);*/
					System.out.println(objVal);
					dict.put(objKey,objVal);
					}
				}
			}			
		}
	}
	
	public static void putExcelData(String sheetRef, String colTcid,Workbook readWorkBook,String passValue,String wrtFile) throws RowsExceededException, WriteException, IOException{
		String objKey, objVal;
		WritableWorkbook wrtWorkBook = Workbook.createWorkbook(new File(wrtFile),readWorkBook);
		int sheetRefColcnt = wrtWorkBook.getSheet(sheetRef).getColumns();
		int sheetRefRowcnt = wrtWorkBook.getSheet(sheetRef).getRows();
		for (int itrSheetRow = 1; itrSheetRow < sheetRefRowcnt ; itrSheetRow++) {
			String sheetColid = wrtWorkBook.getSheet(sheetRef).getCell(0,itrSheetRow).getContents();
 			if(sheetColid.equalsIgnoreCase(colTcid)){
 				String getColHead = wrtWorkBook.getSheet(sheetRef).getCell(sheetRefColcnt-1,0).getContents();
 					if(getColHead.equalsIgnoreCase("OrderNumber")){
					     WritableSheet wrtSheet = wrtWorkBook.getSheet(sheetRef);
					     Label wrtLabel = new Label(sheetRefColcnt-1,itrSheetRow,passValue);
					     wrtSheet.addCell(wrtLabel);
					     wrtWorkBook.write();	
					     objKey = wrtWorkBook.getSheet(sheetRef).getCell(sheetRefColcnt-1,0).getContents();
					     objVal = wrtWorkBook.getSheet(sheetRef).getCell(sheetRefColcnt-1,itrSheetRow).getContents();
					    					     
					     if(objVal!=""){
					    	 /*System.out.println(objKey);
							 System.out.println(objVal);*/
							 dict.put(objKey,objVal);
						}
 					}else{
 					    wrtWorkBook.write();	 					
 					}
			}	 			
		}
		wrtWorkBook.close();
	}
	

	
	public static String add(String a,String b){
	String c;
	c=a+b;
	return c;
	}
	
	public static void Login(){
		System.out.println("I am in ExcelLibrary!! Function-Login");
	}
	
	public static void OMS_Order_Entry(){
		System.out.println("I am in ExcelLibrary Function-OMS_Order_Entry!!");
	}
	
	public static void CSR_Inquiry(){
		System.out.println("I am in ExcelLibrary Function-CSR_Inquiry!!");
	}
}

