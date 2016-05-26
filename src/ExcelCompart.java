import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;

public class ExcelCompart {

	
	public static void checkColumns(HSSFWorkbook readSrcWorkBook,HSSFWorkbook readDestWorkBook) throws IOException{//,HSSFWorkbook writeDestWorkBook){
		boolean doesExist=false;
		String srColName=null;
		String destColName= null;
		HSSFWorkbook writeDestWorkBook =  readSrcWorkBook;
		FileOutputStream fileOut = null;
		
		int destColCount = readDestWorkBook.getSheetAt(0).getRow(1).getLastCellNum();
		int srcColCount = readSrcWorkBook.getSheetAt(0).getRow(1).getLastCellNum();
		String name = "D:\\"+readSrcWorkBook.toString()+".xls";
		
		
		fileOut = new FileOutputStream(name);
		
		
		//Check if the column names are same :
		for(short i = 0; i < srcColCount; i++){
			srColName=readSrcWorkBook.getSheetAt(0).getRow(0).getCell(i).getRichStringCellValue().toString();
			for (short j=0; j<destColCount;j++)
			{
				destColName=readDestWorkBook.getSheetAt(0).getRow(0).getCell(j).getRichStringCellValue().toString();
				if(srColName.equals(destColName)){
					doesExist = true;
				}
				
			}
			if(!doesExist){
				try {
					fileOut = new FileOutputStream(name);
					HSSFCellStyle style = writeDestWorkBook.createCellStyle();
					HSSFCell testColor =null;
					for (int k=0; k<=readDestWorkBook.getSheetAt(0).getLastRowNum();k++){
					
					testColor =writeDestWorkBook.getSheetAt(0).getRow(k).getCell(i);
					style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					testColor.setCellStyle(style);
					writeDestWorkBook.write(fileOut);
					
					}
					
				} catch (FileNotFoundException e) {
				
					e.printStackTrace();
				} catch (IOException e) {
					
					e.printStackTrace();
				}
				System.out.println(srColName+" is not present in the destination folder");
				
			}
			doesExist=false;
		}
		
		fileOut.close();
		
	}
	
	public static String getvalue(HSSFWorkbook WorkBook,int rows,short cols,short colu){

		String data= null;
		switch (WorkBook.getSheetAt(0).getRow(rows).getCell(cols).getCellType())
        {
            case Cell.CELL_TYPE_NUMERIC:
            	data = "0"+WorkBook.getSheetAt(0).getRow(rows).getCell(colu).getNumericCellValue();
            	 
            	 
                break;
            case Cell.CELL_TYPE_STRING:
            	data = "1"+WorkBook.getSheetAt(0).getRow(rows).getCell(colu).getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_FORMULA:
             
                break;
        }
		return data;
		
	}
	
	
	
	public static void compare(HSSFWorkbook readSrcWorkBook,HSSFWorkbook readDestWorkBook,String Pk, String fileName) throws Exception{
		int srcColCount = readSrcWorkBook.getSheetAt(0).getRow(0).getLastCellNum();
		int srcRowCount = readSrcWorkBook.getSheetAt(0).getPhysicalNumberOfRows();
		int destColCount = readDestWorkBook.getSheetAt(0).getRow(0).getLastCellNum();
		int destRowCount = readDestWorkBook.getSheetAt(0).getPhysicalNumberOfRows();
		HSSFWorkbook writeScrWorkBook =  readSrcWorkBook;
		HSSFWorkbook writeDestWorkBook =  readDestWorkBook;
		
		FileOutputStream DestFileOut =  new FileOutputStream(fileName);
		int counter = 0;
		
		
		
		for(short cols = 0; cols < srcColCount; cols++){
			if(readSrcWorkBook.getSheetAt(0).getRow(0).getCell(cols).getRichStringCellValue().toString().equals(Pk)){//HEADER
				for(int rows =1; rows<srcRowCount;rows++){
					
				
				
				double colSrcText = 0;
				String colSrcText1 = null;
				double colDestText=0;
				String colDestText1= null;
				
				String srcData=getvalue(readSrcWorkBook , rows, cols,cols);
				if(srcData.startsWith("0")){
					colSrcText = Integer.parseInt(srcData.substring(1));
				}
				else{
					colSrcText1 =srcData.substring(1);
				}
				
				
				int flag =0;
				
				for(int row = 1; row < destRowCount; row++){
					
					
					
				String DestData=getvalue(readDestWorkBook , row, cols,cols);
					if(DestData.startsWith("0")){
						colDestText = Integer.parseInt(DestData.substring(1));
						if(colDestText==colSrcText){
                    		 flag++;
                    		 if(flag==0){
                    		 for(short colu = 0;colu<srcColCount;colu++){
                    			 			      
                    			 
                    			 
                    			srcData=getvalue(readSrcWorkBook , rows,cols, colu);
         						if(srcData.startsWith("0")){
         							colSrcText = Integer.parseInt(srcData.substring(1));
         						}
         						else{
         							colSrcText1 =srcData.substring(1);
         						}
                    			 
                    			 DestData=getvalue(readDestWorkBook , row,cols, colu);
         						if(DestData.startsWith("0")){
         							colDestText = Integer.parseInt(DestData.substring(1));
         							if(colSrcText==colDestText){		
     		         				}else{
     		         					FileOutputStream srcFileOut =  new FileOutputStream(fileName);
     		         					System.out.println("Match Not Found at Row: " +rows +" Source text: " +colSrcText +" Destination text: " +colDestText);
     		         					HSSFCellStyle style = writeScrWorkBook.createCellStyle();
     		       					HSSFCell testColor =writeScrWorkBook.getSheetAt(0).getRow(rows).getCell(colu);
     		       					style.setFillForegroundColor(HSSFColor.RED.index);
     		       					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
     		       					testColor.setCellStyle(style);
     		       					writeScrWorkBook.write(srcFileOut);
     		       					srcFileOut.close();
     		         					
     		         				}
         						}
         						else{
         							colDestText1 =DestData.substring(1);
         							 if(colSrcText1.equalsIgnoreCase(colDestText1)){		
         		         				}else{
         		         					FileOutputStream srcFileOut =  new FileOutputStream(fileName);
         		         					System.out.println("Match Not Found at Row: " +rows +" Source text: " +colSrcText1 +" Destination text: " +colDestText1);
         		         				
         		         					HSSFCellStyle style = writeScrWorkBook.createCellStyle();
	         		       					HSSFCell testColor =writeScrWorkBook.getSheetAt(0).getRow(rows).getCell(colu);
	         		       					style.setFillForegroundColor(HSSFColor.RED.index);
	         		       					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	         		       					testColor.setCellStyle(style);
	         		       					writeScrWorkBook.write(srcFileOut);
	         		       					srcFileOut.close();
         		         				}
         						}
                    			                     			                    			 
                    		 }
                    		 }
                    		 else{
                    			 System.out.println("no or duplicate entries exist");
                    		 }
                    		 
                    		 
                    	 }
					}
					else{
						colDestText1 =DestData.substring(1);
						 if(colDestText1.equals(colSrcText1)){
                    		 
                    		 
                    	 	 if(colDestText==colSrcText){
	                    		 flag++;
	                    		 if(flag==1){
	                    			 if(srcColCount>destColCount){
	                    			 counter =destColCount;
	                    			 }
	                    			 else{
	                    			counter = srcColCount;	 
	                    			 }
	                    		 for(short colu = 0;colu<counter;colu++){
	                    			 
		                    			srcData=getvalue(readSrcWorkBook , rows, cols,colu);
		         						if(srcData.startsWith("0")){
		         							colSrcText = Integer.parseInt(srcData.substring(1));
		         						}
		         						else{
		         							colSrcText1 =srcData.substring(1);
		         						}
	                    			 				   
		         						 
		                    			DestData=getvalue(readDestWorkBook , row, cols,colu);
		         						if(DestData.startsWith("0")){
		         							colDestText = Integer.parseInt(DestData.substring(1));
		         							 if(colSrcText==colDestText){		
		         		         				}else{
		         		         					FileOutputStream srcFileOut =  new FileOutputStream(fileName);
		         		         					System.out.println("Match Not Found at Row: " +rows +" Source text: " +colSrcText +" Destination text: " +colDestText);
		         		         					HSSFCellStyle style = writeScrWorkBook.createCellStyle();
			         		       					HSSFCell testColor =writeScrWorkBook.getSheetAt(0).getRow(rows).getCell(colu);
			         		       					style.setFillForegroundColor(HSSFColor.RED.index);
			         		       					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			         		       					testColor.setCellStyle(style);
			         		       				writeScrWorkBook.write(srcFileOut);
			         		       			srcFileOut.close();
		         		         					
		         		         				}
		         						}
		         						else{
		         							colDestText1 =DestData.substring(1);
		         							if(colSrcText1.equalsIgnoreCase(colDestText1)){		
	         		         				}else{
	         		         					FileOutputStream srcFileOut =  new FileOutputStream(fileName);
	         		         					System.out.println("Match Not Found at Row: " +rows +" Source text: " +colSrcText1 +" Destination text: " +colDestText1);
	         		         					HSSFCellStyle style = writeScrWorkBook.createCellStyle();
		         		       					HSSFCell testColor =writeScrWorkBook.getSheetAt(0).getRow(rows).getCell(colu);
		         		       					style.setFillForegroundColor(HSSFColor.RED.index);
		         		       					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		         		       				testColor.setCellStyle(style);
		         		       				writeScrWorkBook.write(srcFileOut);
	         		         				srcFileOut.close();
	         		         				}
		         						}
	                    			 }
	                    		 }
	                    		 else{
	                    			 System.out.println("no or duplicate entries exist");
	                    		 }
	                    				                    		
	                    	 }
	                    	 else{				                    		 
	                    	 }
                    	 }
                    	 else{
                    		 
                    	 }
					}
					
					
				}
			}
			}
		}
	}
	
	
	
	
	
	
	public static void main(String[] args) {
		
		//int i=0;
		
		try {
		/*	FileInputStream srcfile = new FileInputStream(new File("D:\\Tax 1099 Report_as_scr.xls"));
			FileInputStream destfile = new FileInputStream(new File("D:\\Tax 1099 Report_as_dest.xls"));*/
			// Inputs from the Sheet and UI.
		/*
		 * String sourceFile ="";
		 * String destinationFile= "";
		 * Boolean Execute= true;
		 * Boolean Header = true
		 * short toleranceLevel = 0;
		 * ArrayList<String> toleranceColumn[] = new ArrayList<String>
		 */	
			
			
			
			
			
			
			
			
			Path rn_demo = Paths.get("D:\\", "Tax 1099 Report_as_scr.xls");
			InputStream srcfile = Files.newInputStream(rn_demo);
			Path rn_demo1 = Paths.get("D:\\", "Tax 1099 Report_as_dest.xls");
			InputStream destfile = Files.newInputStream(rn_demo1);
			int srcRowCount =0;
			int srcColCount=0;
			String Pk = "Trade ID";
			HSSFWorkbook readSrcWorkBook = new HSSFWorkbook(srcfile);
			HSSFWorkbook readDestWorkBook = new HSSFWorkbook(destfile);
				
			 /*rn_demo = Paths.get("D:\\", "workbook.xls");
			InputStream inputStream = Files.newInputStream(rn_demo);
			 writeDestWorkBook = new HSSFWorkbook(inputStream);*/
			
			srcRowCount = readSrcWorkBook.getSheetAt(0).getPhysicalNumberOfRows();
			srcColCount = readSrcWorkBook.getSheetAt(0).getRow(1).getLastCellNum();
			int destRowCount = readDestWorkBook.getSheetAt(0).getPhysicalNumberOfRows();
			int destColCount = readDestWorkBook.getSheetAt(0).getRow(1).getLastCellNum();//GETROW(1) = HEADER START
			String srcData= null;
			String DestData= null;
			//Check the number of rows and columns in the source and destination files.
			
			if(srcRowCount!=destRowCount){
				System.out.println("Number of rows do not match.\n number of rows in source file: "+srcRowCount+"\nnumber of rows in destination file: "+destRowCount);
			}
			if(srcColCount!=destColCount){
				System.out.println("Number of columns do not match.\n number of columns in source file: "+srcColCount+"\nnumber of columns in destination file: "+destColCount);
			}
			//STATIC
		
			checkColumns(readSrcWorkBook,readDestWorkBook);//,writeDestWorkBook);
			checkColumns(readDestWorkBook,readSrcWorkBook);//,writeDestWorkBook);
			
			//FileOutputStream srcFileOut =  new FileOutputStream(name);
			
				System.out.println("PrimaryKey: "+Pk);
				String sourceFileMod="D:\\SourceFileModified.xls";		
				String destFileMod="D:\\DestFileModified.xls";
				
				compare(readSrcWorkBook, readDestWorkBook, Pk, sourceFileMod);
				compare(readDestWorkBook, readSrcWorkBook, Pk, destFileMod);
				
				
				
				
				/*checkColumns(readSrcWorkBook,readDestWorkBook);//,writeDestWorkBook);
				checkColumns(readDestWorkBook,readSrcWorkBook);//,writeDestWorkBook);
				//Check if the column names are same :
								
				
				//srcFileOut.close();
				// fileOut = new FileOutputStream("D:\\workbook.xls");
				//writeDestWorkBook.write(fileOut);
				/*writeDestWorkBook.write(fileOut);	*/
		}catch (Exception e) {
			e.printStackTrace();
		}	
	}
}
