package Writer;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import Reader.ExcelReader;
import Reader.excel_object;
import parameters.Parameters;


/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 *
 * @author www.codejava.net
 *
 */
public class Writer_dubg {
	
	    private  List<excel_object> objects = new ArrayList<>();
	    private static int LastRowWithOldData =2652-1;
	    private static List<Cell> Cells = new ArrayList<Cell>(); // list of copied cells from the last row
		private static List<String> ListOfCellsInRow = new ArrayList<String>();		
    	private static List<CellStyle> CellStyleList = new ArrayList<CellStyle>();

	   
	    public Sheet getSheet() throws EncryptedDocumentException, InvalidFormatException, IOException
	    {
	    	 FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFile()));
             Workbook workbook = WorkbookFactory.create(inputStream);	     
             Sheet sheet = workbook.getSheetAt(0); 
	    	
	    	return sheet;
	    }
	    

	    
	    public int GetAllWrittenRows(Sheet sheet, Workbook wb)
	    {
	    	int notNullCount = 0;
	    	 sheet = wb.getSheetAt(0);
	    	for (Row row : sheet) {
	    	    for (Cell cell : row) {
	    	        if (cell.getCellType() != Cell.CELL_TYPE_BLANK) {
	    	            if (cell.getCellType() != Cell.CELL_TYPE_STRING ||
	    	                cell.getStringCellValue().length() > 0) {
	    	                notNullCount++;
	    	                break;
	    	            }
	    	        }
	    	    }
	    	}
	    	return notNullCount;
	    }
	    
	    
	    
	    public void write(String data) throws ParseException
	    {			
	    	boolean EmptyList = false;
	    	
	            try {
	            	System.out.println("                                ");           	
	            	System.out.println("--------------------------------- ");  
	            	System.out.println("take data from reader: ");   
	            	
	            	objects.clear();
	            	ExcelReader read = new ExcelReader();            	
	            	objects = read.Read_from_file(data);
	            	
	            	if(objects.size() == 0 )
	            	{
	            		EmptyList =true;
	            	}
	            	
	            	System.out.println("objects.size: " + objects.size());
	            	
	    		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
	    			// TODO Auto-generated catch block
	    			e.printStackTrace();
	    		}
	            
	          
	           if(EmptyList == false)
	           {
	            //saving data to given file works fine, but syntax do not allow to manipulate over excel_object object list
	            try {
	                FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFile()));
	                BufferedInputStream bis = new BufferedInputStream(new FileInputStream(new File(Parameters.getPathToIrekFile())));
	                Workbook workbook = WorkbookFactory.create(bis);	     
	                Sheet sheet = workbook.getSheetAt(0);        
	                
	                System.out.println("Sheet in workbook: " + workbook.getSheetAt(0));
	                
	                
	                
	          
	                
	                int rowCount = GetAllWrittenRows(sheet, workbook);            
	   
	                System.out.println("Row count: " + rowCount);
	                
	                
	           //     CellStyle style = workbook.createCellStyle();  
	                
	                System.out.println("Objects size: " + objects.size());
	                for(int j = 0 ;  j < objects.size() ; j++)
	                {
	                	Row row = sheet.createRow(rowCount++);
	                	
	                	int columnCount  = 0;
	                	
	                	  Cell cell = row.createCell(columnCount);
	                      cell.setCellValue(rowCount);
	                      
	                      Cell cell_0 = row.createCell(columnCount++);
	                      if(objects.get(j).getCountry() instanceof String)
	                      {
	                    	  cell_0.setCellValue((String) objects.get(j).getCountry() );
	                 //   	  GetCellStyle(cell_0,workbook,style);
	                    	                  
	                      }
	                      
	                      Cell cell_1 = row.createCell(columnCount++);
	                      if(objects.get(j).getClient() instanceof String)
	                      {
	                    	  cell_1.setCellValue((String) objects.get(j).getClient() );
	             //       	  GetCellStyle(cell_1,workbook,style);

	                      }
	                      
	                      Cell cell_2 = row.createCell(columnCount++);
	                      if(objects.get(j).getMachine_type() instanceof String)
	                      {
	                    	  cell_2.setCellValue((String) objects.get(j).getMachine_type() );
	             //       	  GetCellStyle(cell_2,workbook,style);

	                      }
	                      
	                      Cell cell_3 = row.createCell(columnCount++);
	                      if(objects.get(j).getSN()  instanceof String  )
	                      {
	                    	  cell_3.setCellValue( Integer.parseInt(objects.get(j).getSN()) );
	         //           	  GetCellStyle(cell_3,workbook,style);

	                      }
	                      
	                      Cell cell_4 = row.createCell(columnCount++);
	                      if(objects.get(j).getQuantity() instanceof String)
	                      {
	                    	  double d = 0;
	                    	  
	                    	  if(objects.get(j).getQuantity().length() >=1)
	                    	  {
	                    		  Number number = NumberFormat.getInstance().parse(objects.get(j).getQuantity());                   		   
		                    	  d = number.doubleValue();
	                    	  }
	                    	  
	                    	  cell_4.setCellValue(d );
	                   // 	  GetCellStyle(cell_4,workbook,style);

	                      }
	                      
	                      Cell cell_5 = row.createCell(columnCount++);
	                      if(objects.get(j).getDate() instanceof String)
	                      {
	                    	  cell_5.setCellValue((String) objects.get(j).getDate() );
	               //     	  GetCellStyle(cell_5,workbook,style);

	                      }
	                      
	                      Cell cell_6 = row.createCell(columnCount++);
	                      if(objects.get(j).getYear() instanceof String)
	                      {
	                    	  double d = 0;
	                    	  if(objects.get(j).getYear().length() >= 4)
	                    	  {
		                    	  
		                    	  Number number = NumberFormat.getInstance().parse(objects.get(j).getYear());                   		   
		                    	  d = number.doubleValue();

	                    	  }
	                    	  cell_6.setCellValue(d);
	                //    	  GetCellStyle(cell_6,workbook,style);

	                      }
	                      
	                      Cell cell_7 = row.createCell(columnCount++);
	                      if(objects.get(j).getValue_EUR() instanceof String)
	                      {
	                    	 
	                    	  double d = 0;
	                    	  if(objects.get(j).getValue_EUR().contains(","))
	                    			  {
	                    		  		Number number = NumberFormat.getInstance().parse(objects.get(j).getValue_EUR());                   		   
	                    		   		d = number.doubleValue();
	                    			  } 
                   		 

	                    	  cell_7.setCellValue( d);
	                    //	  GetCellStyle(cell_7,workbook,style);

	                      }
	                      
	                      Cell cell_8 = row.createCell(columnCount++);
	                      if(objects.get(j).getValue_PLN() instanceof String)
	                      {
	                    	
	                    	  double d = 0;
	                    	  
	                    	  if(objects.get(j).getValue_PLN().contains(","))
	                    	  {                    		   
	                    		   Number number = NumberFormat.getInstance().parse(objects.get(j).getValue_PLN());   
	                    		    d = number.doubleValue();	   
	                    	  }
	                    	  cell_8.setCellValue(d );
	                    	//  GetCellStyle(cell_8,workbook,style);

	                      }
	                      
	                      Cell cell_9 = row.createCell(columnCount++);
	                      if(objects.get(j).getKurs_EUR() instanceof String)
	                      {
	                    	  cell_9.setCellValue((String) objects.get(j).getKurs_EUR() );
	                    //	  GetCellStyle(cell_9,workbook,style);

	                      }
	                      
	                      
	                }
	                
	                for(excel_object c : objects )
	                	 c.printObject();
	                
                         
	     
	                inputStream.close();
	                FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFile());
	                workbook.write(outputStream);
	                workbook.close();
	                outputStream.close();
	                 
	            } catch (IOException | EncryptedDocumentException
	                    | InvalidFormatException ex) {
	                ex.printStackTrace();
	            }
	          }
	        
	    }
 
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		// to tests but function need to be static
//		 Writer w = new Writer();
//		 w.RemoveRows(w.getSheet());
//		 w.write("02.2019");
		 
		// System.out.println("methods(without get) : " );
		
	}
	
	public static void GetCellStyle(Cell cell, Workbook wb, CellStyle style )
	{
         style = wb.createCellStyle();  
        
        style.setAlignment(HorizontalAlignment.CENTER);
        
        style.setBorderBottom(BorderStyle.THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());  
        
        
        style.setBorderRight(BorderStyle.THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());  
        
        style.setBorderTop(BorderStyle.THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());  
        
        style.setBorderLeft(BorderStyle.THIN);
        
        cell.setCellStyle(style);  
	}
	
	
	public static void removeRow(Sheet sheet, int rowIndex)  {
	    
          
		
		int lastRowNum=sheet.getLastRowNum();
	   
	    
	    if(rowIndex>=0 && rowIndex<lastRowNum){
	        sheet.shiftRows(rowIndex+1,lastRowNum, -1);
	    }
	    if(rowIndex==lastRowNum){
	    	Row removingRow=sheet.getRow(rowIndex);
	    	    	
	        if(removingRow!=null){
	         //   System.out.println("removed : " + rowIndex + "  ereased value : " + sheet.getRow(rowIndex).getCell(1).getStringCellValue());
	            sheet.removeRow(removingRow);
	        }
	    }
	    
	   
	}
	
	public void RemoveRows() throws EncryptedDocumentException, InvalidFormatException, IOException
	{

		  FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFile()));
		  BufferedInputStream bis = new BufferedInputStream(new FileInputStream(new File(Parameters.getPathToIrekFile())));
        Workbook workbook = WorkbookFactory.create(bis);	     
        Sheet sheet1 = workbook.getSheetAt(0);       
		
		// LastRowWithOldData -> 2672
		int lastRow = sheet1.getLastRowNum();
		
		System.out.println("last row -> " + lastRow + ", 2652 ->" + LastRowWithOldData);
		
				if(lastRow >= LastRowWithOldData)
		{
			System.out.println("lastRow > lastRowiwWithOldData");
	    	// remove last rows :
	        for(int i = lastRow ; i>= LastRowWithOldData ; i--)
	        {
	        	removeRow( sheet1,  i); 
	        }
		}
        
				bis.close();
        FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFile());
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
	}

	

	
	public void PushSavedRow(List<Cell> cells, Workbook wb, Sheet sheet) throws IOException
	{
		System.out.println("Push to last row has begun");
		
		 FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFileBackup()));

		
		sheet = wb.getSheetAt(0);
		
		   Row row = sheet.createRow(sheet.getLastRowNum()+2);
           
	        int columnCount = 0;
	        for(int i = 0 ; i < cells.size(); i++)
	        {
	      	    Cell cell = row.createCell(columnCount++);   
	      	  	cell = cells.get(i);
	        }
	        
	        inputStream.close();

	        FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFileBackup());
	        wb.write(outputStream);
	        wb.close();
	        outputStream.close();
	        
			System.out.println("End of pushing");

		
	}
	
	public void RemoveLastRow(Sheet sheet)
	{
		int rowIndex = sheet.getLastRowNum();
		
	    int lastRowNum=sheet.getLastRowNum();
	    if(rowIndex>=0 && rowIndex<lastRowNum){
	        sheet.shiftRows(rowIndex+1,lastRowNum, -1);
	    }
	    if(rowIndex==lastRowNum){
	    	Row removingRow=sheet.getRow(rowIndex);
	    	    	
	        if(removingRow!=null){
	            sheet.removeRow(removingRow);
	        }
	    }
	}
	
	
	public  List<Cell> GetLastRow_Copy_it_then_remove_from_sheet() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		int HowManyCellsInTheRow = 10;
		int SheetNumeber = 0;
		 int SpaceBettwenLastRow = 12;

		
		System.out.println("GetLastRow begin: ");
        
    
  	
		 FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFileBackup()));
		 BufferedInputStream bis = new BufferedInputStream(new FileInputStream(new File(Parameters.getPathToIrekFileBackup())));
		 	Workbook workbook = WorkbookFactory.create(bis);
		 	Sheet sheet = workbook.getSheetAt(SheetNumeber); 

        DataFormatter formatter = new DataFormatter(); 
        int rowCount = sheet.getLastRowNum();
        
    	System.out.println("rowCount : " +rowCount);

        
        
        // get row 
        for(int i = 0 ; i < HowManyCellsInTheRow ; i++)
        {
        	String value_of_cell = formatter.formatCellValue(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i));
 	
        	System.out.println("value_of_cell : " +value_of_cell);

        	
        	ListOfCellsInRow.add(value_of_cell);
        	
        	System.out.println("ListOfCellsInRow : " + ListOfCellsInRow.get(i));
        	
        	CellStyle newCellStyle = workbook.createCellStyle();
   //     	newCellStyle.setQuotePrefixed(false); // should remove  " ' " sign from excel sheet
        	if(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i) != null)
        		newCellStyle.cloneStyleFrom(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i).getCellStyle());
        	
        		CellStyleList.add(newCellStyle);
        	
        }
        

        // remowe row
        removeRow(sheet, rowCount);
        
      
//        //insert to sheet
//        Row row = sheet.createRow(sheet.getLastRowNum() + SpaceBettwenLastRow); 
//        int columnCount = 0;
//        for(int i = 0 ; i < ListOfCellsInRow.size(); i++)
//        {
//      	  Cell cell = row.createCell(columnCount++);   
//      	  cell.setCellStyle(CellStyleList.get(i));
//          cell.setCellFormula( ListOfCellsInRow.get(i) );   
//          
//        	  Cells.add(cell);
//        }
        
        
      
        
        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFileBackup());
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        
        
		return Cells;             
	}	
	
	public void InsertToTheSheet( ) throws IOException, EncryptedDocumentException, InvalidFormatException
	{
		int SheetNumeber = 0;
		 int SpaceBettwenLastRow = 3;
		
		 FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFileBackup()));
		 BufferedInputStream bis = new BufferedInputStream(new FileInputStream(new File(Parameters.getPathToIrekFileBackup())));
		 	Workbook workbook = WorkbookFactory.create(bis);
		 	Sheet sheet = workbook.getSheetAt(SheetNumeber);
		
		 	System.out.println("elements in listofcellinrow: " + ListOfCellsInRow.size());
		 	
		 	for(String s : ListOfCellsInRow)
			 	System.out.println("elements : " + s);

		 	
		 	
	     //insert to sheet
        Row row = sheet.createRow(sheet.getLastRowNum() + SpaceBettwenLastRow); 
        int columnCount = 0;
        for(int i = 0 ; i < ListOfCellsInRow.size(); i++)
        {
      	  Cell cell = row.createCell(columnCount++);   
      	//  cell.setCellStyle(CellStyleList.get(i)); // could be problematic
      	  
      	  CellStyle style = workbook.createCellStyle();
       	  style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

          style.setFillBackgroundColor(IndexedColors.BROWN.getIndex());
          style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
          
          style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
          style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
          style.setRightBorderColor(IndexedColors.BLACK.getIndex());
          style.setTopBorderColor(IndexedColors.BLACK.getIndex());
          
          style.setAlignment(HorizontalAlignment.CENTER);
      	  
      	  cell.setCellStyle(style);
      	  
          cell.setCellFormula( ListOfCellsInRow.get(i) );   
          
        	  Cells.add(cell);
        }
        
        
        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFileBackup());
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
		
		
	}
	
	
	
	private static void PrintCellInArray(List<Cell> Cells)
	{
		for(Cell c : Cells)
		{
			if(c.getCellTypeEnum() == CellType.STRING)
				System.out.println("values: " + c.getStringCellValue());
			else if(c.getCellTypeEnum() == CellType.NUMERIC)
				System.out.println("values: " + c.getNumericCellValue());
			else if(c.getCellTypeEnum() == CellType.FORMULA)
				System.out.println("values: " + c.getCellFormula());
		}
	}
	
	
	
	
	
	
	
}