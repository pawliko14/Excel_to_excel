package Writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
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
public class Writer {
	
	  
	    private  List<excel_object> objects = new ArrayList<>();
	    
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
	    
	    
	    
	    public void write(String data)
	    {				  
	            try {
	            	System.out.println("                                ");           	
	            	System.out.println("--------------------------------- ");           	

	            	System.out.println("take data from reader: ");           	
	            	ExcelReader read = new ExcelReader();            	
	            	objects =read.Read_from_file(data);
	            	
	    		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
	    			// TODO Auto-generated catch block
	    			e.printStackTrace();
	    		}
	            
	          
	           
	            
	            //saving data to given file works fine, but syntax do not allow to manipulate over excel_object object list
	            try {
	                FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFile()));
	                Workbook workbook = WorkbookFactory.create(inputStream);	     
	                Sheet sheet = workbook.getSheetAt(0);          
	                
	                
	                
	            //    List<Cell> cells =  GetLastRow_new_version_3rd_version();
	                
	                int rowCount = GetAllWrittenRows(sheet, workbook);            
	   
       
	                // create rows based on data taken form excelreader
	                for(int j = 0 ;  j < objects.size()-1 ; j++)
	                {
	                	Row row = sheet.createRow(rowCount++);
	                	
	                	int columnCount  = 0;
	                	
	                	  Cell cell = row.createCell(columnCount);
	                      cell.setCellValue(rowCount);
	                      
	                      Cell cell_0 = row.createCell(columnCount++);
	                      if(objects.get(j).getCountry() instanceof String)
	                      {
	                    	  cell_0.setCellValue((String) objects.get(j).getCountry() );
	                    	  GetCellStyle(cell_0,workbook);
	                      }
	                      
	                      Cell cell_1 = row.createCell(columnCount++);
	                      if(objects.get(j).getClient() instanceof String)
	                      {
	                    	  cell_1.setCellValue((String) objects.get(j).getClient() );
	                    	  GetCellStyle(cell_1,workbook);

	                      }
	                      
	                      Cell cell_2 = row.createCell(columnCount++);
	                      if(objects.get(j).getMachine_type() instanceof String)
	                      {
	                    	  cell_2.setCellValue((String) objects.get(j).getMachine_type() );
	                    	  GetCellStyle(cell_2,workbook);

	                      }
	                      
	                      Cell cell_3 = row.createCell(columnCount++);
	                      if(objects.get(j).getSN() instanceof String)
	                      {
	                    	  cell_3.setCellValue((String) objects.get(j).getSN() );
	                    	  GetCellStyle(cell_3,workbook);

	                      }
	                      
	                      Cell cell_4 = row.createCell(columnCount++);
	                      if(objects.get(j).getQuantity() instanceof String)
	                      {
	                    	  cell_4.setCellValue((String) objects.get(j).getQuantity() );
	                    	  GetCellStyle(cell_4,workbook);

	                      }
	                      
	                      Cell cell_5 = row.createCell(columnCount++);
	                      if(objects.get(j).getDate() instanceof String)
	                      {
	                    	  cell_5.setCellValue((String) objects.get(j).getDate() );
	                    	  GetCellStyle(cell_5,workbook);

	                      }
	                      
	                      Cell cell_6 = row.createCell(columnCount++);
	                      if(objects.get(j).getYear() instanceof String)
	                      {
	                    	  cell_6.setCellValue((String) objects.get(j).getYear() );
	                    	  GetCellStyle(cell_6,workbook);

	                      }
	                      
	                      Cell cell_7 = row.createCell(columnCount++);
	                      if(objects.get(j).getValue_EUR() instanceof String)
	                      {
	                    	  cell_7.setCellValue((String) objects.get(j).getValue_EUR() );
	                    	  GetCellStyle(cell_7,workbook);

	                      }
	                      
	                      Cell cell_8 = row.createCell(columnCount++);
	                      if(objects.get(j).getValue_PLN() instanceof String)
	                      {
	                    	  cell_8.setCellValue((String) objects.get(j).getValue_PLN() );
	                    	  GetCellStyle(cell_8,workbook);

	                      }
	                      
	                      Cell cell_9 = row.createCell(columnCount++);
	                      if(objects.get(j).getKurs_EUR() instanceof String)
	                      {
	                    	  cell_9.setCellValue((String) objects.get(j).getKurs_EUR() );
	                    	  GetCellStyle(cell_9,workbook);

	                      }                                
	                }
	                
	                

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
 
	public static void GetCellStyle(Cell cell, Workbook wb)
	{
        CellStyle style = wb.createCellStyle();  
        
        style.setAlignment(HorizontalAlignment.CENTER);
        
        style.setBorderBottom(BorderStyle.THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());  
        
        
        style.setBorderRight(BorderStyle.THIN);  
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());  
        
        style.setBorderTop(BorderStyle.THIN);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());  
        
        style.setBorderLeft(BorderStyle.THIN);
        
     //   style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
    //    style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        
        cell.setCellStyle(style);  
	}
	
	
	public static void removeRow(Sheet sheet, int rowIndex) {
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
	
	public void RemoveRows(Sheet sheet)
	{
    	// remove last rows tests:
        for(int i = 0 ; i < 9999; i++)
        	removeRow( sheet,  2672 + i); 
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
	
	public static List<Cell> GetLastRow() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		int HowManyCellsInTheRow = 10;
		int SpaceBettwenLastRow = 2;
		int SheetNumeber = 0;
		
		
		System.out.println("GetLastRow begin: ");
        
    	List<String> ListOfCellsInRow = new ArrayList<String>();		
    	List<CellStyle> CellStyleList = new ArrayList<CellStyle>();
    	List<Cell> Cells = new ArrayList<Cell>();
  	
		 FileInputStream inputStream = new FileInputStream(new File(Parameters.getPathToIrekFileBackup()));
		 	Workbook workbook = WorkbookFactory.create(inputStream);


        Sheet sheet = workbook.getSheetAt(SheetNumeber); 


        DataFormatter formatter = new DataFormatter(); 
        int rowCount = sheet.getLastRowNum();
        
        
        // get row 
        for(int i = 0 ; i < HowManyCellsInTheRow ; i++)
        {
        	String value_of_cell = formatter.formatCellValue(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i));
 	
        	ListOfCellsInRow.add(value_of_cell);
        	
        	CellStyle newCellStyle = workbook.createCellStyle();
        	if(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i) != null)
        		newCellStyle.cloneStyleFrom(workbook.getSheetAt(SheetNumeber).getRow(rowCount).getCell(i).getCellStyle());
        		CellStyleList.add(newCellStyle);
        		    	
        }
        

        // remowe row
        removeRow(sheet, rowCount);
        
      
        //insert to sheet
        Row row = sheet.createRow(sheet.getLastRowNum()+ SpaceBettwenLastRow); 
        int columnCount = 0;
        for(int i = 0 ; i < ListOfCellsInRow.size(); i++)
        {
      	  Cell cell = row.createCell(columnCount++);   
      	  cell.setCellStyle(CellStyleList.get(i));
          cell.setCellFormula( ListOfCellsInRow.get(i) );   
        	  
        	  Cells.add(cell);
        }
        
        
        PrintCellInArray(Cells);

        
        System.out.println("End of function");
        
        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream(Parameters.getPathToIrekFileBackup());
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        
        
		return Cells;             
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