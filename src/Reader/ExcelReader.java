package Reader;



import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.SystemOutLogger;

import parameters.Parameters;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {
	 
    private  List<excel_object> objects = new ArrayList<>();

    public static void main(String[] args) throws IOException, InvalidFormatException {
    	
    	//to tests
    	ExcelReader read = new ExcelReader();
    	read.Read_from_file("04.2019");

    }
        
    public enum IrekExcelColumnName
    {

    	
    	COUNTRY(0),
    	CLIENT(1),
    	MACHINE_TYPE(2),
    	SN(3),
    	QUANTITY(4),
       	VALUE_PLN(8),
    //	YEAR(6),
    	VALUE_EUR(5),
    	DATE(6), 
    	KURS_EUR(9);
    	
    	
    	public final int value;	
    	private IrekExcelColumnName(int label)
    	{
    		this.value = label;
    	}
    	
    }
    
    public enum IzaExcelColumnName
    {
    	LP(0),
    	CUSTOMER(1),
    	TYPE_OF_LATHE(2),
    	SN(3),
    	QTY(4),
    	VALUE(5),
    	DATE_OF_SHIPMENT(6),
    	QTY_2(7),
    	VALUE_OF_THE_INVOICE(8),
    	UWAGI(9);
    	
    	public final int value;
    	private IzaExcelColumnName(int label)
    	{
    		this.value = label;
    	}
    }
          
    
    public List<String> ReadHowManySheetsInIzasFile() throws EncryptedDocumentException, InvalidFormatException, IOException
    {
    	List<String> ListOfSheets = new ArrayList<String>();
        Workbook workbook = WorkbookFactory.create(new File(Parameters.getSampleXlsxFilePath()));
        
        for(Sheet sheet : workbook)
        {
        		//  "\\s" <- it means take anyting thas is a space characeter, then change it to ""
        	String SheetNameTrimmed = sheet.getSheetName().replaceAll("\\s", "");
            ListOfSheets.add(SheetNameTrimmed);
        }
        
        return ListOfSheets;

    }
    
   
    
    public int GetNumberOfSheetsInDocument(Workbook workbook)
    {
    	return workbook.getNumberOfSheets();
    }
    
    
    public static String getStringBetweenBrackets(String st)
    {
    	String string_between_Brackets = "";
    
    		if(st.contains("("))
    		{
    			String parts[] = st.split("[(]");
    			string_between_Brackets = parts[1];
    		}
    	
    		if(string_between_Brackets.contains(")"))
    		{
    			String part[] = string_between_Brackets.split("[)]");
    			string_between_Brackets = part[0];  			
    		}

 	
 
    	return string_between_Brackets;
    }
    
    public int GetIndexOfTheSheet(List<String> ListOfSheets ,String dateMonth_to_retrive )
    {
    	int sheetNumber = 0;
    	
    	 for(int i = 0 ; i < ListOfSheets.size() ; i++)
         {
      	   if(ListOfSheets.get(i).equals(dateMonth_to_retrive) || ListOfSheets.get(i) == dateMonth_to_retrive )
      	   {
      		   sheetNumber = i;
      	   }
         }
    	
    	return sheetNumber;
    }
    
    public List<excel_object> Read_from_file(String dateMonth_to_retrive) throws IOException, EncryptedDocumentException, InvalidFormatException
    {
    	int index_of_the_sheet = 0;
    	int beginning_row = 3;
    	int ending_row = 19;
    	int SheetHasAnyRecord = -1;
    	
    	List<String> ListOfSheets = ReadHowManySheetsInIzasFile();  
    	

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(Parameters.getSampleXlsxFilePath()));

        //to be sure it is truely a string
        dateMonth_to_retrive.toString();
        
        index_of_the_sheet = GetIndexOfTheSheet(ListOfSheets, dateMonth_to_retrive);
        
      
        
        int biggestValue_previous = 0 ; // it means how many machines has been sold in this month
        int biggestValue = 0 ; // it means how many machines has been sold in this month

        DataFormatter formatter = new DataFormatter();
        
        SheetHasAnyRecord =  CheckIfSheetHasAnyRecord(workbook,index_of_the_sheet, beginning_row, ending_row);
        
        System.out.println("SheetHasAnyRecord: "+ SheetHasAnyRecord);
        
        
	        // print all A column from sheet [number] 
	        // range between (3;19) is necessary to get proper data from excel file, while data starts with 3rd row to 19 ( 20nd row is summary)
	        for(int i = beginning_row ; i < ending_row ; i++)
	        {
	           
	            // if next cell IS NOT EMPTY (cell of the 1 column (not 0 ) )
	            if(!formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(1)).isEmpty() )
	            {
	            	            	
		            biggestValue_previous   = Integer.parseInt(formatter.formatCellValue((workbook.getSheetAt(index_of_the_sheet).getRow(i-1).getCell(0))));
		            biggestValue 			= Integer.parseInt(formatter.formatCellValue((workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(0))));
		            
		          if(biggestValue_previous> biggestValue)
		        	biggestValue =biggestValue_previous;
	            
	            }
	         
	        }
	        
	        if(SheetHasAnyRecord == 1)
	        {
	        	
	        	biggestValue =  1;
	        }
	        else if(SheetHasAnyRecord == 0)
	        {
	        	return objects;
	        }
	        
	        System.out.println("reader -> biggestvalue " + biggestValue);
	                
	        // print all depends on numeber of rows:
	        for(int i = 2 ; i < biggestValue + 2  ;i++)
	        {  	 
	       	  // get substring of full date ( get only year)
		       	String year = "";
		       	if (formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.DATE.value)).length() > 4) 
		       	     year  = formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.DATE.value)).substring(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.DATE.value)).length() - 4);
		       	
		       	
		       	// trim cell for example DORMAC(NL) from Izas sheet to NL, because country is needed
		       	String country = getStringBetweenBrackets(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.CLIENT.value)));
	
	       	 	
		       	// creating excel object based on inner class 'Builder'
	      	 	excel_object object = new excel_object.Builder()
	       	 			.Country(country)
	       	 			.Client(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.CLIENT.value)))
	       	 			.Machine_type(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.MACHINE_TYPE.value)))
	       	 			.SN(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.SN.value)))
	       	 			.Quantity(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.QUANTITY.value)))
	       	 			.Date(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.DATE.value)))
	       	 			.Year(year)
	       	 			.Value_EUR(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.VALUE_EUR.value)))
	       	 			.Value_PLN(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.VALUE_PLN.value)))
	       	 			.Kurs_EUR(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(IrekExcelColumnName.KURS_EUR.value)))
	       	 			.build();
	      	 	
	      
	       	 	
	       	 		objects.add(object);
	        }
	        
	      System.out.println("reader -> objectsSize " + objects.size());
        
      
        
        
        for(excel_object s : objects)
        	s.printObject();

        // Closing the workbook
        workbook.close();
        
		return objects;
    }



	private int CheckIfSheetHasAnyRecord(Workbook workbook, int index_of_the_sheet, int beginning_row, int ending_row) {
		
        DataFormatter formatter = new DataFormatter();

		int HowManyRecords = 0;
		
        Sheet sheet = workbook.getSheetAt(index_of_the_sheet);     
        Row row = sheet.getRow(beginning_row-1);
        
        int DoesSheetHasAnyFilledRow = -1;

        
        System.out.println("begin : " + beginning_row + " , end: "+ ending_row);
		
        
		for(int i = beginning_row-1 ; i < ending_row; i++)
		{
			Cell c = row.getCell(IzaExcelColumnName.CUSTOMER.value);	
			PrintCellValue(c);
			
		
			
			
			
			if(formatter.formatCellValue(workbook.getSheetAt(index_of_the_sheet).getRow(i).getCell(1)).isEmpty())
			{
				System.out.println("null , " + i);
			}
			else
			{
				System.out.println("rekord dodany ");

				HowManyRecords++;
			}
			row = sheet.getRow(i);
			
		}
		
		System.out.println("done, " + HowManyRecords);
		
		if(HowManyRecords >=2)
			DoesSheetHasAnyFilledRow =2;
		else if(HowManyRecords > 0)
			DoesSheetHasAnyFilledRow =1;
		else if(HowManyRecords ==0 )
			DoesSheetHasAnyFilledRow = 0;
		
		return DoesSheetHasAnyFilledRow;		
	}
	
	

	
	
	private void PrintCellValue(Cell c)
	{
		if(c != null) {
			
			if(c.getCellTypeEnum() == CellType.FORMULA)
				System.out.println(c.getCellFormula());
			else if(c.getCellTypeEnum() == CellType.STRING)
				System.out.println(c.getStringCellValue());
			else if(c.getCellTypeEnum() == CellType.NUMERIC)
				System.out.println(c.getNumericCellValue());		
		}
	}
	
	

    
}
