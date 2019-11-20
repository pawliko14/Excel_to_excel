package Reader;



import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {
	
    public static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\el08\\Desktop\\Irek_exel\\WYSY£KI MASZYN.xls";
    public static final String PATH_TO_FOLDER = "C:\\Users\\el08\\Desktop\\Irek_exel\\"; 
    private static List<excel_object> objects = new ArrayList<>();

    public static void main(String[] args) throws IOException, InvalidFormatException {
    	
    	Read_from_file("02.2019");

    }
    
    // enum are not in use yet, since its not necessary in this etap
    // anyway should be used for better code review
    
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
    
    public enum IzasExcelColumnName
    {
    	LP(0),
    	CUSTOMER(1),
    	TYPE_OF_LATHE(2),
    	SN(3),
    	QTY_EUR(4),
    	VALUE(5),
    	DATE_OF_SHIPMENT(6),
    	QTY_PLN(7),
    	VALUE_OF_THE_INVOICE_PLN(8),
    	UWAGI(9);
    	
    	public final int value; 	
    	private IzasExcelColumnName(int value)
    	{
    		this.value =value;
    	}
    }
      
    
    public static List<String> ReadHowManySheetsInIzasFile() throws EncryptedDocumentException, InvalidFormatException, IOException
    {
    	List<String> ListOfSheets = new ArrayList<String>();
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        
        for(Sheet sheet : workbook)
        {
            ListOfSheets.add(sheet.getSheetName());
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
    
    public static List<excel_object> Read_from_file(String dateMonth_to_retrive) throws IOException, EncryptedDocumentException, InvalidFormatException
    {
  	
    	
    	List<String> ListOfSheets = ReadHowManySheetsInIzasFile();
    	
    	PrintStream fileout = new PrintStream(PATH_TO_FOLDER + "LogFile.txt");
    	

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        int index_of_sheet_02_2019 = 0;
     //   String  temp_02_2019= "02.2019";
        
       for(int i = 0 ; i < ListOfSheets.size() ; i++)
       {
    	   if(ListOfSheets.get(i).equals(dateMonth_to_retrive))
    			   index_of_sheet_02_2019 = i;
       }
       
       int sheet_number = index_of_sheet_02_2019;


       // set out to file
     //  	System.setOut(fileout);

 
        int biggestValue_previous = 0 ; // it means how many machines has been sold in this month
        int biggestValue = 0 ; // it means how many machines has been sold in this month

        DataFormatter formatter = new DataFormatter();
        
        // print all A column from sheet [number] 
        // range between (3;19) is necessary to get proper data from excel file, while data starts with 3rd row to 19 ( 20nd row is summary)
        for(int i = 3 ; i < 19 ; i++)
        {
           
            // if next cell IS NOT EMPTY
            if(!formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(0)).isEmpty() )
            {
            	            	
	            biggestValue_previous   = Integer.parseInt(formatter.formatCellValue((workbook.getSheetAt(sheet_number).getRow(i-1).getCell(0))));
	            biggestValue 			= Integer.parseInt(formatter.formatCellValue((workbook.getSheetAt(sheet_number).getRow(i).getCell(0))));
	            
	          if(biggestValue_previous> biggestValue)
	        	biggestValue =biggestValue_previous;
            
            }
        }
                
        // print all depends on numeber of rows:
        for(int i = 2 ; i < biggestValue + 2  ;i++)
        {  	 
       	  // get substring of full date ( get only year)
	       	String year = "";
	       	if (formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.DATE.value)).length() > 4) 
	       	     year  = formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.DATE.value)).substring(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.DATE.value)).length() - 4);
	       	
	       	
	       	// trim cell for example DORMAC(NL) from Izas sheet to NL, because country is needed
	       	String country = getStringBetweenBrackets(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.CLIENT.value)));

       	 	
	       	// creating excel object based on inner class 'Builder'
      	 	excel_object object = new excel_object.Builder()
       	 			.Country(country)
       	 			.Client(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.CLIENT.value)))
       	 			.Machine_type(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.MACHINE_TYPE.value)))
       	 			.SN(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.SN.value)))
       	 			.Quantity(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.QUANTITY.value)))
       	 			.Date(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.DATE.value)))
       	 			.Year(year)
       	 			.Value_EUR(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.VALUE_EUR.value)))
       	 			.Value_PLN(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.VALUE_PLN.value)))
       	 			.Kurs_EUR(formatter.formatCellValue(workbook.getSheetAt(sheet_number).getRow(i).getCell(IrekExcelColumnName.KURS_EUR.value)))
       	 			.build();
      	 	
      
       	 	
       	 		objects.add(object);
        }
        
        //print all object from the list
        System.out.println("all objects in the arraylist: ");
        
//        for(excel_object ob : objects)
//        {
//        	ob.printObject();
//        }
	
        // Closing the workbook
        workbook.close();
        
		return objects;
    }
    
}
