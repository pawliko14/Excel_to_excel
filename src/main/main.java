package main;

import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import Reader.ExcelReader;
import Writer.Writer;
import Writer.Writer_dubg;
import Writer.Writer_tester;


/*
 * 
 * Tasks to do next week :
 * 
 * - separete code in writer to  : wrtier,row remower and row adder
 * - add unit tests for every funcitonality
 * 
 * 
 * 
 * 
 * 
 */


public class main {
	
	private static List<String> DatesStaringOn2019;
	private static List<String> FinalListOfDates;

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException {
	
	
		
		// get all data bettwen 01.12.2018 till now()
		GetListOfMonth();
		
		//initialize writing section
		Writer_dubg write = new Writer_dubg();

		// get last row of the summary
		write.GetLastRow_Copy_it_then_remove_from_sheet();

		// remove last row from the sheet
		write.RemoveRows();
		
		
		
		/*
		 *  MAIN FUNCTION, ITERATION ARE BASED ON LIST OF ELEMENT THAT ARE GRATER THAN 12.2018 ( OLDER DATA IS INCLUDED TO THE SHEET)
		 *  
		 *  insertion rows from Izas excel to Irek
		 */	
	for(int i = 0 ;i < FinalListOfDates.size();i++)
			write.write(FinalListOfDates.get(i));
		
	
		// write again copied row to specific position
		write.InsertToTheSheet();
		
		

		
		System.out.println("Done without errors");
		

	}
	
	
	
	private static void GetListOfMonth()  throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException
	{
		FinalListOfDates = new ArrayList<>();
		
		SimpleDateFormat f = new SimpleDateFormat("MM.yyyy");	
		ExcelReader read = new ExcelReader();
		
		DatesStaringOn2019 = read.ReadHowManySheetsInIzasFile();
		
		for(int i = 0;  i < DatesStaringOn2019.size() ; i++)
		{		
			if( !DatesStaringOn2019.get(i).contains("Arkusz"))
			{				
				Date date = f.parse("12.2018");			
				Date date2 = f.parse( DatesStaringOn2019.get(i).substring(DatesStaringOn2019.get(i).length()-7, DatesStaringOn2019.get(i).length()));
				
				if(date2.after(date))					
					FinalListOfDates.add(DatesStaringOn2019.get(i));			
												
			}		
		}		
	}

}
