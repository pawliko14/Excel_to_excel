package main;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

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

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		
		
		//working good
		Writer_dubg write = new Writer_dubg();
		
		write.RemoveRows(write.getSheet());
		
		write.write("02.2019");
	
		//what happened to 05.2019 sheet? is it coruppted or what
//		write.write("05.2019");
		
//		write.write("03.2019");
								// error on 04.2019
	//	write.write("06.2019"); 
		
//		Writer_dubg writer2 = new Writer_dubg();
//		writer2.write("03.2019");
		
		//to tests  <- not working
//		Writer_tester write = new Writer_tester();
//		write.RemoveRows(write.getSheet());
//		
//		write.write("02.2019");
		
		System.out.println("Done without errors");
		
		

	}

}
