package main;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import Writer.Writer;
import Writer.Writer_dubg;

public class main {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		Writer_dubg write = new Writer_dubg();
		
		write.RemoveRows(write.getSheet());
		write.write("03.2019");
		
		Writer_dubg writer2 = new Writer_dubg();
//		write.write("04.2019"); // error because rows are empty for this month
		writer2.write("02.2019");
		
		Writer_dubg writer3 = new Writer_dubg();

		writer3.write("06.2019");
		
		Writer_dubg writer4= new Writer_dubg();

		writer4.write("07.2019");
		
		System.out.println("Done without errors");
		
		

	}

}
