package main;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import Writer.Writer;

public class main {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		Writer write = new Writer();
		
		write.RemoveRows(write.getSheet());
//		write.write("03.2019");
		

//		write.write("04.2019"); // error because rows are empty for this month
		write.write("02.2019");

			
		
		System.out.println("Done without errors");
		
		

	}

}
