package Tests;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

import Reader.ExcelReader;

class Read_from_file_tests {

	@Test
	void test() throws EncryptedDocumentException, InvalidFormatException, IOException {


		ExcelReader reader = new ExcelReader();
		
		reader.ReadHowManySheetsInIzasFile();
		
	
		
	}

}
