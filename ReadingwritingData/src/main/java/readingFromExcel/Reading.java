package readingFromExcel;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading {

	public static void main(String[] args) throws Exception {

		
		
		String path="C:\\TestFolder\\SampleExcel.XlSX.xlsx";
	
		// Create an object of file class and pass the path

				File src = new File(path); // File is a class

				// Pass this source to the file input stream

				FileInputStream fis = new FileInputStream(src); // It will convert  excel data to binary data

				XSSFWorkbook workbook = new XSSFWorkbook(fis); //creat an object of the workbook

				XSSFSheet sheet = workbook.getSheetAt(0);

				for (Row row : sheet) { // Ehhanced for loop.it will go and read rows

					for (Cell cell : row) {

						System.out.print(cell.getStringCellValue() + "\t");

					}

					System.out.println();

				}
				workbook.close();
				fis.close();

	}

}
