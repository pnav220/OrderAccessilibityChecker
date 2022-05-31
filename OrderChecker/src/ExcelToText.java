import java.util.*;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.BufferedWriter;
import java.io.BufferedReader;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class ExcelToText {
	File filename;
	public ExcelToText(File filename){
		//default
		this.filename = filename;
	}
	public static void main(String[] args) {
		
	}
	public void dumpCodes(File filename) {
		try {
			//File file = new File("SaleOffer - S01687.xlsx");
			FileInputStream input = new FileInputStream(filename);
			XSSFWorkbook excel = new XSSFWorkbook(input);
			XSSFSheet sheet = excel.getSheetAt(1);
			BufferedWriter writer = new BufferedWriter(new FileWriter("output.txt"));
			Iterator<Row> rowItr = sheet.iterator();
			while(rowItr.hasNext()) {
				Row row = rowItr.next();
				Iterator<Cell> cellItr = row.cellIterator();
				while(cellItr.hasNext()) {
				Cell cell = cellItr.next();
				if(cell.getColumnIndex() == 4) {
				switch(cell.getCellTypeEnum()) {
				case STRING:
					writer.write(cell.getStringCellValue()+"\n");
					break;
				case NUMERIC:
					writer.write(cell.getNumericCellValue()+"\n");
					break;
				default:
					writer.write("");
				}
			}
			}
		}
			writer.close();
			excel.close();
			input.close();
	}
		catch (Exception e) {
			e.printStackTrace();
		}
	} 
	public void checkStatus() throws EncryptedDocumentException, InvalidFormatException {
		//csv format is code;status
		String line = "";
		String split = ";";
		boolean check = false;
		ArrayList<String> SICodes = new ArrayList<String>();
		ArrayList<String> status = new ArrayList<String>();
		try {
			BufferedReader reader = new BufferedReader(new FileReader("Sales Items Status.csv"));
			BufferedReader reader2 = new BufferedReader(new FileReader("output.txt"));
			BufferedWriter writer = new BufferedWriter(new FileWriter("outputStatus.txt"));
			while((line = reader.readLine())!=null){
				String[] splitLine = line.split(split);
				SICodes.add(splitLine[0]);
				if(splitLine.length == 2)
					status.add(splitLine[1]);
			}
			while((line = reader2.readLine())!= null){
				if(line.compareTo("SI Code") == 0) {
				}
				else {
					for(int i = 0; i < SICodes.size(); i++) {
						if(line.compareTo(SICodes.get(i))==0) {
							writer.write(line +": "+status.get(i) + "\n");
							check = true;
							break;
						}
					}
					if(check != true)
						writer.write(line +": N/A\n");
					
				}
			}
			reader.close();
			reader2.close();
			writer.close();
			BufferedReader reader3 = new BufferedReader(new FileReader("output.txt"));
			FileInputStream input = new FileInputStream("Orderability Status Check Input Form_V1.0.xlsm");
			XSSFWorkbook excel = new XSSFWorkbook(input);
			XSSFSheet sheet = excel.getSheetAt(0);
			int rowIndex = 3;
			while((line = reader3.readLine())!=null) {
				if(line.equalsIgnoreCase("SI Code")) {
					//nothing happens
				}
				else {
					Row row = sheet.getRow(rowIndex);
					Cell cell = row.getCell(0);
					cell.setCellValue(line);
				}
				rowIndex++;
			}
			input.close();
			FileOutputStream output = new FileOutputStream("Orderability Status Check Input Form_V1.0.xlsm");
			excel.write(output);
			excel.close();
			output.close();
			Runtime.getRuntime().exec("/Users/pnav/Library/CloudStorage/OneDrive-Nokia/myVBS.vbs");
			/*Iterator<Row> rowItr = sheet.iterator();
			while(rowItr.hasNext()) {
				Row row = rowItr.next();
				Iterator<Cell> cellItr = row.cellIterator();
				while(cellItr.hasNext()) {
				Cell cell = cellItr.next();
				if(cell.getColumnIndex() == 0) {
				switch(cell.getCellTypeEnum()) {
				case STRING:
					writer.write(cell.getStringCellValue()+"\n");
					break;
				case NUMERIC:
					writer.write(cell.getNumericCellValue()+"\n");
					break;
				default:
					writer.write("");
				}
			}
			}
		}*/
		}
		catch (IOException e) {
			e.printStackTrace();
		}
	}
}
