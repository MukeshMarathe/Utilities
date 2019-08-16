package org.zebra.io;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertExcelInOCC {
	public static final String PROPERTIES_FILE_PATH = "C:\\\\Users\\\\Mukesh\\\\Documents\\\\workspace-spring-tool-suite-4-4.2.0.RELEASE\\\\ZebraUtility\\\\Resource\\\\config.properties";
	public static final String EXCEL_FILE_PATH = "C:\\Users\\Mukesh\\Desktop\\Zebra\\profile.xlsx";
	public static final String CSV_FILE_PATH = "C:\\Users\\Mukesh\\Desktop\\Zebra\\profile.csv";
	
	public static void main(String[] args) throws Exception {
		readExcel(EXCEL_FILE_PATH);
	}

	public static String getColumnName(String propertyName) {
		FileReader reader;
		String pName = null;
		try {
			reader = new FileReader(PROPERTIES_FILE_PATH);
			Properties properties = new Properties();
			properties.load(reader);

			pName = properties.getProperty(propertyName);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return pName;
	}

	
	  public static void readExcel(String excelFilePath) throws Exception {
		  FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
		  Workbook workbook = new XSSFWorkbook(excelFile);
		  org.apache.poi.ss.usermodel.Sheet datatypeSheet = workbook.getSheetAt(0);
		  Iterator<Row> iterator = datatypeSheet.iterator();
		  int count =0;
	
		  FileWriter csvWriter = new FileWriter(CSV_FILE_PATH); 
		  while (iterator.hasNext()) { 
			  Row currentRow = iterator.next(); 
			  Iterator<Cell>cellIterator = currentRow.iterator();
			  while (cellIterator.hasNext()) { 
				  Cell currentCell = cellIterator.next(); 
				  
			  		if(currentCell.getCellTypeEnum() == CellType.STRING) {
			  			System.out.print(currentCell.getStringCellValue());
			  			if(count==0) {
			  				String cName = getColumnName(currentCell.getStringCellValue());
			  				csvWriter.append(cName);
			  			}else {
			  				csvWriter.append(currentCell.getStringCellValue());
			  			}
			  			
			  		}else if(currentCell.getCellTypeEnum() == CellType.NUMERIC) {
			  			if(count==0) {
			  				String cName = getColumnName(currentCell.getStringCellValue());
			  				csvWriter.append(cName);
			  			}else {
			  				System.out.print(String.valueOf(currentCell.getNumericCellValue()));
			  				csvWriter.append(String.valueOf(currentCell.getNumericCellValue()));
			  			}
			  		}else if(currentCell.getCellType() == CellType.BOOLEAN) {
			  			System.out.print(currentCell.getBooleanCellValue());
		  				csvWriter.append(String.valueOf(currentCell.getBooleanCellValue()));
		  				
			  		}else if(currentCell.getCellType() == CellType.BLANK) {
			  			System.out.println("");
		  				csvWriter.append(null);
			  		}
			  		
			  		if(cellIterator.hasNext()) {
		  				csvWriter.append(",");
		  			}
			  	}
		  		System.out.println("\n");
		  		csvWriter.append("\n");
		  		count ++;
		  		System.out.println(count);
		  	}
		  csvWriter.flush(); 
		  csvWriter.close();
		  }
}