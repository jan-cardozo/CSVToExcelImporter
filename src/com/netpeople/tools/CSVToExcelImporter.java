package com.netpeople.tools;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.FileAlreadyExistsException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVToExcelImporter {

	public static void main(String[] args) {
		try {
			File in = new File(args[0]);
			if (!in.canRead()) {
				throw new FileNotFoundException();
			}

			File out = new File(args[1]);
			if (!out.createNewFile()) {
				throw new FileAlreadyExistsException(null);
			}
			FileOutputStream outStream = new FileOutputStream(out);

			BufferedReader reader = new BufferedReader(new InputStreamReader(
					new FileInputStream(in)));

			Workbook workbook = new XSSFWorkbook();
			
			Sheet sheet = workbook.createSheet();
			workbook.setSheetName(0, "Report");
			
			String line = new String();
			int i = 0;
			while ((line = reader.readLine()) != null) {
				Row row = sheet.createRow(i);
				String[] values = line.split("\\|");
				int j = 0;
				
				for (String value : values) {
					Cell cell = row.createCell(j);
					cell.setCellValue(value);
					j++;
				}
				i++;
			}
			
			workbook.write(outStream);
			reader.close();
			outStream.close();
		} catch (FileAlreadyExistsException e) {
			System.err.println("The output file could not be created, perhaps it is already there?");
			System.exit(1);
		} catch (FileNotFoundException e) {
			System.err.println("The input file could not be read");
			System.exit(2);
		} catch (IOException e) {
			System.err.println("There was an IO error.");
			System.exit(3);
		}
		System.exit(0);
	}

}
