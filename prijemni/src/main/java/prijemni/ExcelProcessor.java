package prijemni;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelProcessor {
//	String outPath = System.getProperty("user.home");
//	System.out.println("== out home == " + outPath);

	public static void main(String[] args) {
		String inPath = args[0];
		String sfrDel = args[1];
		Row delRow = null;
		
		try (OutputStream os = new FileOutputStream("workbook.xls")) {
	        DataFormatter dataFormatter = new DataFormatter();
			
			FileInputStream fis = new FileInputStream(inPath);
			HSSFWorkbook wb = new HSSFWorkbook(fis);
			HSSFSheet sheet = wb.getSheet("RangListPMF");
			Iterator<Row> rows =sheet.rowIterator();	
			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();
				
				while (cells.hasNext()) {
					Cell cell = cells.next();
					String cellVal = dataFormatter.formatCellValue(cell);
					if (cellVal.equals(sfrDel)) {
						delRow = row;
					}
				}
			}
			
			System.out.println("== deleting student " + sfrDel + " at row " + delRow.getRowNum());
			int lastRow = sheet.getLastRowNum();
			sheet.removeRow(delRow);
			sheet.shiftRows(delRow.getRowNum(), lastRow, -1);
			sheet.shiftRows(delRow.getRowNum(), lastRow, -1);
			
			int idx = delRow.getRowNum();
			while(sheet.getRow(idx) != null && sheet.getRow(idx).getCell(1) != null) {
				Row row = sheet.getRow(idx);
				idx+=2;
				String cellVal = dataFormatter.formatCellValue(row.getCell(1));
				System.out.println("cell val bf =" + cellVal);
				
				cellVal = cellVal.substring(0, cellVal.indexOf('.'));
				
				System.out.println("cell val =" + cellVal);
				row.getCell(1).setCellValue("" + (Integer.valueOf(cellVal) - 1) + ".");
			}
			
			wb.write(os);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void whiteCell(Workbook wb, Row row, String txt) {
		CellStyle style = wb.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
		style.setFillPattern( FillPatternType.SOLID_FOREGROUND );
		Cell cell = row.createCell((short) 1);
		cell.setCellValue(txt);
		cell.setCellStyle(style);
	}
}
