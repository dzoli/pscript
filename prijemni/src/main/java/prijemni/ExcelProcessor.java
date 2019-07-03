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
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;

public class ExcelProcessor {

	public static void main(String[] args) {
		String inPath = args[0];
		String sfrDel = args[1];
		System.out.println("in = " + inPath);
		System.out.println("sfrdel = " + sfrDel);
		Row delRow = null;
		Row firstStudRow = null;
		HSSFWorkbook wb = null;
		
		try (OutputStream os = new FileOutputStream("workbook.xls")) {
			DataFormatter dataFormatter = new DataFormatter();

			FileInputStream fis = new FileInputStream(inPath);
			wb = new HSSFWorkbook(fis);
			HSSFSheet sheet = wb.getSheet("RangListPMF");
			Iterator<Row> rows = sheet.rowIterator();
			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();

				while (cells.hasNext()) {
					Cell cell = cells.next();
					String cellVal = dataFormatter.formatCellValue(cell);
					if (cellVal.equals(sfrDel)) {
						delRow = row;
					}
					if (cellVal.equals("1.")) {
						firstStudRow = row;
					}
				}
			}

			System.out.println("== deleting student " + sfrDel + " at row " + delRow.getRowNum());
			int lastRow = sheet.getLastRowNum();
			sheet.removeRow(delRow);
			sheet.shiftRows(delRow.getRowNum(), lastRow, -1);
			sheet.shiftRows(delRow.getRowNum(), lastRow, -1);

			// row number
			int idx = delRow.getRowNum();
			while (sheet.getRow(idx) != null && sheet.getRow(idx).getCell(1) != null) {
				Row row = sheet.getRow(idx);

				idx += 2;
				String cellVal = dataFormatter.formatCellValue(row.getCell(1));
				cellVal = cellVal.substring(0, cellVal.indexOf('.'));
				row.getCell(1).setCellValue("" + (Integer.valueOf(cellVal) - 1) + ".");
			}

			// row color
			idx = firstStudRow.getRowNum();
			boolean rowToFill = false;
			int isGrey = 0;
			while (sheet.getRow(idx) != null && sheet.getRow(idx).getCell(1) != null) {
				Row row = sheet.getRow(idx);
				idx += 2;
				
				Iterator<Cell> cells = row.cellIterator();
				while (cells.hasNext()) {
					Cell cell = cells.next();
					CellAddress ca = cell.getAddress();
					String cellVal = dataFormatter.formatCellValue(cell);
					System.out.println(cellVal);
					if (cellVal.equals("1.")) {
						rowToFill = true;
					}
					if (rowToFill) {
						if (ca.getColumn() >= 1 && ca.getColumn() <= 9) {
							if ((isGrey % 2) == 0) {
								System.out.println("grey " + ca.getColumn());
								CellStyle cs = wb.createCellStyle();
								cs.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
								cs.setFillPattern(FillPatternType.BIG_SPOTS);
								cell.setCellStyle(cs);
							} else {
								System.out.println(isGrey + " no fil ");
								CellStyle cs = wb.createCellStyle();
								cs.setFillPattern(FillPatternType.NO_FILL);
								cell.setCellStyle(cs);
							}
						} else {
							CellStyle cs = wb.createCellStyle();
							cs.setFillPattern(FillPatternType.NO_FILL);
							cell.setCellStyle(cs);
						}
					}
				}
				isGrey ++;
			}

			wb.write(os);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
