package exelparser;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XmlTransform {
	public static void main(String[] args) throws IOException {

		File excel = new File("/home/sven/Apps/SWWerke.xlsx");
		String savePath = "/home/sven/Apps/SWWerke.xml";
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheetAt(0);
		StringBuilder sb = new StringBuilder();
		
		
		modsBuildHeader(sb);

		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow myrow = sheet.getRow(i);
			sb.append("<mods:titleInfo xml:lang=" + "\"de\"");
			for (Cell mycell : myrow) {
				if (mycell.getCellTypeEnum() == CellType.NUMERIC) {
					mycell.setCellType(CellType.STRING);
					sb.append(" ID=" + "\"" + mycell + "\">\n");
				} else {
					sb.append("<mods:title>" + mycell + "</mods:title>\n");
				}
			}
			sb.append("</mods:titleInfo>\n");
		}
		sb.append("</mods:mods>");

		System.out.println(sb.toString());
		
		
		saveFile(sb, savePath);

		/*
		 * Iterator<Row> itr = sheet.iterator(); while (itr.hasNext()) { Row row
		 * = itr.next(); // Iterating over each column of Excel file
		 * Iterator<Cell> cellIterator = row.cellIterator(); while
		 * (cellIterator.hasNext()) { Cell cell = cellIterator.next();
		 * 
		 * switch (cell.getCellType()) { case Cell.CELL_TYPE_STRING:
		 * System.out.print(cell.getStringCellValue() + "\t"); break; case
		 * Cell.CELL_TYPE_NUMERIC: System.out.print(cell.getNumericCellValue() +
		 * "\t"); break; case Cell.CELL_TYPE_BOOLEAN:
		 * System.out.print(cell.getBooleanCellValue() + "\t"); break; default:
		 * 
		 * } } System.out.println(""); }
		 */
	}

	public static void saveFile(StringBuilder sb, String path) throws IOException {
		File file = new File(path);
		try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
		    writer.write(sb.toString());
		}
	}

	public static void modsBuildHeader(StringBuilder sb) {
		sb.append("<mods:mods"+ " xmlns:mods=" + "\"http://www.loc.gov/mods/v3\"" + " xmlns:xlink=" +
		"\"http://www.w3.org/1999/xlink\"" + " xmlns:xsi=" + "\"http://www.w3.org/2001/XMLSchema-instance\"" +
				" version=" +"\"3.6\">\n");
	}
}



