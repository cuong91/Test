package tranning_them;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Test {
public static void main(String[] args) {
	HSSFWorkbook wb = new HSSFWorkbook();
	HSSFSheet sheet = wb.createSheet("FirstSheet");
	HSSFRow rowhead = sheet.createRow(0); 
	HSSFCellStyle style = wb.createCellStyle();
	HSSFFont font = wb.createFont();
	font.setFontName(HSSFFont.FONT_ARIAL);
	font.setFontHeightInPoints((short)10);
	font.setBold(true);
	style.setFont(font);
	rowhead.createCell(0).setCellValue("ID");
	rowhead.createCell(1).setCellValue("First");
	rowhead.createCell(2).setCellValue("Second");
	rowhead.createCell(3).setCellValue("Third");
	for(int j = 0; j<=3; j++)
	rowhead.getCell(j).setCellStyle(style);
}
}
