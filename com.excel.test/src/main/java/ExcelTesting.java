

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTesting {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		String timeStamp = new SimpleDateFormat("yyyy-MM-dd HHmmss").format(Calendar.getInstance().getTime());
		String fileName = "C:\\Users\\Or.Garfunkel\\Desktop\\Test\\" + timeStamp + ".xlsx";		
		String [] lineList = new String[]{"Application id","Application name","marketPlace user","submission email","publisher email","support email","marketPlace password","Price","Application status", "Message code","Validation Result"};
		String [] listForCreatingone = new String[]{"123","My App","userName","email","publisher email","support email"," password","0.99","Application status", "Message code","Pending"};
		String [] listForCreating = new String[]{"123","My App","userName","email","publisher email","support email"," password","0.99","Application status", "Message code","Approved"};
		String [] listForCreatingTwo = new String[]{"123","My App","userName","email","publisher email","support email"," password","0.99","Application status", "Message code","Pending"};
		try {
			createExcel(fileName);
			createHeaderLine(lineList, fileName);
			createBodyLine(listForCreatingone, fileName);
			createBodyLine(listForCreating, fileName);
			createBodyLine(listForCreatingTwo, fileName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println(fileName);
	}

	
	/**
	 * The function creates excel file
	 * */
	public static void createExcel(String filename) throws IOException {
	     XSSFWorkbook workbook=new XSSFWorkbook();
	     /*XSSFSheet sheet =*/  workbook.createSheet("FirstSheet");  
	     FileOutputStream fileOut =  new FileOutputStream(filename);
	     workbook.write(fileOut);
	     fileOut.close();
	}
	
	/**
	 * The function creates Header line in the excel file
	 * */
	public static void createHeaderLine(String [] listForCreating, String filename) throws IOException { 
	    FileInputStream file = new FileInputStream(new File(filename));
	    XSSFWorkbook workbook = new XSSFWorkbook(file);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    
	    int rowNumber = sheet.getLastRowNum()+1;
	    if(rowNumber == 1)
	    	rowNumber = 0;
	    sheet.createFreezePane(0, 1);

	    XSSFRow rowhead=   sheet.createRow((short)rowNumber);

	    XSSFFont font= workbook.createFont();
	    font.setFontHeightInPoints((short)12);
	    font.setFontName("Arial");
	    font.setColor(IndexedColors.AQUA.getIndex());
	    font.setBold(true);
	    font.setItalic(false);

	    CellStyle style = workbook.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setAlignment(CellStyle.ALIGN_CENTER);
	    style.setFont(font);	    
	    
	    for(int i=0 ; i < listForCreating.length ; i++) {
	    	 Cell cell = rowhead.createCell((short) i);
	    	 cell.setCellValue(listForCreating[i]);
	    	 cell.setCellStyle(style);
	    	 sheet.autoSizeColumn(i);
	    }
	    
	    FileOutputStream fileOut =  new FileOutputStream(filename);
	    workbook.write(fileOut);
	    fileOut.close();
	}
	
	/**
	 * The function creates line in the excel file
	 * */
	public static void createBodyLine(String [] listForCreating, String filename) throws IOException { 
	    FileInputStream file = new FileInputStream(new File(filename));
	    XSSFWorkbook workbook = new XSSFWorkbook(file);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    
	    int rowNumber = sheet.getLastRowNum()+1;

	    CellStyle style = workbook.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    
	    XSSFRow rowhead=   sheet.createRow((short)rowNumber);
	    
	    for(int i=0 ; i < listForCreating.length ; i++) {
	    	if(listForCreating[i].equals("Approved"))
	    	{
		    	 Cell cell = rowhead.createCell((short) i);
		    	 cell.setCellValue(listForCreating[i]);
		    	 style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		    	 cell.setCellStyle(style);
	    	}
	    	else if(listForCreating[i].equals("Pending"))
	    	{
	    		Cell cell = rowhead.createCell((short) i);
		    	 cell.setCellValue(listForCreating[i]);
		    	 style.setFillForegroundColor(IndexedColors.RED.getIndex());
		    	 cell.setCellStyle(style);
	    	}
	    	else{
	    		rowhead.createCell((short) i).setCellValue(listForCreating[i]);
	    	}
	    }
	    FileOutputStream fileOut =  new FileOutputStream(filename);
	    workbook.write(fileOut);
	    fileOut.close();
	}
	
//	private XSSFCellStyle setHeaderStyle(XSSFWorkbook sampleWorkBook)
//	{
//		XSSFFont font = sampleWorkBook.createFont();
//		font.setFontName("Arial");
//		font.setColor(IndexedColors.DARK_BLUE.getIndex());
//		font.setBold(true);
//		HSSFCellStyle cellStyle = sampleWorkBook.createCellStyle();
//		cellStyle.setFont(font);
//		return cellStyle;
//	}

	
}
