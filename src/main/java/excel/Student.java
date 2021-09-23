package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Student extends StudentData{
private	XSSFSheet sheet;
private XSSFWorkbook workbook;
private FileInputStream file;
Student() throws IOException {
	
	file=new FileInputStream("C:\\Users\\LENOVO\\Desktop\\NewExcel\\newExcelp.xlsx");
	workbook=new XSSFWorkbook(file);
sheet=workbook.getSheet("sheet3");
}

@Override
public void student() {
	Iterator iterator=sheet.iterator();
	while(iterator.hasNext()) {
	XSSFRow row=(XSSFRow) iterator.next();
	Iterator cellIterator=row.cellIterator();
	while(cellIterator.hasNext()) {
	XSSFCell cell=	(XSSFCell) cellIterator.next();
	switch(	cell.getCellType()) {
	case STRING:System.out.print(cell.getStringCellValue());
	break;
	case NUMERIC:System.out.print(cell.getNumericCellValue());
	break;
	case BOOLEAN:System.out.print(cell.getBooleanCellValue());
	}
		
	}
	System.out.println(" ");
	}
	
}
public static void main(String args[]) throws IOException {
	StudentData d=new Student();
	d.student();
}
}
