package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

public class TestData {
	
	public ArrayList<String> getdata(String TestcaseName) throws IOException
	{
        ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C://Users//HARDIK RAJPUT//OneDrive//Desktop//DataDriven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int SheetSize = workbook.getNumberOfSheets();
		
		for(int i=0;i<SheetSize;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1"))
			{
		XSSFSheet sheet = workbook.getSheetAt(i);
		//Scan the entire first row to get cell from row which has testcase
		
		            Iterator<Row> Rows =sheet.rowIterator();
		            Row Firstrow = Rows.next();
		            Iterator<Cell> Cells = Firstrow.cellIterator();
		            int k =0;
		            int column = 0;
		            while(Cells.hasNext())
		            {
		            	Cell Value = Cells.next();
		            	if(Value.getStringCellValue().equalsIgnoreCase("TestCase"))
		            	{
		            		column=k;
		            	}
		            	k++;
		            }
		            System.out.println(column);
		            while(Rows.hasNext())
		            {
		            	Row r=Rows.next();
		            	if(r.getCell(column).getStringCellValue().equalsIgnoreCase(TestcaseName))
{
	Iterator<Cell> cv= r.cellIterator();
	while(cv.hasNext())
	{
		Cell c = cv.next();
		if(c.getCellType()==CellType.STRING)
		{
		a.add(c.getStringCellValue());
		}
		else
		{
			
			a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
			
		}
		
	}
}
		            	
		            }
			}
			}
		return a;
		
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	}

}
