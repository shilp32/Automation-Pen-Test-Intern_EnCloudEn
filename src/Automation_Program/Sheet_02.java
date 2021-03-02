package Automation_Program;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Sheet_02
{
	public static String data( int r, int c)
	{
		try
		{
			FileInputStream file = new FileInputStream(new File("./Data/input.xlsx"));
			
			Workbook wb = WorkbookFactory.create(file);
			
			Sheet sh = wb.getSheet("sheet_04");
			DataFormatter da = new DataFormatter();
			String data =  da.formatCellValue(sh.getRow(r).getCell(c));
			return data;
			
			
		}
		catch (Exception e) 
		{
			return null;
		}
	}
	
	public static void main(String[] args) throws Exception
	{
		FileInputStream file = new FileInputStream(new File("./Data/input.xlsx"));

		Workbook wb = WorkbookFactory.create(file);

		Sheet sh = wb.createSheet("Sheet_02");
		int rowid = 0;

		for(int i =0; i<9; i++)
		{
			Row row = sh.createRow(rowid++);
			int cellid = 0;
			for(int j=0; j<9; j++)
			{
				String da = data(j, i);
				System.out.print(da+"\t");
				row.createCell(cellid++).setCellValue(da);

			}
			System.out.println();
			
		}
		FileOutputStream fos = new FileOutputStream(new File("./Data/input.xlsx"));
		wb.write(fos);
		fos.close();
		System.out.println("don");
		
	}

}
