package Automation_Program;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Sheet_01 
{
	public static void main(String[] args) throws Exception
	{
		FileInputStream file = new FileInputStream(new File("./Data/input.xlsx"));

		Workbook wb = WorkbookFactory.create(file);

		Sheet sh = wb.createSheet("Sheet_01");
		int rowid = 0;
		for(int i=1; i<=10; i++)
		{ 
			Row row = sh.createRow(rowid++);
			int cellid = 0;
			for(int j=2; j<=10; j++)
			{
				int res = i*j;
				System.out.print(res+"\t");
				row.createCell(cellid++).setCellValue(res);
			}
			System.out.println();
		}
		FileOutputStream fos = new FileOutputStream(new File("./Data/input.xlsx"));
		wb.write(fos);
		fos.close();
		System.out.println("don");


	}

}
