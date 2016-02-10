package pisual.momo.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import pisual.bank.model.Unit;

public class WriteFilePoi {
	public void WriteFilePoi(List<Unit> unitList) throws IOException {
		int excelNum = unitList.size();
		for(int j=0;j<excelNum;j++)
		{
		InputStream fe = new FileInputStream(unitList.get(j).getSourceFile().getPath());
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fe);
		HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
		int totalNum = unitList.get(j).getSourceLocation().size();
		double sun = 0;
		for(int i=0;i<totalNum;i++)
		{
			sun=sun+unitList.get(j).getSourceLocation().get(i).getResult();
		}
		 HSSFRow row = hssfSheet.getRow(unitList.get(j).getTargetLocationY());
		 HSSFCell cell = row.createCell(unitList.get(j).getTargetLocationX());
		 cell.setCellValue(sun);
		 OutputStream out = new FileOutputStream(unitList.get(j).getSourceFile().getPath());
		 hssfWorkbook.write(out);
		 out.close();
		}
	}
	}