package pisual.momo.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import pisual.bank.model.Unit;

public class ReadFilePoi {
	/**
	 * 
	 * 
	 * **/
	public void ReadFilePoi(Unit unit) throws IOException {
		int totalNum = unit.getSourceLocation().size();
		for(int i=0;i<totalNum;i++)
		{
			InputStream fe = new FileInputStream(unit.getSourceLocation().get(i).getFile().getPath());
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fe);
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
			HSSFRow hssfRow = hssfSheet.getRow(unit.getSourceLocation().get(i).getLocationY());	
			HSSFCell date = hssfRow.getCell(unit.getSourceLocation().get(i).getLocationX());
			System.out.println(this.getValue(date));
			unit.getSourceLocation().get(i).setResult(Double.parseDouble(this.getValue(date)));
		}
	}
	
	/**
	 * 
	 * **/
	 private String getValue(HSSFCell hssfCell) {
	        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
	            return String.valueOf(hssfCell.getBooleanCellValue());
	        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
	            return String.valueOf(hssfCell.getNumericCellValue());
	        } else {
	            return String.valueOf(hssfCell.getStringCellValue());
	        }
	    }
}
