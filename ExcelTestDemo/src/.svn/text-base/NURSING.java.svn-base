import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
public class NURSING {
	public static void main(String[] args) {
		String fileName = "F:\\MyDocs\\workspace\\TestExcel\\xls\\NURSNG.xls";
		Vector dataHolder = ReadCSV(fileName);
		printCellDataToConsole(dataHolder);
	}

	public static Vector ReadCSV(String fileName) {
		
		Vector cellVectorHolder = new Vector();

		try {
			FileInputStream myInput = new FileInputStream(fileName);
			POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
			HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
			HSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator rowIter = mySheet.rowIterator();

			while (rowIter.hasNext()) {
				HSSFRow myRow = (HSSFRow) rowIter.next();
				Iterator cellIter = myRow.cellIterator();
				Vector cellStoreVector = new Vector();
				while (cellIter.hasNext()) {
					HSSFCell myCell = (HSSFCell) cellIter.next();
					cellStoreVector.addElement(myCell);
				}
				cellVectorHolder.addElement(cellStoreVector);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return cellVectorHolder;
	}

	private static void printCellDataToConsole(Vector dataHolder) {
		for (int i = 0; i < dataHolder.size(); i++) {
			Vector cellStoreVector = (Vector) dataHolder.elementAt(i);

			String sql = "INSERT INTO EDS_NURSING ( "
					+ "NURSCODE,  "
					+ "NURS_NAME,"
					+ "ADDR1, "
					+ "ADDR2, "
					+ "ADDR3, "
					+ "DISCSTAT, "
					+ "OPTPRN, "
					+ "FOOTPRN "
					+ ")		"
					+

					"VALUES (";
					
					
					
			
			if (cellStoreVector.elementAt(0) == null) {
				sql +=  "";
			} else {
				sql +=  ((HSSFCell) cellStoreVector.elementAt(0)).getNumericCellValue();
			}
			
			if (((HSSFCell)cellStoreVector.elementAt(1)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(1)).getStringCellValue() + "'";
			}

			if (((HSSFCell)cellStoreVector.elementAt(2)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(2)).getStringCellValue() + "'";
			}
			if (((HSSFCell)cellStoreVector.elementAt(3)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(3)).getStringCellValue() + "'";
			}
			if (((HSSFCell)cellStoreVector.elementAt(4)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(4)).getStringCellValue() + "'";
			}
			if ((((HSSFCell)cellStoreVector.elementAt(5)).getNumericCellValue())== -999999999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(5)).getNumericCellValue();
			}
			if (((HSSFCell)cellStoreVector.elementAt(6)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(6)).getStringCellValue() + "'";
			}
			if (((HSSFCell)cellStoreVector.elementAt(7)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(7)).getStringCellValue() + "'";
			}
			sql += ");" ;
			System.out.println(sql);
			
			
		}
	}
}
