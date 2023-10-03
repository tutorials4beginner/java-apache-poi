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

public class Dept {
	public static void main(String[] args) {
		String fileName = "D:\\MyDocs\\Workspace\\repos\\trunk\\EdsDiag\\sql\\CollectedData\\dept_data.xls";
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

			String sql = "INSERT INTO EDS_DEPT ( "
					+ "DEPT_ID,  "
					+ "DEPT_NAME, "
					+ "DEPT_NOTE, "
					+ "PARENT_ID, "
					+ "DEPT_LEVEL "
					+ ")		"
					+

					"VALUES (EDS_DEPT_SEQ.NEXTVAL,";	
			if (cellStoreVector.elementAt(0) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(0)).getStringCellValue() + "'";
			}
			if (cellStoreVector.elementAt(1) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(1)).getStringCellValue() + "'";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(2) == null) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(2)).getNumericCellValue() + "";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(3) == null) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(3)).getNumericCellValue() + "";
			}
			
			sql += ");" ;
			System.out.println(sql);
			
			

		}
	}
}
