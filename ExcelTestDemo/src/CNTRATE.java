import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Vector;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
public class CNTRATE {
	public static void main(String[] args) {
		String fileName = "F:\\MyDocs\\workspace\\TestExcel\\xls\\CNTRATE.xls";
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

			String sql = "INSERT INTO EDS_CNTRATE ( "
					+ "TESTCODE,  "
					+ "RATE,"
					+ "DEPTCODE, "
					+ "COLCOM1, "
					+ "COLCOM2"
					+ ")		"
					+

					"VALUES (";
					
					
					
			
			if ((((HSSFCell)cellStoreVector.elementAt(0)).getNumericCellValue())== -9999) {
				sql +=  "";
			} else {
				sql +=  ((HSSFCell) cellStoreVector.elementAt(0)).getNumericCellValue();
			}
			
			if ((((HSSFCell)cellStoreVector.elementAt(1)).getNumericCellValue())== -9999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(1)).getNumericCellValue();
			}
			if ((((HSSFCell)cellStoreVector.elementAt(2)).getNumericCellValue())== -9999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(2)).getNumericCellValue();
			}
			if ((((HSSFCell)cellStoreVector.elementAt(3)).getNumericCellValue())== -9999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(3)).getNumericCellValue();
			}
			if ((((HSSFCell)cellStoreVector.elementAt(4)).getNumericCellValue())== -9999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(4)).getNumericCellValue();
			}
			if ((((HSSFCell)cellStoreVector.elementAt(5)).getNumericCellValue())== -9999) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(5)).getNumericCellValue();
			}
			sql += ");" ;
			System.out.println(sql);
			
			
		}
	}
}
