import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class BACTR {
	public static void main(String[] args) {
		String fileName = "F:\\MyDocs\\workspace\\TestExcel\\xls\\BACTR.xls";
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

			String sql = "INSERT INTO EDS_BACTR ( "
					+ "BACT_CODE,  "
					+ "BACT_NAME"
					+ ")		"
					+

					"VALUES ("
					+

					"EDS_BACTR_SEQ.NEXTVAL ";
			
			
			sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(1)).getStringCellValue() + "'";
			sql += ");" ;
			
			System.out.println(sql);
		}
	}
}
