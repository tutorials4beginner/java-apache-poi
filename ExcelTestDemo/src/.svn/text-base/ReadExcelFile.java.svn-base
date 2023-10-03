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

public class ReadExcelFile {

	public static void main(String[] args) {
		String fileName = "F:\\MyDocs\\workspace\\TestExcel\\xls\\CHM_PATH.xls";
		Vector dataHolder = ReadCSV(fileName);
		printCellDataToConsole(dataHolder);
	}

	public static Vector ReadCSV(String fileName) {
		String TST_PARMT, REMARKS, TST_UNIT, AGE_FROM, AGE_TO, SEX, RNG_FROM, RNG_TO, DAY_1, DAY_2, REFRNG;

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

			String sql = "INSERT INTO EDS_CHEM_PATH_STAND ( "
					+ "TST_CODE,  "
					+ "TST_PARMT  ,"
					+ " REMARKS  , "
					+ "TST_UNIT   , "
					+ "AGE_FROM , "
					+ "AGE_TO , "
					+ "SEX  , "
					+ "RNG_FROM , "
					+ "RNG_TO , "
					+ "DAY_1 , "
					+ "DAY_2 , "
					+ "REFRNG "
					+ ")		"
					+

					"VALUES ("
					+

					"EDS_CHEM_PATH_SEQ.NEXTVAL ";
					
					
			
			if (cellStoreVector.elementAt(1) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(1)).getStringCellValue() + "'";
			}
			
			if (((HSSFCell)cellStoreVector.elementAt(2)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(2)).getStringCellValue() + "'";
			}

			if ((HSSFCell) cellStoreVector.elementAt(3) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(3)).getStringCellValue() + "'";
			}

			if ((HSSFCell) cellStoreVector.elementAt(4) == null) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(4)).getNumericCellValue() + "";
			}

			if ((HSSFCell) cellStoreVector.elementAt(5) == null) {
				sql +=  ", ";
			} else {
				sql += ", " + ((HSSFCell) cellStoreVector.elementAt(5)).getNumericCellValue() + "";
			}

			if ((HSSFCell) cellStoreVector.elementAt(6) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(6)).getStringCellValue() + "'";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(7) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(7)).getNumericCellValue() + "'";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(8) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(8)).getNumericCellValue() + "'";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(9) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(9)).getNumericCellValue() + "'";
			}
			
			if ((HSSFCell) cellStoreVector.elementAt(10) == null) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(10)).getNumericCellValue() + "'";
			}
			
			if (((HSSFCell)cellStoreVector.elementAt(11)).getStringCellValue().equals("AAAAAAAAAA")) {
				sql +=  ", ''";
			} else {
				sql += ", '" + ((HSSFCell) cellStoreVector.elementAt(11)).getStringCellValue() + "'";
			}
			
			sql += ");" ;
			System.out.println(sql);
			
			File file = new File("Result.txt");

			FileWriter writer;
			try {
				writer = new FileWriter(file, true);
				writer.write(sql);
				writer.write("\n");
				writer.flush();
				writer.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		}
	}
}