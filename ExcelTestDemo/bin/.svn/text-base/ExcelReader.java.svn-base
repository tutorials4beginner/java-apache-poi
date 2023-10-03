

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Vector;

public class ExcelReader {

	public static void main(String[] args) throws Exception {

		String filename = "C:\\Test\\CHM_PATH.xls";
		FileInputStream fis = null;

		
		String TST_PARMT, REMARKS, TST_UNIT, SEX, REFRNG;
		int TST_CODE, AGE_FROM, AGE_TO, DAY_1, DAY_2;
		float RNG_FROM, RNG_TO;
		try {
			fis = new FileInputStream(filename);
			System.out.println("Hello....1");	
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			System.out.println("Hello....1");
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			System.out.println("Hello....sheet"+sheet);
			Iterator rowIter = sheet.rowIterator();
			
			while (rowIter.hasNext()) {
				HSSFRow myRow = (HSSFRow) rowIter.next();
				Iterator cellIter = myRow.cellIterator();
				Vector<String> cellStoreVector = new Vector<String>();
				
				while (cellIter.hasNext()) {
					HSSFCell myCell = (HSSFCell) cellIter.next();
					String cellvalue = myCell.getStringCellValue();
					cellStoreVector.addElement(cellvalue);
				}

				int i = 0;
				TST_CODE = Integer.parseInt(cellStoreVector.get(i));
				TST_PARMT = cellStoreVector.get(i + 1).toString();
				REMARKS = cellStoreVector.get(i + 2).toString();
				TST_UNIT = cellStoreVector.get(i + 3).toString();
				AGE_FROM = Integer.parseInt(cellStoreVector.get(i + 4)
						.toString());
				AGE_TO = Integer
						.parseInt(cellStoreVector.get(i + 5).toString());
				SEX = cellStoreVector.get(i + 6).toString();
				RNG_FROM = Integer.parseInt(cellStoreVector.get(i + 7)
						.toString());
				RNG_TO = Integer
						.parseInt(cellStoreVector.get(i + 8).toString());
				DAY_1 = Integer.parseInt(cellStoreVector.get(i + 9).toString());
				DAY_2 = Integer
						.parseInt(cellStoreVector.get(i + 10).toString());
				REFRNG = cellStoreVector.get(i + 11).toString();

				insertQuery(TST_PARMT, REMARKS, TST_UNIT, SEX, REFRNG,
						TST_CODE, AGE_FROM, AGE_TO, DAY_1, DAY_2, RNG_FROM,
						RNG_TO);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fis != null) {
				fis.close();
			}
		}
		// showExelData(sheetData);
	}

	private static void insertQuery(String TST_PARMT, String REMARKS,
			String TST_UNIT, String SEX, String REFRNG, int TST_CODE,
			int AGE_FROM, int AGE_TO, int DAY_1, int DAY_2, float RNG_FROM,
			float RNG_TO) {
		System.out.println(TST_PARMT + " " + REMARKS + " " + TST_UNIT + " "
				+ SEX + " " + REFRNG + " " + TST_CODE + " " + AGE_FROM + " "
				+ AGE_TO + " " + DAY_1 + " " + DAY_2 + " " + RNG_FROM + " "
				+ RNG_TO);
	}
}