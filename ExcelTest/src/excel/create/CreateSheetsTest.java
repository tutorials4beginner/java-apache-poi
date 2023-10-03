package excel.create;
import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class CreateSheetsTest {

  /**
   * @param args
   * @throws IOException 
   * @throws IOException 
   * @throws WriteException 
   * @throws BiffException 
   */
  public static void main(String[] args) throws IOException, WriteException {
    //Creates a writable workbook with the given file name
    WritableWorkbook workbook = Workbook.createWorkbook(new File("D:/MySheets.xls"));
    
    //Create sheets with given name
    for (int i = 0 ; i < 6 ; i ++) {
      workbook.createSheet("My Sheet-" + (i+1), i);
    }
    
    //Writes out the data held in this workbook in Excel format
    workbook.write(); 
    
    //Close and free allocated memory 
    workbook.close(); 
  }

}