package excel.insert;
import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class InsertRowTest {

  /**
   * @param args
   * @throws IOException 
   * @throws IOException 
   * @throws WriteException 
   * @throws BiffException 
   */
  public static void main(String[] args) throws IOException, WriteException {
    //Creates a writable workbook with the given file name
    WritableWorkbook workbook = Workbook.createWorkbook(new File("D:/InsertRow.xls"));
    WritableSheet sheet = workbook.createSheet("My Sheet", 0);
    
    // Create cell font and format
    WritableFont cellFont = new WritableFont(WritableFont.TIMES, 16);
    cellFont.setColour(Colour.BLUE);
    
    WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
    cellFormat.setBackground(Colour.ORANGE);
    cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
    
    Label lable = null;
    int insertRow = 10;
    
    for (int i = 0 ; i < 4 ; i ++) {
      sheet.insertRow(insertRow); //Insert new row at row 2 [Moves the row 2 to next row]
      lable = new Label(0, insertRow, " " + i , cellFormat);
      sheet.addCell(lable);
    }
    
    //Writes out the data held in this workbook in Excel format
    workbook.write(); 
    
    //Close and free allocated memory 
    workbook.close(); 
  }

}