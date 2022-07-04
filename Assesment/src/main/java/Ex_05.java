import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;


public class Ex_05

{
    public static void main(String[] args) {

        String path=System.getProperty("user.dir")+"\\Book.xlsx";
        try {
            readingFromExcel(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void readingFromExcel(String path) throws IOException {
        File excelFile=new File(path);
        FileInputStream fis=new FileInputStream(excelFile);
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        XSSFSheet sheet=workbook.getSheetAt(0);

        Iterator<Row> rowIt=sheet.iterator();
        while (rowIt.hasNext()){
            Row row=rowIt.next();

            Iterator<Cell> cellIterator=row.cellIterator();
            while (cellIterator.hasNext()){
                Cell cell= cellIterator.next();

                System.out.print(cell.toString()+ "\s");
            }
            System.out.println();
        }
        workbook.close();
        fis.close();
    }
}

