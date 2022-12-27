package _JDBC.Gun2;

import _JDBC.JDBCParent;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

// Soru: Actor table'ındaki tüm verileri yeni excel'e yazdırınız

public class _04_Soru extends JDBCParent {

    @Test
    public void test1() throws SQLException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("ActorData");

        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("actor_id");
        row.createCell(1).setCellValue("first_name");
        row.createCell(2).setCellValue("last_name");
        row.createCell(3).setCellValue("last_update");

        ResultSet rs = statement.executeQuery("SELECT * FROM actor LIMIT 20");

        int count = 1;
        while (rs.next())
        {
            int actor_id = rs.getInt("actor_id");
            String first_name = rs.getString("first_name");
            String last_name = rs.getString("last_name");
            String last_update = rs.getString("last_update");

            row= sheet.createRow(count++);

            row.createCell(0).setCellValue(actor_id);
            row.createCell(1).setCellValue(first_name);
            row.createCell(2).setCellValue(last_name);
            row.createCell(3).setCellValue(last_update);
        }
        try {
            FileOutputStream outputStream =
              new FileOutputStream("src/test/java/ApachePOI/resource/ActorData.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (Exception e) {
            System.out.println(e);
        }
        }
}





