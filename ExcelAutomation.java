package pratik;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelAutomation {

    @Test
    public void test() throws IOException {

        String path="src/test/java/resources/ulkeler.xlsx";
        FileInputStream fileInputStream = new FileInputStream(path);
        Workbook workbook= WorkbookFactory.create(fileInputStream);

        //- 1.satirdaki 2.hucreye gidelim ve yazdiralim
        Cell satir1Hucre2=workbook.getSheet("Sayfa1").getRow(0).getCell(1);
        System.out.println(satir1Hucre2);


        //- 1.satirdaki 2.hucreyi bir string degiskene atayalim ve yazdiralim
        String satir1Hucre2Str= satir1Hucre2.getStringCellValue();
        System.out.println(satir1Hucre2Str);

        //- 2.satir 4.cell’in afganistan’in baskenti oldugunu test edelim
        Cell satir2Cell4=workbook.getSheet("Sayfa1").getRow(1).getCell(3);
        Assert.assertEquals(satir2Cell4.getStringCellValue(),"Kabil","eşit degil");

        //- Satir sayisini bulalim
        System.out.println(workbook.getSheet("Sayfa1").getLastRowNum());//190

        //- Fiziki olarak kullanilan satir sayisini bulun
        System.out.println(workbook.getSheet("Sayfa1").getPhysicalNumberOfRows());//191

        //- Ingilizce Ulke isimleri ve baskentleri bir map olarak kaydedelim
        Map<String,String> map= new HashMap<>();
        String key="";
        String value="";
        for (int i = 0; i <=workbook.getSheet("Sayfa1").getLastRowNum() ; i++) {

            key=workbook.getSheet("Sayfa1").getRow(i).getCell(0).getStringCellValue();
            value=workbook.getSheet("Sayfa1").getRow(i).getCell(1).getStringCellValue();
            map.put(key,value);
        }

        System.out.println(map);


    }
}
