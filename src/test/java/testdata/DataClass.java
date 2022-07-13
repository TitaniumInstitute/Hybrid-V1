package testdata;

import static com.ti.framework.utils.data.ExcelUtils.getExcelTableArray;

import com.ti.framework.config.PropertyManager;
import java.lang.reflect.Method;
import org.testng.annotations.DataProvider;

public class DataClass {
  static Object[][] testObjArray;
  private static String excelWorkBook = PropertyManager.getInstance().getProperty("DataEngine");
  private static String excelValidWorkSheet = PropertyManager.getInstance().getProperty("TestCaseSheet");

  @DataProvider(name="ExcelData")
  public static Object[][] getDataFromExcel(Method method){
    testObjArray = getExcelTableArray(excelWorkBook, excelValidWorkSheet);
    return testObjArray;
  }
}
