package test;

import static com.ti.framework.config.Constants.KEYWORD_FAIL;
import static com.ti.framework.config.Constants.KEYWORD_PASS;
import static com.ti.framework.config.Constants.TEST_CASE_RESULT_COL;
import static com.ti.framework.utils.extentreports.ExtentTestManager.startTest;

import com.ti.framework.config.ReadObject;
import java.lang.reflect.InvocationTargetException;
import org.apache.poi.ss.usermodel.Sheet;
import org.testng.Assert;
import org.testng.annotations.Test;
import testdata.DataClass;

public class ExecutionEngine extends Base{

  private static void executeActions() throws InvocationTargetException, IllegalAccessException {
    for (int i=0; i< methods.length;i++){
      if (methods[i].getName().equals(testStep.getActionKeyword())){
        testStepResult = (boolean) methods[i].invoke(actionKeywords, ReadObject.getInstance().getProperty(testStep.getPageObject()), testStep.getData());
        break;
      }
    }
  }

  @Test(dataProvider = "ExcelData", dataProviderClass = DataClass.class)
  public static void DriverScript(String id, String tcId, String desc, String runMode, String dataDriven, String result)
      throws InvocationTargetException, IllegalAccessException {

    Sheet dataSheet = null;
    testCaseResult = true;
    boolean tsResult = true;
    int rowStepError = 0;

    testCase.setRunMode(runMode);
    testCase.setTcKey((int)Double.parseDouble(id));
    startTest(tcId,desc);

    if (testCase.getRunMode().equalsIgnoreCase("yes")){
      if (!dataDriven.equals("")){
        dataSheet = excelReader.getWorkSheet(dataDriven);
        totalDataRows = dataSheet.getLastRowNum() - dataSheet.getFirstRowNum();
      }else{
        totalDataRows = 1;
      }

      for (int ddt = 1; ddt <= totalDataRows; ddt++){
        int ddtTotalCols = 0;

        Sheet stepsSheet = excelReader.getWorkSheet(excelTSWorkSheet);
        int totalRows = stepsSheet.getLastRowNum() - stepsSheet.getFirstRowNum();
        int totalCols = stepsSheet.getRow(0).getLastCellNum();

        for (startStepsRow = 0; startStepsRow < totalRows; startStepsRow++){
          row = stepsSheet.getRow(startStepsRow+1);
          for (startStepsCol = 0; startStepsCol < totalCols; startStepsCol ++){
            if(startStepsCol == 0 && !tcId.equals(row.getCell(startStepsCol).toString())){
              breaking = true;
              break;
            }

            switch (startStepsCol){
              case 1:
                testStep.setId(row.getCell(startStepsCol).toString());
                break;
              case 2:
                testStep.setDescription(row.getCell(startStepsCol).toString());
                break;
              case 3:
                testStep.setPageName(row.getCell(startStepsCol).toString());
                break;
              case 4:
                testStep.setPageObject(row.getCell(startStepsCol).toString());
                break;
              case 5:
                testStep.setActionKeyword(row.getCell(startStepsCol).toString());
                break;
              case 6:
                try {
                  testStep.setData(row.getCell(startStepsCol).toString()==null?"":row.getCell(startStepsCol).toString());
                }catch (NullPointerException np){
                  testStep.setData("");
                }

                param = "";
                if (!testStep.getData().equals("") && !dataDriven.equals("")){
                  ddtTotalCols = dataSheet.getRow(0).getLastCellNum();
                  for (int startDataCol = 0; startDataCol < ddtTotalCols; startDataCol++){
                    param = dataSheet.getRow(0).getCell(startDataCol).toString();
                    if (testStep.getData().equals(param)){
                      try{
                        param = dataSheet.getRow(ddt).getCell(startDataCol).toString();
                      }catch (Exception e){
                        param = "";
                      }
                      testStep.setData(param);
                      break;
                    }
                  }
                }
                break;
            }
          }
          if (!breaking){
            executeActions();
            if (testStepResult){
              excelReader.setCellData(KEYWORD_PASS, startStepsRow+1, totalCols-1, excelTSWorkSheet, excelWorkBook);

              if(!dataDriven.equals("")){
                excelReader.setCellData(KEYWORD_PASS, ddt, ddtTotalCols -1, dataDriven, excelWorkBook);
              }
            }else{
              testCaseResult = false;
              if (totalDataRows > 1){
                excelReader.setCellData(KEYWORD_FAIL, ddt, ddtTotalCols -1, dataDriven, excelWorkBook);
                tsResult = false;
                rowStepError = startStepsRow + 1;
              }
              excelReader.setCellData(KEYWORD_FAIL, startStepsRow+1, totalCols-1, excelTSWorkSheet, excelWorkBook);
              excelReader.setCellData(KEYWORD_FAIL, testCase.getTcKey(), TEST_CASE_RESULT_COL, excelTCWorkSheet, excelWorkBook);
              softAssert.assertTrue(testStepResult);
            }
          }
          breaking = false;
        }

        if (!tsResult){
          excelReader.setCellData(KEYWORD_FAIL, rowStepError, totalCols-1, excelTSWorkSheet, excelWorkBook);
        }

      }

      if (testCaseResult){
        excelReader.setCellData(KEYWORD_PASS, testCase.getTcKey(), TEST_CASE_RESULT_COL, excelTCWorkSheet, excelWorkBook);
      }
    }
  }
}
