package test;

import static com.ti.framework.config.Constants.BASE_URL;
import static com.ti.framework.config.Constants.VIDEO_FOLDER;

import com.ti.framework.base.BasePage;
import com.ti.framework.base.BrowserType;
import com.ti.framework.base.DriverFactory;
import com.ti.framework.config.PropertyManager;
import com.ti.framework.utils.data.ExcelUtils;
import com.ti.framework.utils.logs.Log;
import com.ti.framework.utils.ui.ActionKeywords;
import com.ti.framework.utils.video.SpecializedScreenRecorder;
import java.lang.reflect.Method;
import models.TestCase;
import models.TestStep;
import org.apache.poi.ss.usermodel.Row;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;
import org.testng.asserts.SoftAssert;

public class Base extends BasePage {
    protected static SpecializedScreenRecorder videoRec;
    protected static Method[] methods;
    protected static ActionKeywords actionKeywords;
    protected static TestStep testStep;
    protected static TestCase testCase;
    protected static String excelWorkBook = PropertyManager.getInstance().getProperty("DataEngine");
    protected static String excelTSWorkSheet = PropertyManager.getInstance().getProperty("TestStepSheet");
    protected static String excelTCWorkSheet = PropertyManager.getInstance().getProperty("TestCaseSheet");
    protected static int startStepsRow;
    protected static int startStepsCol;
    protected static boolean breaking = false;
    protected static boolean testCaseResult;
    protected static boolean testStepResult = true;
    protected static ExcelUtils excelReader;
    protected static Row row;
    protected static int totalDataRows;
    protected static String param= "";
    protected static SoftAssert softAssert;

    public Base(){
        testStep = new TestStep();
        testCase = new TestCase();
        softAssert = new SoftAssert();
    }

    @BeforeTest
    @Parameters("browser")
    void setup(String browser) {
        DriverFactory.getInstance().setDriver(BrowserType.valueOf(browser));
        if(Boolean.parseBoolean(PropertyManager.getInstance().getProperty("Video"))){
            videoRec.startRecording(VIDEO_FOLDER);
        }
        DriverFactory.getInstance().getDriver().navigate().to(BASE_URL);
        Log.info("Tests is starting!");
        actionKeywords = new ActionKeywords();
        methods = actionKeywords.getClass().getMethods();
    }

    @AfterTest
    void turnDown() throws Exception {
        Log.info("Tests are ending!");
        DriverFactory.getInstance().removeDriver();
        if(videoRec != null){
            videoRec.stopRecording();
        }
        softAssert.assertAll();
    }
}
