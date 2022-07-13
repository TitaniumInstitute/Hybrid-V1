package com.ti.framework.utils.listeners;

import static com.ti.framework.utils.extentreports.ExtentTestManager.getTest;
import static com.ti.framework.utils.logs.Log.info;

import com.aventstack.extentreports.Status;
import com.ti.framework.base.DriverFactory;
import com.ti.framework.utils.extentreports.ExtentManager;
import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

public class TestListener extends DriverFactory implements ITestListener {
    private static String getTestMethodName(ITestResult iTestResult){
        return iTestResult.getMethod().getConstructorOrMethod().getName();
    }

    @Override
    public void onStart(ITestContext iTestContext){
        info("Starting method " + iTestContext.getName());
        iTestContext.setAttribute("WebDriver", getInstance().getDriver());
    }

    @Override
    public void onFinish(ITestContext iTestContext){
        info("Finishing method " + iTestContext.getName());
        info("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX          -E--N--D           XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
        ExtentManager.extentReports.flush();
    }

    @Override
    public void onTestFailure(ITestResult iTestResult){
        info(getTestMethodName(iTestResult) + " test is failed.");
        getTest().log(Status.FAIL, "Test Failed");
        //generate screenshot
    }

    @Override
    public void onTestSkipped(ITestResult iTestResult){
        info(getTestMethodName(iTestResult) + " test is skipped.");
        getTest().log(Status.SKIP, "Test Skipped");
    }

    @Override
    public void onTestSuccess(ITestResult iTestResult){
        info(getTestMethodName(iTestResult) + " test is succed.");
        getTest().log(Status.PASS, "Test Passed");
    }

    @Override
    public void onTestStart(ITestResult iTestResult){
        info("************************************************************************");
        info(getTestMethodName(iTestResult) + "Test is starting");
        info("************************************************************************");
    }

    @Override
    public void onTestFailedButWithinSuccessPercentage(ITestResult iTestResult){
        info("Test failed but it is in defined success ratio " + getTestMethodName(iTestResult));
    }
}
