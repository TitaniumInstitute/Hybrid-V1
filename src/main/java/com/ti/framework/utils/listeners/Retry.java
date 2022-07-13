package com.ti.framework.utils.listeners;

import static com.ti.framework.utils.extentreports.ExtentTestManager.getTest;

import com.aventstack.extentreports.Status;
import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;

public class Retry implements IRetryAnalyzer {
    private byte count = 0;
    private static byte maxTry = 1;

    @Override
    public boolean retry(ITestResult iTestResult) {
        if(!iTestResult.isSuccess()){
            if(count < maxTry){
                count++;
                iTestResult.setStatus(ITestResult.FAILURE);
                extendReportsFailOperation(iTestResult);
            }
        }
        return false;
    }

    private void extendReportsFailOperation(ITestResult iTestResult) {
        getTest().log(Status.FAIL, "Test Failed");
        //add screenshot
    }


}
