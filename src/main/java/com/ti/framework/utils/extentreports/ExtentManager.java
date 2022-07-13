package com.ti.framework.utils.extentreports;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;

public class ExtentManager {
    public static final ExtentReports extentReports = new ExtentReports();

    public synchronized static ExtentReports createExtentReports(){
        ExtentSparkReporter reporter = new ExtentSparkReporter("./results/extent-report.html");
        reporter.config().setReportName("Test Execution Report");
        extentReports.attachReporter(reporter);
        extentReports.setSystemInfo("Plataforma", "Titanium Institute");
        extentReports.setSystemInfo("Author", "Gilberto Sánchez");

        return extentReports;
    }
}
