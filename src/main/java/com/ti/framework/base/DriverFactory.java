package com.ti.framework.base;

import io.github.bonigarcia.wdm.WebDriverManager;
import io.github.bonigarcia.wdm.config.DriverManagerType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.safari.SafariDriver;

public class DriverFactory {

    private String getOS = System.getProperty("os.name").toLowerCase();
    private String driverProperty = "webdriver.chrome.driver";
    private String driverPath = System.getProperty("user.dir") + "/src/main/resources/";
    private String driverName = (getOS.contains("mac")) ? "chromedriver" : "chromedriver.exe";
    private String braveLocation = "C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe";

    private static DriverFactory instance = new DriverFactory();

    public static DriverFactory getInstance() {
        return instance;
    }

    // Thread local driver object for webdriver
    ThreadLocal<WebDriver> driver = new ThreadLocal<WebDriver>();

    /**
     * Method to get the driver object and launch the browser
     *
     * @return
     */
    public WebDriver getDriver() {
        return driver.get();
    }

    /**
     * Method to set the driver object and select the browser
     *
     * @param
     * @return
     */
    public void setDriver(BrowserType browser) {

        if (getOS.contains("mac") && browser.equals(BrowserType.SAFARI)) {
            driver.set(new SafariDriver());
        } else if (browser.equals(BrowserType.BRAVE)) {
            System.setProperty(driverProperty, driverPath + driverName);
            driver.set(new ChromeDriver(new ChromeOptions().setBinary(braveLocation)));
        } else {
            WebDriverManager.getInstance(DriverManagerType.valueOf(browser.toString())).setup();
            switch (browser) {
                case CHROME:
                    driver.set(new ChromeDriver());
                    break;
                case FIREFOX:
                    driver.set(new FirefoxDriver());
                    break;
                case EDGE:
                    driver.set(new EdgeDriver());
                    break;
                case IEXPLORER:
                    if (getOS.contains("win")) {
                        driver.set(new InternetExplorerDriver());
                    }
                    break;
            }
        }

        driver.get().manage().window().maximize();
    }

    /**
     * Method to quits the driver and closes the browser
     */
    public void removeDriver() {
        if (driver.get() != null) {
            try {
                driver.get().quit(); // First quit WebDriver session gracefully
                driver.remove(); // Remove WebDriver reference from the ThreadLocal variable.
            } catch (Exception e) {
                System.err.println("Unable to gracefully quit WebDriver."); // We'll replace this with actual Loggers later - don't worry !
            }
        }
    }
}
