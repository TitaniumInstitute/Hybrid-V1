package com.ti.framework.utils.ui;

import static com.ti.framework.config.Constants.SCREENSHOT_FOLDER;
import static com.ti.framework.config.CreateFolder.createFolder;

import com.ti.framework.base.DriverFactory;
import com.ti.framework.base.FrameworkException;
import com.ti.framework.utils.logs.Log;
import java.io.File;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Calendar;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class SeleniumUtil {

  private static WebDriver driver = DriverFactory.getInstance().getDriver();
  protected static WebElement element;

  protected static WebElement findElement(String locator){
    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
    try{
      return driver.findElement(By.id(locator));
    }catch (NoSuchElementException ne){
      return returnElementName(locator);
    }
  }

  private static WebElement returnElementName(String name){
    try{
      return driver.findElement(By.name(name));
    }catch (NoSuchElementException ne){
      return returnElementXpath(name);
    }
  }

  private static WebElement returnElementXpath(String xpath) {
    try{
      return driver.findElement(By.xpath(xpath));
    }catch (NoSuchElementException ne){
      return returnElementCssSelector(xpath);
    }
  }

  private static WebElement returnElementCssSelector(String cssSelector) {
    try{
      return driver.findElement(By.cssSelector(cssSelector));
    }catch (NoSuchElementException ne){
      return returnElementClassName(cssSelector);
    }
  }

  private static WebElement returnElementClassName(String className) {
    try{
      return driver.findElement(By.className(className));
    }catch (NoSuchElementException ne){
      return returnElementLinkText(className);
    }
  }

  private static WebElement returnElementLinkText(String linkText) {
    try{
      return driver.findElement(By.linkText(linkText));
    }catch (NoSuchElementException ne){
      return returnElementTagName(linkText);
    }
  }

  private static WebElement returnElementTagName(String tag) {
    try{
      return driver.findElement(By.tagName(tag));
    }catch (NoSuchElementException ne){
      new FrameworkException("Class SeleniumUtil | Method findElement | Exception desc: Can not find the element: " + tag);
      return null;
    }
  }

  public static WebElement highLight(WebElement element) {
    for (int i = 0; i < 3; i++) {
      //Creating JavaScriptExecuter Interface
      JavascriptExecutor js = (JavascriptExecutor) driver;
      for (int iCnt = 0; iCnt < 3; iCnt++) {
        js.executeScript("arguments[0].setAttribute('style','background: yellow')",
            element);
        try {
          Thread.sleep(20);
        } catch (InterruptedException e) {
          Thread.currentThread().interrupt();
        }
        js.executeScript("arguments[0].setAttribute('style','background:')", element);
      }
    }
    return element;
  }

  public static String getScreenShot(WebElement element) {
    String ssDateTime = new SimpleDateFormat("yyMMddHHmmss")
        .format(Calendar.getInstance().getTime());
    String file = null;
    try {
      file = createFolder(SCREENSHOT_FOLDER) + "/" + ssDateTime + ".png";
    } catch (FrameworkException e) {
      e.printStackTrace();
    }

    try {
      File srcFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
      FileUtils.copyFile(
          (element == null ? srcFile : element.getScreenshotAs(OutputType.FILE)),
          new File(file));
    } catch (Exception e) {
      Log.error(
          "Class SeleniumUtils | Method takeSnapShot | Exception desc: " + e.getMessage());
    }
    return file;
  }

  public static void scrollToElement(WebElement element) throws FrameworkException {
    JavascriptExecutor js = (JavascriptExecutor) driver;
    try{
      js.executeScript("arguments[0].scrollIntoView(true)", element);
    }catch (JavascriptException jse){
      throw new FrameworkException(jse.getMessage());
    }
  }

  public static void refresh(){
    driver.navigate().refresh();
  }
}
