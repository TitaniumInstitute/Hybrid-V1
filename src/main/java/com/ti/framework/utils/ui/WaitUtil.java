package com.ti.framework.utils.ui;

import com.ti.framework.base.DriverFactory;
import com.ti.framework.base.FrameworkException;
import java.time.Duration;
import java.util.List;
import org.openqa.selenium.JavascriptException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class WaitUtil {

  private static WebDriver driver = DriverFactory.getInstance().getDriver();

  public static void sync() throws FrameworkException {
    String pageLoadStatus;
    try {
      JavascriptExecutor js = (JavascriptExecutor) driver;
      do {
        pageLoadStatus = (String) js.executeScript("return document.readyState");
      }while (!pageLoadStatus.equals("complete"));
    }catch(JavascriptException jse){
      throw new FrameworkException(jse.getMessage());
    }
  }

  private static WebDriverWait explicitWait(){
    return new WebDriverWait(driver, Duration.ofSeconds(10));
  }

  public static void elementVisible(WebElement element){
    explicitWait().until(ExpectedConditions.visibilityOf(element));
  }

  public static void allElementsVisible(List<WebElement> elements){
    explicitWait().until(ExpectedConditions.visibilityOfAllElements(elements));
  }

  public static void elementClickable(WebElement element){
    explicitWait().until(ExpectedConditions.elementToBeClickable(element));
    element.click();
  }

  public static void elementNotVisible(WebElement element){
    explicitWait().until(ExpectedConditions.invisibilityOf(element));
  }

  public static void wairForAlert(){
    explicitWait().until(ExpectedConditions.alertIsPresent());
  }
}
