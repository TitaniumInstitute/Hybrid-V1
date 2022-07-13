package com.ti.framework.utils.verifications;

import org.openqa.selenium.WebElement;

public interface VerificationInterface {

  void verifyTestIsDisplayed(WebElement element, String text);

  void verifyTextAreEquals(String actualText, String expectedText);

  void verifyElementExist(WebElement element);

}
