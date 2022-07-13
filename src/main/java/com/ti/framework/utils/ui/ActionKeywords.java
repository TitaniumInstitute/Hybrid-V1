package com.ti.framework.utils.ui;

import static com.ti.framework.utils.ui.WaitUtil.elementNotVisible;
import static com.ti.framework.utils.ui.WaitUtil.sync;
import static com.ti.framework.utils.logs.Log.*;

import com.ti.framework.base.FrameworkException;
import org.openqa.selenium.support.ui.Select;

public class ActionKeywords extends SeleniumUtil{
  public boolean preLoading(String object, String data){
    try{
      sync();
      elementNotVisible(highLight(findElement(object)));
      return true;
    }catch (Exception e){
      refresh();
      preLoading(object,data);
      return false;
    }
  }

  public boolean type(String object, String data){
    try {
      info("Entering the text: " + data + "in " + object);
      element = highLight(findElement(object));
      element.clear();
      element.sendKeys(data);
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean click(String object, String data){
    try {
      info("Clicking on Object " + object);
      highLight(findElement(object)).click();
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean getElementText(String object, String data){
    try {
      info("Getting Text: " + data + " from " + object);
      String text = highLight(findElement(object)).getText();
      if (data.equals("")){
        info("Getting the text: " + text);
      }else if (!data.equals("")&&data.equals(text)){
        info("The text: " + text + " is equals to: " + data);
      }else if (!data.equals("")&&!data.equals(text)){
        info("The text: " + text + " is not equals to: " + data);
        new FrameworkException("The text: " + text + " is not equals to: " + data);
        return false;
      }
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean getTextBoxText(String object, String data){
    try {
      info("Getting the text from: " + data + " from " + object);
      String text = highLight(findElement(object)).getAttribute("value");
      if (data.equals("")){
        info("Getting the text: " + text);
      }else if (!data.equals("")&&data.equals(text)){
        info("The text: " + text + " is equals to: " + data);
      }else if (!data.equals("")&&!data.equals(text)){
        info("The text: " + text + " is not equals to: " + data);
        new FrameworkException("The text: " + text + " is not equals to: " + data);
        return false;
      }
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean submitForm(String object, String data){
    try {
      info("Form Submitted");
      highLight(findElement(object)).submit();
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean selectByTextFromList(String object, String data){
    try {
      if (!data.equals("")){
        info("Selecting " + data + " from dropdown " + object);
        new Select(highLight(findElement(object))).selectByVisibleText(data);
      }
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean selectByValueFromList(String object, String data){
    try {
      if (!data.equals("")){
        info("Selecting " + data + " from dropdown " + object);
        new Select(highLight(findElement(object))).selectByValue(data);
      }
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

  public boolean selectByIndexFromList(String object, String data){
    try {
      if (!data.equals("")){
        info("Selecting " + data + " from dropdown " + object);
        new Select(highLight(findElement(object))).selectByIndex(Integer.parseInt(data));
      }
      return true;
    }catch (Exception e){
      new FrameworkException("Not able to type text ", e);
      return false;
    }
  }

}
