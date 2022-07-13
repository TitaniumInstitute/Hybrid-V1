package com.ti.framework.base;

import com.ti.framework.utils.verifications.VerificationUtil;
import org.openqa.selenium.support.PageFactory;

public class BasePage {
    public static BasePage actualPage;
    public VerificationUtil verification = new VerificationUtil();

    public <TPage extends BasePage> TPage getInstance(Class<TPage> page){
        Object obj = PageFactory.initElements(DriverFactory.getInstance().getDriver(), page);
        return page.cast(obj);
    }

    public <TPage extends BasePage> TPage as(Class<TPage> pageInstance){
        try{
            return (TPage)this;
        }catch (Exception e){
            e.getStackTrace();
        }
        return null;
    }
}
