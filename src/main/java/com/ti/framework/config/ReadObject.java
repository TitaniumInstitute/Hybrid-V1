package com.ti.framework.config;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class ReadObject implements PropertyManagerInterface{
  private static ReadObject instance;
  private static final Object lock = new Object();
  private static String propertyFilePath = System.getProperty("user.dir") + "/src/main/resources/ObjRepo.txt";
  private Properties properties = new Properties();

  public static ReadObject getInstance(){
    if(instance == null){
      synchronized (lock){
        instance = new ReadObject();
        instance.loadData();
      }
    }
    return instance;
  }

  @Override
  public void loadData(){
    try {
      properties.load(new FileInputStream(propertyFilePath));
    }catch (IOException e){
      System.err.println("Class: PropertyManager | Method: loadData | Exception:" + e.getMessage());
    }
  }

  @Override
  public String getProperty(String aproperty){
    return properties.getProperty(aproperty);
  }
}
