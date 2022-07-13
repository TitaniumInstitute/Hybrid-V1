package models;

import lombok.Data;

@Data
public class TestStep extends Test{
  private String pageName;
  private String pageObject;
  private String ActionKeyword;
  private String data;
}
