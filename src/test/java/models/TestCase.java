package models;

import lombok.Data;

@Data
public class TestCase extends Test{
  private String runMode;
  private int tcKey;
}
