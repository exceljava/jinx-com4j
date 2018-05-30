package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileFindListBy implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoListbyName(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoListbyTitle(2),
  ;

  private final int value;
  MsoFileFindListBy(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
