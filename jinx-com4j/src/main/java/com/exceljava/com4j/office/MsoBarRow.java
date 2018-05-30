package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBarRow implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoBarRowFirst(0),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoBarRowLast(-1),
  ;

  private final int value;
  MsoBarRow(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
