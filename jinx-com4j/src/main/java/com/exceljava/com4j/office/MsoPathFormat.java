package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPathFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPathTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoPathTypeNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoPathType1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoPathType2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoPathType3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoPathType4(4),
  ;

  private final int value;
  MsoPathFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
