package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlColumnDataType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlGeneralFormat(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTextFormat(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMDYFormat(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDMYFormat(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlYMDFormat(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlMYDFormat(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlDYMFormat(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlYDMFormat(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlSkipColumn(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlEMDFormat(10),
  ;

  private final int value;
  XlColumnDataType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
