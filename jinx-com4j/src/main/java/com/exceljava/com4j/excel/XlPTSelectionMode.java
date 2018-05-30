package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPTSelectionMode implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlLabelOnly(1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlDataAndLabel(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDataOnly(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlOrigin(3),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlButton(15),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlBlanks(4),
  /**
   * <p>
   * The value of this constant is 256
   * </p>
   */
  xlFirstRow(256),
  ;

  private final int value;
  XlPTSelectionMode(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
