package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlCategoryLabelLevel implements ComEnum {
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  xlCategoryLabelLevelNone(-3),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  xlCategoryLabelLevelCustom(-2),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlCategoryLabelLevelAll(-1),
  ;

  private final int value;
  XlCategoryLabelLevel(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
