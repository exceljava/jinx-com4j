package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlHAlign implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4108
   * </p>
   */
  xlHAlignCenter(-4108),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlHAlignCenterAcrossSelection(7),
  /**
   * <p>
   * The value of this constant is -4117
   * </p>
   */
  xlHAlignDistributed(-4117),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlHAlignFill(5),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlHAlignGeneral(1),
  /**
   * <p>
   * The value of this constant is -4130
   * </p>
   */
  xlHAlignJustify(-4130),
  /**
   * <p>
   * The value of this constant is -4131
   * </p>
   */
  xlHAlignLeft(-4131),
  /**
   * <p>
   * The value of this constant is -4152
   * </p>
   */
  xlHAlignRight(-4152),
  ;

  private final int value;
  XlHAlign(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
