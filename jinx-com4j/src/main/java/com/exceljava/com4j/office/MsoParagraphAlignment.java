package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoParagraphAlignment implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoAlignMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAlignLeft(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAlignCenter(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoAlignRight(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoAlignJustify(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoAlignDistribute(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoAlignThaiDistribute(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoAlignJustifyLow(7),
  ;

  private final int value;
  MsoParagraphAlignment(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
