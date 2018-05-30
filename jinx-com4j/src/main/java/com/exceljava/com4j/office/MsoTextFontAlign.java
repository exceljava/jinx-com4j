package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextFontAlign implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoFontAlignMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoFontAlignAuto(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFontAlignTop(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFontAlignCenter(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFontAlignBaseline(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFontAlignBottom(4),
  ;

  private final int value;
  MsoTextFontAlign(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
