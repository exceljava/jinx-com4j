package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBlackWhiteMode implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBlackWhiteMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBlackWhiteAutomatic(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBlackWhiteGrayScale(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBlackWhiteLightGrayScale(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoBlackWhiteInverseGrayScale(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoBlackWhiteGrayOutline(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoBlackWhiteBlackTextAndLine(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoBlackWhiteHighContrast(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoBlackWhiteBlack(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoBlackWhiteWhite(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoBlackWhiteDontShow(10),
  ;

  private final int value;
  MsoBlackWhiteMode(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
