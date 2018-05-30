package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlThemeColor implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlThemeColorDark1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlThemeColorLight1(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlThemeColorDark2(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlThemeColorLight2(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlThemeColorAccent1(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlThemeColorAccent2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlThemeColorAccent3(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlThemeColorAccent4(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlThemeColorAccent5(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlThemeColorAccent6(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlThemeColorHyperlink(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlThemeColorFollowedHyperlink(12),
  ;

  private final int value;
  XlThemeColor(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
