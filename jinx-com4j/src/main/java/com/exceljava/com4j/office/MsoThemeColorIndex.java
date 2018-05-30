package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoThemeColorIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoThemeColorMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoNotThemeColor(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoThemeColorDark1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoThemeColorLight1(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoThemeColorDark2(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoThemeColorLight2(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoThemeColorAccent1(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoThemeColorAccent2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoThemeColorAccent3(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoThemeColorAccent4(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoThemeColorAccent5(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoThemeColorAccent6(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoThemeColorHyperlink(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoThemeColorFollowedHyperlink(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoThemeColorText1(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoThemeColorBackground1(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoThemeColorText2(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoThemeColorBackground2(16),
  ;

  private final int value;
  MsoThemeColorIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
