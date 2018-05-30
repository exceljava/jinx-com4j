package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoThemeColorSchemeIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoThemeDark1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoThemeLight1(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoThemeDark2(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoThemeLight2(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoThemeAccent1(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoThemeAccent2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoThemeAccent3(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoThemeAccent4(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoThemeAccent5(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoThemeAccent6(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoThemeHyperlink(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoThemeFollowedHyperlink(12),
  ;

  private final int value;
  MsoThemeColorSchemeIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
