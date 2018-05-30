package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBackgroundStyleIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBackgroundStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoBackgroundStyleNotAPreset(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBackgroundStylePreset1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBackgroundStylePreset2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBackgroundStylePreset3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoBackgroundStylePreset4(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoBackgroundStylePreset5(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoBackgroundStylePreset6(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoBackgroundStylePreset7(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoBackgroundStylePreset8(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoBackgroundStylePreset9(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoBackgroundStylePreset10(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoBackgroundStylePreset11(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoBackgroundStylePreset12(12),
  ;

  private final int value;
  MsoBackgroundStyleIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
