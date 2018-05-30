package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoButtonStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoButtonAutomatic(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoButtonIcon(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoButtonCaption(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoButtonIconAndCaption(3),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoButtonIconAndWrapCaption(7),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoButtonIconAndCaptionBelow(11),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoButtonWrapCaption(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoButtonIconAndWrapCaptionBelow(15),
  ;

  private final int value;
  MsoButtonStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
