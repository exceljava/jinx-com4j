package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoButtonStyleHidden implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoButtonWrapText(4),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoButtonTextBelow(8),
  ;

  private final int value;
  MsoButtonStyleHidden(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
