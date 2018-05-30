package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoHorizontalAnchor implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoHorizontalAnchorMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAnchorNone(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAnchorCenter(2),
  ;

  private final int value;
  MsoHorizontalAnchor(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
