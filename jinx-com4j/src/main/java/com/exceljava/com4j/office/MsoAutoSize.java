package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoAutoSize implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoAutoSizeMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoAutoSizeNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAutoSizeShapeToFitText(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAutoSizeTextToFitShape(2),
  ;

  private final int value;
  MsoAutoSize(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
