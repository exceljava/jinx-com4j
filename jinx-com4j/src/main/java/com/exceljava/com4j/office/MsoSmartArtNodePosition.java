package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSmartArtNodePosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSmartArtNodeDefault(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSmartArtNodeAfter(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoSmartArtNodeBefore(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoSmartArtNodeAbove(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoSmartArtNodeBelow(5),
  ;

  private final int value;
  MsoSmartArtNodePosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
