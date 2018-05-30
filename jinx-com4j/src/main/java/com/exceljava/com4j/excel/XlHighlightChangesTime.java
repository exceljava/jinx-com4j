package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlHighlightChangesTime implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSinceMyLastSave(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAllChanges(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlNotYetReviewed(3),
  ;

  private final int value;
  XlHighlightChangesTime(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
