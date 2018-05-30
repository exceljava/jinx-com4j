package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCommentDisplayMode implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlNoIndicator(0),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlCommentIndicatorOnly(-1),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCommentAndIndicator(1),
  ;

  private final int value;
  XlCommentDisplayMode(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
