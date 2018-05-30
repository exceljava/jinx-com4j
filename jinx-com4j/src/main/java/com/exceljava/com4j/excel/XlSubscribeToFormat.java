package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSubscribeToFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4147
   * </p>
   */
  xlSubscribeToPicture(-4147),
  /**
   * <p>
   * The value of this constant is -4158
   * </p>
   */
  xlSubscribeToText(-4158),
  ;

  private final int value;
  XlSubscribeToFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
