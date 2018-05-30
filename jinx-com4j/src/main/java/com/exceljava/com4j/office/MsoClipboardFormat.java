package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoClipboardFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoClipboardFormatMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoClipboardFormatNative(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoClipboardFormatHTML(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoClipboardFormatRTF(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoClipboardFormatPlainText(4),
  ;

  private final int value;
  MsoClipboardFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
