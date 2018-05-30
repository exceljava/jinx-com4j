package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPasteType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4104
   * </p>
   */
  xlPasteAll(-4104),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlPasteAllUsingSourceTheme(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlPasteAllMergingConditionalFormats(14),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlPasteAllExceptBorders(7),
  /**
   * <p>
   * The value of this constant is -4122
   * </p>
   */
  xlPasteFormats(-4122),
  /**
   * <p>
   * The value of this constant is -4123
   * </p>
   */
  xlPasteFormulas(-4123),
  /**
   * <p>
   * The value of this constant is -4144
   * </p>
   */
  xlPasteComments(-4144),
  /**
   * <p>
   * The value of this constant is -4163
   * </p>
   */
  xlPasteValues(-4163),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlPasteColumnWidths(8),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlPasteValidation(6),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlPasteFormulasAndNumberFormats(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlPasteValuesAndNumberFormats(12),
  ;

  private final int value;
  XlPasteType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
