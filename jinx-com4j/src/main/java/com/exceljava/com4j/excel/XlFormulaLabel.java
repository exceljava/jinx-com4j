package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFormulaLabel implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlNoLabels(-4142),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRowLabels(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlColumnLabels(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMixedLabels(3),
  ;

  private final int value;
  XlFormulaLabel(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
