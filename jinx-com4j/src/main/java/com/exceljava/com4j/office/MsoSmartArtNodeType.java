package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoSmartArtNodeType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSmartArtNodeTypeDefault(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoSmartArtNodeTypeAssistant(2),
  ;

  private final int value;
  MsoSmartArtNodeType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
