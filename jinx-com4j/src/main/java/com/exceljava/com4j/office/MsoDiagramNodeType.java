package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoDiagramNodeType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoDiagramNode(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoDiagramAssistant(2),
  ;

  private final int value;
  MsoDiagramNodeType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
