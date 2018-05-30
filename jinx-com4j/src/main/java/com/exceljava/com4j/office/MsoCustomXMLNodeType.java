package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCustomXMLNodeType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCustomXMLNodeElement(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCustomXMLNodeAttribute(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCustomXMLNodeText(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCustomXMLNodeCData(4),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoCustomXMLNodeProcessingInstruction(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoCustomXMLNodeComment(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoCustomXMLNodeDocument(9),
  ;

  private final int value;
  MsoCustomXMLNodeType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
