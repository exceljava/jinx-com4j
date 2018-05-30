package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoHTMLProjectState implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoHTMLProjectStateDocumentLocked(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoHTMLProjectStateProjectLocked(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoHTMLProjectStateDocumentProjectUnlocked(3),
  ;

  private final int value;
  MsoHTMLProjectState(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
