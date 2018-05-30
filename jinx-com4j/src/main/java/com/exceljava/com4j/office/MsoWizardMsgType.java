package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoWizardMsgType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoWizardMsgLocalStateOn(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoWizardMsgLocalStateOff(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoWizardMsgShowHelp(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoWizardMsgSuspending(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoWizardMsgResuming(5),
  ;

  private final int value;
  MsoWizardMsgType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
