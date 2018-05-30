package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoAutomationSecurity implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAutomationSecurityLow(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAutomationSecurityByUI(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoAutomationSecurityForceDisable(3),
  ;

  private final int value;
  MsoAutomationSecurity(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
