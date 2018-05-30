package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBalloonButtonType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -15
   * </p>
   */
  msoBalloonButtonYesToAll(-15),
  /**
   * <p>
   * The value of this constant is -14
   * </p>
   */
  msoBalloonButtonOptions(-14),
  /**
   * <p>
   * The value of this constant is -13
   * </p>
   */
  msoBalloonButtonTips(-13),
  /**
   * <p>
   * The value of this constant is -12
   * </p>
   */
  msoBalloonButtonClose(-12),
  /**
   * <p>
   * The value of this constant is -11
   * </p>
   */
  msoBalloonButtonSnooze(-11),
  /**
   * <p>
   * The value of this constant is -10
   * </p>
   */
  msoBalloonButtonSearch(-10),
  /**
   * <p>
   * The value of this constant is -9
   * </p>
   */
  msoBalloonButtonIgnore(-9),
  /**
   * <p>
   * The value of this constant is -8
   * </p>
   */
  msoBalloonButtonAbort(-8),
  /**
   * <p>
   * The value of this constant is -7
   * </p>
   */
  msoBalloonButtonRetry(-7),
  /**
   * <p>
   * The value of this constant is -6
   * </p>
   */
  msoBalloonButtonNext(-6),
  /**
   * <p>
   * The value of this constant is -5
   * </p>
   */
  msoBalloonButtonBack(-5),
  /**
   * <p>
   * The value of this constant is -4
   * </p>
   */
  msoBalloonButtonNo(-4),
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  msoBalloonButtonYes(-3),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBalloonButtonCancel(-2),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoBalloonButtonOK(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoBalloonButtonNull(0),
  ;

  private final int value;
  MsoBalloonButtonType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
