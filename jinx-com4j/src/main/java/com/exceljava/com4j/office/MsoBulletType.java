package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBulletType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBulletMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoBulletNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBulletUnnumbered(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBulletNumbered(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBulletPicture(3),
  ;

  private final int value;
  MsoBulletType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
