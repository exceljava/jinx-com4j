package com.exceljava.com4j.vbide  ;

import com4j.*;

/**
 */
public enum vbext_ComponentType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  vbext_ct_StdModule(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  vbext_ct_ClassModule(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  vbext_ct_MSForm(3),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  vbext_ct_ActiveXDesigner(11),
  /**
   * <p>
   * The value of this constant is 100
   * </p>
   */
  vbext_ct_Document(100),
  ;

  private final int value;
  vbext_ComponentType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
