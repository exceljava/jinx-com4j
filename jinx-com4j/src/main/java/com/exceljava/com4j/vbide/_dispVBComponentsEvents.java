package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface _dispVBComponentsEvents extends Com4jObject {
  // Methods:
  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   */

  @DISPID(1)
  void itemAdded(
    com.exceljava.com4j.vbide._VBComponent vbComponent);


  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   */

  @DISPID(2)
  void itemRemoved(
    com.exceljava.com4j.vbide._VBComponent vbComponent);


  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   * @param oldName Mandatory java.lang.String parameter.
   */

  @DISPID(3)
  void itemRenamed(
    com.exceljava.com4j.vbide._VBComponent vbComponent,
    java.lang.String oldName);


  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   */

  @DISPID(4)
  void itemSelected(
    com.exceljava.com4j.vbide._VBComponent vbComponent);


  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   */

  @DISPID(5)
  void itemActivated(
    com.exceljava.com4j.vbide._VBComponent vbComponent);


  /**
   * @param vbComponent Mandatory com.exceljava.com4j.vbide._VBComponent parameter.
   */

  @DISPID(6)
  void itemReloaded(
    com.exceljava.com4j.vbide._VBComponent vbComponent);


  // Properties:
}
