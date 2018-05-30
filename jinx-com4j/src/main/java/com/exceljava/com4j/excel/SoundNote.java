package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface SoundNote extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   */

  @DISPID(117)
  java.lang.Object delete();


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(917)
  java.lang.Object _import(
    java.lang.String filename);


  /**
   */

  @DISPID(918)
  java.lang.Object play();


  /**
   */

  @DISPID(919)
  java.lang.Object record();


  // Properties:
}
