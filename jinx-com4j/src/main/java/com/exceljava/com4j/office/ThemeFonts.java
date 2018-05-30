package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03A4-0000-0000-C000-000000000046}")
public interface ThemeFonts extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory com.exceljava.com4j.office.MsoFontLanguageIndex parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.ThemeFont
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.ThemeFont item(
    com.exceljava.com4j.office.MsoFontLanguageIndex index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
