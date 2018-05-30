package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0914-0000-0000-C000-000000000046}")
public interface WebPageFonts extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory com.exceljava.com4j.office.MsoCharacterSet parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.WebPageFont
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.office.WebPageFont getItem(
    com.exceljava.com4j.office.MsoCharacterSet index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
