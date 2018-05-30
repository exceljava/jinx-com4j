package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0301-0000-0000-C000-000000000046}")
public interface _IMsoOleAccDispObj extends com.exceljava.com4j.office.IAccessible {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(28)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(29)
  int getCreator();


  // Properties:
}
