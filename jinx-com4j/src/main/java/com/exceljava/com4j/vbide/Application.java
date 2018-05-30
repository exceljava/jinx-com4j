package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E158-0000-0000-C000-000000000046}")
public interface Application extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Version"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(7)
  java.lang.String getVersion();


  // Properties:
}
