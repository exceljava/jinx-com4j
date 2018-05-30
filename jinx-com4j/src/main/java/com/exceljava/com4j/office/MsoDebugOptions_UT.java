package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C038B-0000-0000-C000-000000000046}")
public interface MsoDebugOptions_UT extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "CollectionName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getCollectionName();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.MsoDebugOptions_UTRunResult
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.MsoDebugOptions_UTRunResult run();


  // Properties:
}
