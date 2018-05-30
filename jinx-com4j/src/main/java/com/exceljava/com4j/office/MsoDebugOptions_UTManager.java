package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0389-0000-0000-C000-000000000046}")
public interface MsoDebugOptions_UTManager extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "UnitTests"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoDebugOptions_UTs
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.MsoDebugOptions_UTs getUnitTests();


  @VTID(9)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.MsoDebugOptions_UTs.class})
  com.exceljava.com4j.office.MsoDebugOptions_UT getUnitTests(
    int index);

  /**
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  void notifyStartOfTestSuiteRun();


  /**
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void notifyEndOfTestSuiteRun();


  /**
   * <p>
   * Getter method for the COM property "ReportErrors"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  boolean getReportErrors();


  /**
   * <p>
   * Setter method for the COM property "ReportErrors"
   * </p>
   * @param pfReportErrors Mandatory boolean parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  void setReportErrors(
    boolean pfReportErrors);


  // Properties:
}
