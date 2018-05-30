package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface RTD extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "ThrottleInterval"
   * </p>
   */

  @DISPID(2240)
  @PropGet
  int getThrottleInterval();


  /**
   * <p>
   * Setter method for the COM property "ThrottleInterval"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(2240)
  @PropPut
  void setThrottleInterval(
    int rhs);


  /**
   */

  @DISPID(2241)
  void refreshData();


  /**
   */

  @DISPID(2242)
  void restartServers();


  // Properties:
}
