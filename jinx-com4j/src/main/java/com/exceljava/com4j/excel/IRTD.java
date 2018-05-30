package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002446E-0001-0000-C000-000000000046}")
public interface IRTD extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "ThrottleInterval"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(7)
  int getThrottleInterval();


  /**
   * <p>
   * Setter method for the COM property "ThrottleInterval"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(8)
  void setThrottleInterval(
    int rhs);


  /**
   */

  @VTID(9)
  void refreshData();


  /**
   */

  @VTID(10)
  void restartServers();


  // Properties:
}
