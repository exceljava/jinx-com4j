package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002441B-0001-0000-C000-000000000046}")
public interface IRefreshEvents extends Com4jObject {
  // Methods:
  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(7)
  void beforeRefresh(
    Holder<Boolean> cancel);


  /**
   * @param success Mandatory boolean parameter.
   */

  @VTID(8)
  void afterRefresh(
    boolean success);


  // Properties:
}
