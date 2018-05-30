package com.exceljava.com4j.excel.events;

import com4j.*;

@IID("{0002441B-0000-0000-C000-000000000046}")
public abstract class RefreshEvents {
  // Methods:
  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1596)
  public void beforeRefresh(
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param success Mandatory boolean parameter.
   */

  @DISPID(1597)
  public void afterRefresh(
    boolean success) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
