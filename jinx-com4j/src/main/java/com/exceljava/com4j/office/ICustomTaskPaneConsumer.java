package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C033E-0000-0000-C000-000000000046}")
public interface ICustomTaskPaneConsumer extends Com4jObject {
  // Methods:
  /**
   * @param ctpFactoryInst Mandatory com.exceljava.com4j.office.ICTPFactory parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void ctpFactoryAvailable(
    com.exceljava.com4j.office.ICTPFactory ctpFactoryInst);


  // Properties:
}
