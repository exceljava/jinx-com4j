package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03A7-0000-0000-C000-000000000046}")
public interface IRibbonUI extends Com4jObject {
  // Methods:
  /**
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void invalidate();


  /**
   * @param controlID Mandatory java.lang.String parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  void invalidateControl(
    java.lang.String controlID);


  /**
   * @param controlID Mandatory java.lang.String parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  void invalidateControlMso(
    java.lang.String controlID);


  /**
   * @param controlID Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  void activateTab(
    java.lang.String controlID);


  /**
   * @param controlID Mandatory java.lang.String parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  void activateTabMso(
    java.lang.String controlID);


  /**
   * @param controlID Mandatory java.lang.String parameter.
   * @param namespace Mandatory java.lang.String parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(12)
  void activateTabQ(
    java.lang.String controlID,
    java.lang.String namespace);


  // Properties:
}
