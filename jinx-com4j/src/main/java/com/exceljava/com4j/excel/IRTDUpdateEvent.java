package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{A43788C1-D91B-11D3-8F39-00C04F3651B8}")
public interface IRTDUpdateEvent extends Com4jObject {
  // Methods:
  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(7)
  void updateNotify();


  /**
   * <p>
   * Getter method for the COM property "HeartbeatInterval"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(8)
  int getHeartbeatInterval();


  /**
   * <p>
   * Setter method for the COM property "HeartbeatInterval"
   * </p>
   * @param plRetVal Mandatory int parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(9)
  void setHeartbeatInterval(
    int plRetVal);


  /**
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(10)
  void disconnect();


  // Properties:
}
