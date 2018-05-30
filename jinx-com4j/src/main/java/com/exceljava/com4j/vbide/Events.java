package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E167-0000-0000-C000-000000000046}")
public interface Events extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "ReferencesEvents"
   * </p>
   * @param vbProject Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._ReferencesEvents
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.vbide._ReferencesEvents getReferencesEvents(
    com.exceljava.com4j.vbide._VBProject vbProject);


  /**
   * <p>
   * Getter method for the COM property "CommandBarEvents"
   * </p>
   * @param commandBarControl Mandatory com4j.Com4jObject parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._CommandBarControlEvents
   */

  @DISPID(205) //= 0xcd. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide._CommandBarControlEvents getCommandBarEvents(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject commandBarControl);


  // Properties:
}
