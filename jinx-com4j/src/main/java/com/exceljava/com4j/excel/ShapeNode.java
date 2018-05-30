package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000C0318-0000-0000-C000-000000000046}")
public interface ShapeNode extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "EditingType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoEditingType
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.MsoEditingType getEditingType();


  /**
   * <p>
   * Getter method for the COM property "Points"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPoints();


  /**
   * <p>
   * Getter method for the COM property "SegmentType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSegmentType
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.MsoSegmentType getSegmentType();


  // Properties:
}
