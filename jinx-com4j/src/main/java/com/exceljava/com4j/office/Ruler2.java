package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C1-0000-0000-C000-000000000046}")
public interface Ruler2 extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "Levels"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.RulerLevels2
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.RulerLevels2 getLevels();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.RulerLevels2.class})
  com.exceljava.com4j.office.RulerLevel2 getLevels(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "TabStops"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TabStops2
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.TabStops2 getTabStops();


  @VTID(11)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.TabStops2.class})
  com.exceljava.com4j.office.TabStop2 getTabStops(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  // Properties:
}
