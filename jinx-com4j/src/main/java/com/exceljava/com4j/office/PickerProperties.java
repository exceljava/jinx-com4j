package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03E3-0000-0000-C000-000000000046}")
public interface PickerProperties extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PickerProperty
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.PickerProperty getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.String parameter.
   * @param type Mandatory com.exceljava.com4j.office.MsoPickerField parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PickerProperty
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.PickerProperty add(
    java.lang.String id,
    java.lang.String value,
    com.exceljava.com4j.office.MsoPickerField type);


  /**
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  void remove(
    java.lang.String id);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
