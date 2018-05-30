package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0393-0000-0000-C000-000000000046}")
public interface DocumentInspector extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getDescription();


  /**
   * @param status Mandatory Holder<com.exceljava.com4j.office.MsoDocInspectorStatus> parameter.
   * @param results Mandatory Holder<java.lang.String> parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void inspect(
    Holder<com.exceljava.com4j.office.MsoDocInspectorStatus> status,
    Holder<java.lang.String> results);


  /**
   * @param status Mandatory Holder<com.exceljava.com4j.office.MsoDocInspectorStatus> parameter.
   * @param results Mandatory Holder<java.lang.String> parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  void fix(
    Holder<com.exceljava.com4j.office.MsoDocInspectorStatus> status,
    Holder<java.lang.String> results);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  // Properties:
}
