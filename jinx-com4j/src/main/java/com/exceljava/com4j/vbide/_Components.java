package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E161-0000-0000-C000-000000000046}")
public interface _Components extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._Component
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(7)
  @DefaultMethod
  com.exceljava.com4j.vbide._Component item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Application
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide.Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(9)
  com.exceljava.com4j.vbide._VBProject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param component Mandatory com.exceljava.com4j.vbide._Component parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(12)
  void remove(
    com.exceljava.com4j.vbide._Component component);


  /**
   * @param componentType Mandatory com.exceljava.com4j.vbide.vbext_ComponentType parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._Component
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.vbide._Component add(
    com.exceljava.com4j.vbide.vbext_ComponentType componentType);


  /**
   * @param fileName Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._Component
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.vbide._Component _import(
    java.lang.String fileName);


  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide.VBE getVBE();


  // Properties:
}
