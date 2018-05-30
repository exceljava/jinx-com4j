package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E17A-0000-0000-C000-000000000046}")
public interface _References extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.vbide._VBProject getParent();


  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide.VBE getVBE();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide.Reference
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.vbide.Reference item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param guid Mandatory java.lang.String parameter.
   * @param major Mandatory int parameter.
   * @param minor Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide.Reference
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.vbide.Reference addFromGuid(
    java.lang.String guid,
    int major,
    int minor);


  /**
   * @param fileName Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide.Reference
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.vbide.Reference addFromFile(
    java.lang.String fileName);


  /**
   * @param reference Mandatory com.exceljava.com4j.vbide.Reference parameter.
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  void remove(
    com.exceljava.com4j.vbide.Reference reference);


  // Properties:
}
