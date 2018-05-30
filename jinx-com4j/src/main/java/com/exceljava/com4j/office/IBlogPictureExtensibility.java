package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C5-0000-0000-C000-000000000046}")
public interface IBlogPictureExtensibility extends Com4jObject {
  // Methods:
  /**
   * @param blogPictureProvider Mandatory Holder<java.lang.String> parameter.
   * @param friendlyName Mandatory Holder<java.lang.String> parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void blogPictureProviderProperties(
    Holder<java.lang.String> blogPictureProvider,
    Holder<java.lang.String> friendlyName);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param blogProvider Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  void createPictureAccount(
    java.lang.String account,
    java.lang.String blogProvider,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param image Mandatory com4j.Com4jObject parameter.
   * @param imageType Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(index=4)
  java.lang.String publishPicture(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    com4j.Com4jObject image,
    int imageType);


  // Properties:
}
