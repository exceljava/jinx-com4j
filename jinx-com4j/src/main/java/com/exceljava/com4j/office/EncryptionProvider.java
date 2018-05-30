package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CD809-0000-0000-C000-000000000046}")
public interface EncryptionProvider extends Com4jObject {
  // Methods:
  /**
   * @param encprovdet Mandatory com.exceljava.com4j.office.EncryptionProviderDetail parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProviderDetail(
    com.exceljava.com4j.office.EncryptionProviderDetail encprovdet);


  /**
   * @param parentWindow Mandatory com4j.Com4jObject parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  int newSession(
    com4j.Com4jObject parentWindow);


  /**
   * @param parentWindow Mandatory com4j.Com4jObject parameter.
   * @param encryptionData Mandatory com4j.Com4jObject parameter.
   * @param permissionsMask Mandatory Holder<Integer> parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  int authenticate(
    com4j.Com4jObject parentWindow,
    com4j.Com4jObject encryptionData,
    Holder<Integer> permissionsMask);


  /**
   * @param sessionHandle Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  int cloneSession(
    int sessionHandle);


  /**
   * @param sessionHandle Mandatory int parameter.
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  void endSession(
    int sessionHandle);


  /**
   * @param sessionHandle Mandatory int parameter.
   * @param encryptionData Mandatory com4j.Com4jObject parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  int save(
    int sessionHandle,
    com4j.Com4jObject encryptionData);


  /**
   * @param sessionHandle Mandatory int parameter.
   * @param streamName Mandatory java.lang.String parameter.
   * @param unencryptedStream Mandatory com4j.Com4jObject parameter.
   * @param encryptedStream Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  void encryptStream(
    int sessionHandle,
    java.lang.String streamName,
    com4j.Com4jObject unencryptedStream,
    com4j.Com4jObject encryptedStream);


  /**
   * @param sessionHandle Mandatory int parameter.
   * @param streamName Mandatory java.lang.String parameter.
   * @param encryptedStream Mandatory com4j.Com4jObject parameter.
   * @param unencryptedStream Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  void decryptStream(
    int sessionHandle,
    java.lang.String streamName,
    com4j.Com4jObject encryptedStream,
    com4j.Com4jObject unencryptedStream);


  /**
   * @param sessionHandle Mandatory int parameter.
   * @param parentWindow Mandatory com4j.Com4jObject parameter.
   * @param readOnly Mandatory boolean parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  boolean showSettings(
    int sessionHandle,
    com4j.Com4jObject parentWindow,
    boolean readOnly);


  // Properties:
}
