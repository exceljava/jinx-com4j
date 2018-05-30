package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{EC0E6191-DB51-11D3-8F3E-00C04F3651B8}")
public interface IRtdServer extends Com4jObject {
  // Methods:
  /**
   * @param callbackObject Mandatory com.exceljava.com4j.excel.IRTDUpdateEvent parameter.
   * @return  Returns a value of type int
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(7)
  int serverStart(
    com.exceljava.com4j.excel.IRTDUpdateEvent callbackObject);


  /**
   * @param topicID Mandatory int parameter.
   * @param strings Mandatory Holder<java.lang.Object[]> parameter.
   * @param getNewValues Mandatory Holder<Boolean> parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(8)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object connectData(
    int topicID,
    Holder<java.lang.Object[]> strings,
    Holder<Boolean> getNewValues);


  /**
   * @param topicCount Mandatory Holder<Integer> parameter.
   * @return  Returns a value of type java.lang.Object[]
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.Object[] refreshData(
    Holder<Integer> topicCount);


  /**
   * @param topicID Mandatory int parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(10)
  void disconnectData(
    int topicID);


  /**
   * @return  Returns a value of type int
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(11)
  int heartbeat();


  /**
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(12)
  void serverTerminate();


  // Properties:
}
