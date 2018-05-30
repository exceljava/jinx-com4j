package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208AA-0001-0000-C000-000000000046}")
public interface IRoutingSlip extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Delivery"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlRoutingSlipDelivery
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlRoutingSlipDelivery getDelivery();


  /**
   * <p>
   * Setter method for the COM property "Delivery"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRoutingSlipDelivery parameter.
   */

  @VTID(11)
  void setDelivery(
    com.exceljava.com4j.excel.XlRoutingSlipDelivery rhs);


  /**
   * <p>
   * Getter method for the COM property "Message"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getMessage();


  /**
   * <p>
   * Setter method for the COM property "Message"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(13)
  void setMessage(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Recipients"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRecipients(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getRecipients();

  /**
   * <p>
   * Getter method for the COM property "Recipients"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRecipients(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "Recipients"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setRecipients(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setRecipients(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Recipients"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setRecipients(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(16)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object reset();


  /**
   * <p>
   * Getter method for the COM property "ReturnWhenDone"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getReturnWhenDone();


  /**
   * <p>
   * Setter method for the COM property "ReturnWhenDone"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setReturnWhenDone(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Status"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlRoutingSlipStatus
   */

  @VTID(19)
  com.exceljava.com4j.excel.XlRoutingSlipStatus getStatus();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(20)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSubject();


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(21)
  void setSubject(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TrackStatus"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getTrackStatus();


  /**
   * <p>
   * Setter method for the COM property "TrackStatus"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setTrackStatus(
    boolean rhs);


  // Properties:
}
