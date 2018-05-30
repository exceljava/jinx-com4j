package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface RoutingSlip extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Delivery"
   * </p>
   */

  @DISPID(955)
  @PropGet
  com.exceljava.com4j.excel.XlRoutingSlipDelivery getDelivery();


  /**
   * <p>
   * Setter method for the COM property "Delivery"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRoutingSlipDelivery parameter.
   */

  @DISPID(955)
  @PropPut
  void setDelivery(
    com.exceljava.com4j.excel.XlRoutingSlipDelivery rhs);


  /**
   * <p>
   * Getter method for the COM property "Message"
   * </p>
   */

  @DISPID(954)
  @PropGet
  java.lang.Object getMessage();


  /**
   * <p>
   * Setter method for the COM property "Message"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(954)
  @PropPut
  void setMessage(
    java.lang.Object rhs);


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
   */

  @DISPID(952)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getRecipients();

  /**
   * <p>
   * Getter method for the COM property "Recipients"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(952)
  @PropGet
  java.lang.Object getRecipients(
    @Optional java.lang.Object index);


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

  @DISPID(952)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setRecipients(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Recipients"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(952)
  @PropPut
  void setRecipients(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   */

  @DISPID(555)
  java.lang.Object reset();


  /**
   * <p>
   * Getter method for the COM property "ReturnWhenDone"
   * </p>
   */

  @DISPID(956)
  @PropGet
  boolean getReturnWhenDone();


  /**
   * <p>
   * Setter method for the COM property "ReturnWhenDone"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(956)
  @PropPut
  void setReturnWhenDone(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Status"
   * </p>
   */

  @DISPID(958)
  @PropGet
  com.exceljava.com4j.excel.XlRoutingSlipStatus getStatus();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   */

  @DISPID(953)
  @PropGet
  java.lang.Object getSubject();


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(953)
  @PropPut
  void setSubject(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TrackStatus"
   * </p>
   */

  @DISPID(957)
  @PropGet
  boolean getTrackStatus();


  /**
   * <p>
   * Setter method for the COM property "TrackStatus"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(957)
  @PropPut
  void setTrackStatus(
    boolean rhs);


  // Properties:
}
