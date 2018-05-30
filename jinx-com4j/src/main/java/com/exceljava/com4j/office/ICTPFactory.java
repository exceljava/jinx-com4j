package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C033D-0000-0000-C000-000000000046}")
public interface ICTPFactory extends Com4jObject {
  // Methods:
  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ctpParentWindow is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createCTP(ctpAxID, ctpTitle, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param ctpAxID Mandatory java.lang.String parameter.
   * @param ctpTitle Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office._CustomTaskPane
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office._CustomTaskPane createCTP(
    java.lang.String ctpAxID,
    java.lang.String ctpTitle);

  /**
   * @param ctpAxID Mandatory java.lang.String parameter.
   * @param ctpTitle Mandatory java.lang.String parameter.
   * @param ctpParentWindow Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office._CustomTaskPane
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.office._CustomTaskPane createCTP(
    java.lang.String ctpAxID,
    java.lang.String ctpTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ctpParentWindow);


  // Properties:
}
