package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0356-0000-0000-C000-000000000046}")
public interface HTMLProject extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "State"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoHTMLProjectState
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.MsoHTMLProjectState getState();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter refresh is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * refreshProject(false);
   * </code>
   * </p>
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void refreshProject();

  /**
   * @param refresh Optional parameter. Default value is false
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  void refreshProject(
    @Optional @DefaultValue("-1") boolean refresh);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter refresh is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * refreshDocument(false);
   * </code>
   * </p>
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void refreshDocument();

  /**
   * @param refresh Optional parameter. Default value is false
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void refreshDocument(
    @Optional @DefaultValue("-1") boolean refresh);


  /**
   * <p>
   * Getter method for the COM property "HTMLProjectItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.HTMLProjectItems
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.HTMLProjectItems getHTMLProjectItems();


  @VTID(12)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.HTMLProjectItems.class})
  com.exceljava.com4j.office.HTMLProjectItem getHTMLProjectItems(
    java.lang.Object index);

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


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoHTMLProjectOpen parameter openKind is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(0);
   * </code>
   * </p>
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.office.MsoHTMLProjectOpen.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void open();

  /**
   * @param openKind Optional parameter. Default value is 0
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  void open(
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoHTMLProjectOpen openKind);


  // Properties:
}
