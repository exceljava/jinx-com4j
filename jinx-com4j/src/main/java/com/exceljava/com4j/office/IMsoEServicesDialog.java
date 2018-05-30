package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0372-0000-0000-C000-000000000046}")
public interface IMsoEServicesDialog extends Com4jObject {
  // Methods:
  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter applyWebComponentChanges is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * close(false);
   * </code>
   * </p>
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void close();

  /**
   * @param applyWebComponentChanges Optional parameter. Default value is false
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  void close(
    @Optional @DefaultValue("0") boolean applyWebComponentChanges);


  /**
   * @param domain Mandatory java.lang.String parameter.
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  void addTrustedDomain(
    java.lang.String domain);


  /**
   * <p>
   * Getter method for the COM property "ApplicationName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getApplicationName();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "WebComponent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getWebComponent();


  /**
   * <p>
   * Getter method for the COM property "ClipArt"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getClipArt();


  // Properties:
}
