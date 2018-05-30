package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{4291224C-DEFE-485B-8E69-6CF8AA85CB76}")
public interface IAssistance extends Com4jObject {
  // Methods:
  /**
   * <p>
   * ShowHelp Method
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter helpId is set to ""</li><li>java.lang.String parameter scope is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showHelp("", "");
   * </code>
   * </p>
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.String.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR}, literal = {"", ""})
  void showHelp();

  /**
   * <p>
   * ShowHelp Method
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter scope is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showHelp(helpId, "");
   * </code>
   * </p>
   * @param helpId Optional parameter. Default value is ""
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  void showHelp(
    @Optional @DefaultValue("") java.lang.String helpId);

  /**
   * <p>
   * ShowHelp Method
   * </p>
   * @param helpId Optional parameter. Default value is ""
   * @param scope Optional parameter. Default value is ""
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  void showHelp(
    @Optional @DefaultValue("") java.lang.String helpId,
    @Optional @DefaultValue("") java.lang.String scope);


  /**
   * <p>
   * SearchHelp Method
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter scope is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * searchHelp(query, "");
   * </code>
   * </p>
   * @param query Mandatory java.lang.String parameter.
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  void searchHelp(
    java.lang.String query);

  /**
   * <p>
   * SearchHelp Method
   * </p>
   * @param query Mandatory java.lang.String parameter.
   * @param scope Optional parameter. Default value is ""
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  void searchHelp(
    java.lang.String query,
    @Optional @DefaultValue("") java.lang.String scope);


  /**
   * <p>
   * SetDefaultContext Method
   * </p>
   * @param helpId Mandatory java.lang.String parameter.
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  void setDefaultContext(
    java.lang.String helpId);


  /**
   * <p>
   * ClearDefaultContext Method
   * </p>
   * @param helpId Mandatory java.lang.String parameter.
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  void clearDefaultContext(
    java.lang.String helpId);


  // Properties:
}
