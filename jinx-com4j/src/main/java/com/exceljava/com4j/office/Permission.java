package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0376-0000-0000-C000-000000000046}")
public interface Permission extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.UserPermission
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.UserPermission getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "EnableTrustedBrowser"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  boolean getEnableTrustedBrowser();


  /**
   * <p>
   * Setter method for the COM property "EnableTrustedBrowser"
   * </p>
   * @param enable Mandatory boolean parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  void setEnableTrustedBrowser(
    boolean enable);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter permission is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter expirationDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(userId, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param userId Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.UserPermission
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.UserPermission add(
    java.lang.String userId);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter expirationDate is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(userId, permission, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param userId Mandatory java.lang.String parameter.
   * @param permission Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.UserPermission
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.UserPermission add(
    java.lang.String userId,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object permission);

  /**
   * @param userId Mandatory java.lang.String parameter.
   * @param permission Optional parameter. Default value is com4j.Variant.getMissing()
   * @param expirationDate Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.UserPermission
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.UserPermission add(
    java.lang.String userId,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object permission,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object expirationDate);


  /**
   * @param fileName Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  void applyPolicy(
    java.lang.String fileName);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  void removeAll();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param enabled Mandatory boolean parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  void setEnabled(
    boolean enabled);


  /**
   * <p>
   * Getter method for the COM property "RequestPermissionURL"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  java.lang.String getRequestPermissionURL();


  /**
   * <p>
   * Setter method for the COM property "RequestPermissionURL"
   * </p>
   * @param contact Mandatory java.lang.String parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(20)
  void setRequestPermissionURL(
    java.lang.String contact);


  /**
   * <p>
   * Getter method for the COM property "PolicyName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getPolicyName();


  /**
   * <p>
   * Getter method for the COM property "PolicyDescription"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(22)
  java.lang.String getPolicyDescription();


  /**
   * <p>
   * Getter method for the COM property "StoreLicenses"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(23)
  boolean getStoreLicenses();


  /**
   * <p>
   * Setter method for the COM property "StoreLicenses"
   * </p>
   * @param enabled Mandatory boolean parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(24)
  void setStoreLicenses(
    boolean enabled);


  /**
   * <p>
   * Getter method for the COM property "DocumentAuthor"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(25)
  java.lang.String getDocumentAuthor();


  /**
   * <p>
   * Setter method for the COM property "DocumentAuthor"
   * </p>
   * @param author Mandatory java.lang.String parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(26)
  void setDocumentAuthor(
    java.lang.String author);


  /**
   * <p>
   * Getter method for the COM property "PermissionFromPolicy"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(27)
  boolean getPermissionFromPolicy();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(28)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
