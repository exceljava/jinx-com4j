package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C037C-0000-0000-C000-000000000046}")
public interface SharedWorkspaceFiles extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(9)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFile
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.office.SharedWorkspaceFile getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(11)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter parentFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter overwriteIfFileAlreadyExists is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter keepInSync is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(fileName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFile
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.SharedWorkspaceFile add(
    java.lang.String fileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwriteIfFileAlreadyExists is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter keepInSync is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(fileName, parentFolder, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param parentFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFile
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.SharedWorkspaceFile add(
    java.lang.String fileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentFolder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter keepInSync is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(fileName, parentFolder, overwriteIfFileAlreadyExists, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param parentFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param overwriteIfFileAlreadyExists Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFile
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.SharedWorkspaceFile add(
    java.lang.String fileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwriteIfFileAlreadyExists);

  /**
   * @param fileName Mandatory java.lang.String parameter.
   * @param parentFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param overwriteIfFileAlreadyExists Optional parameter. Default value is com4j.Variant.getMissing()
   * @param keepInSync Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.SharedWorkspaceFile
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.SharedWorkspaceFile add(
    java.lang.String fileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object overwriteIfFileAlreadyExists,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object keepInSync);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "ItemCountExceeded"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  boolean getItemCountExceeded();


  // Properties:
}
