package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0365-0000-0000-C000-000000000046}")
public interface FileDialogFilters extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(10)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  int getCount();


  /**
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.FileDialogFilter
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.office.FileDialogFilter item(
    int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filter is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void delete();

  /**
   * @param filter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  void delete(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filter);


  /**
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  void clear();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter position is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(description, extensions, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param description Mandatory java.lang.String parameter.
   * @param extensions Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.FileDialogFilter
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.FileDialogFilter add(
    java.lang.String description,
    java.lang.String extensions);

  /**
   * @param description Mandatory java.lang.String parameter.
   * @param extensions Mandatory java.lang.String parameter.
   * @param position Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.FileDialogFilter
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.FileDialogFilter add(
    java.lang.String description,
    java.lang.String extensions,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object position);


  // Properties:
}
