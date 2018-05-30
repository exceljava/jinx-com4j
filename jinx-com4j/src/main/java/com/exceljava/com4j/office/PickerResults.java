package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03E5-0000-0000-C000-000000000046}")
public interface PickerResults extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResult
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.PickerResult getItem(
    int index);


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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter sipId is set to ""</li><li>java.lang.Object parameter itemData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subItems is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(id, displayName, type, "", com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param id Mandatory java.lang.String parameter.
   * @param displayName Mandatory java.lang.String parameter.
   * @param type Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResult
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.String.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.BSTR, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.PickerResult add(
    java.lang.String id,
    java.lang.String displayName,
    java.lang.String type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter itemData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter subItems is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(id, displayName, type, sipId, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param id Mandatory java.lang.String parameter.
   * @param displayName Mandatory java.lang.String parameter.
   * @param type Mandatory java.lang.String parameter.
   * @param sipId Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResult
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.PickerResult add(
    java.lang.String id,
    java.lang.String displayName,
    java.lang.String type,
    @Optional @DefaultValue("") java.lang.String sipId);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter subItems is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(id, displayName, type, sipId, itemData, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param id Mandatory java.lang.String parameter.
   * @param displayName Mandatory java.lang.String parameter.
   * @param type Mandatory java.lang.String parameter.
   * @param sipId Optional parameter. Default value is ""
   * @param itemData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResult
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.PickerResult add(
    java.lang.String id,
    java.lang.String displayName,
    java.lang.String type,
    @Optional @DefaultValue("") java.lang.String sipId,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object itemData);

  /**
   * @param id Mandatory java.lang.String parameter.
   * @param displayName Mandatory java.lang.String parameter.
   * @param type Mandatory java.lang.String parameter.
   * @param sipId Optional parameter. Default value is ""
   * @param itemData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param subItems Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResult
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.PickerResult add(
    java.lang.String id,
    java.lang.String displayName,
    java.lang.String type,
    @Optional @DefaultValue("") java.lang.String sipId,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object itemData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subItems);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
