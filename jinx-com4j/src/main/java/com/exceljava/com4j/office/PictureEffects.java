package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03D2-0000-0000-C000-000000000046}")
public interface PictureEffects extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PictureEffect
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.PictureEffect getItem(
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
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter position is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(effectType, -1);
   * </code>
   * </p>
   * @param effectType Mandatory com.exceljava.com4j.office.MsoPictureEffectType parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PictureEffect
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.PictureEffect insert(
    com.exceljava.com4j.office.MsoPictureEffectType effectType);

  /**
   * @param effectType Mandatory com.exceljava.com4j.office.MsoPictureEffectType parameter.
   * @param position Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.PictureEffect
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.PictureEffect insert(
    com.exceljava.com4j.office.MsoPictureEffectType effectType,
    @Optional @DefaultValue("-1") int position);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter index is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(-1);
   * </code>
   * </p>
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  void delete();

  /**
   * @param index Optional parameter. Default value is -1
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  void delete(
    @Optional @DefaultValue("-1") int index);


  // Properties:
}
