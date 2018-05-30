package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C0-0000-0000-C000-000000000046}")
public interface GradientStops extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.GradientStop
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(9)
  @DefaultMethod
  com.exceljava.com4j.office.GradientStop getItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
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
     * <li>int parameter index is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(-1);
   * </code>
   * </p>
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  void delete();

  /**
   * @param index Optional parameter. Default value is -1
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(12)
  void delete(
    @Optional @DefaultValue("-1") int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter transparency is set to 0.0f</li><li>int parameter index is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(rgb, position, 0.0f, -1);
   * </code>
   * </p>
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {float.class, int.class}, nativeType = {NativeType.Float, NativeType.Int32}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_I4}, literal = {"0.0", "-1"})
  void insert(
    int rgb,
    float position);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter index is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(rgb, position, transparency, -1);
   * </code>
   * </p>
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   * @param transparency Optional parameter. Default value is 0.0f
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  void insert(
    int rgb,
    float position,
    @Optional @DefaultValue("0") float transparency);

  /**
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   * @param transparency Optional parameter. Default value is 0.0f
   * @param index Optional parameter. Default value is -1
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(13)
  void insert(
    int rgb,
    float position,
    @Optional @DefaultValue("0") float transparency,
    @Optional @DefaultValue("-1") int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter transparency is set to 0.0f</li><li>int parameter index is set to -1</li><li>float parameter brightness is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert2(rgb, position, 0.0f, -1, 0.0f);
   * </code>
   * </p>
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {float.class, int.class, float.class}, nativeType = {NativeType.Float, NativeType.Int32, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_I4, Variant.Type.VT_R4}, literal = {"0.0", "-1", "0.0"})
  void insert2(
    int rgb,
    float position);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter index is set to -1</li><li>float parameter brightness is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert2(rgb, position, transparency, -1, 0.0f);
   * </code>
   * </p>
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   * @param transparency Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {int.class, float.class}, nativeType = {NativeType.Int32, NativeType.Float}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_R4}, literal = {"-1", "0.0"})
  void insert2(
    int rgb,
    float position,
    @Optional @DefaultValue("0") float transparency);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter brightness is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert2(rgb, position, transparency, index, 0.0f);
   * </code>
   * </p>
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   * @param transparency Optional parameter. Default value is 0.0f
   * @param index Optional parameter. Default value is -1
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"0.0"})
  void insert2(
    int rgb,
    float position,
    @Optional @DefaultValue("0") float transparency,
    @Optional @DefaultValue("-1") int index);

  /**
   * @param rgb Mandatory int parameter.
   * @param position Mandatory float parameter.
   * @param transparency Optional parameter. Default value is 0.0f
   * @param index Optional parameter. Default value is -1
   * @param brightness Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  void insert2(
    int rgb,
    float position,
    @Optional @DefaultValue("0") float transparency,
    @Optional @DefaultValue("-1") int index,
    @Optional @DefaultValue("0") float brightness);


  // Properties:
}
