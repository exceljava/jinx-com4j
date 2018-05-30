package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E18C-0000-0000-C000-000000000046}")
public interface Property extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(7)
  @DefaultMethod
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param lppvReturn Mandatory java.lang.Object parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(8)
  @DefaultMethod
  void setValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object lppvReturn);


  /**
   * <p>
   * Getter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getIndexedValue(index1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object getIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1);

  /**
   * <p>
   * Getter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getIndexedValue(index1, index2, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object getIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2);

  /**
   * <p>
   * Getter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getIndexedValue(index1, index2, index3, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object getIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index3);

  /**
   * <p>
   * Getter method for the COM property "IndexedValue"
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index4);


  /**
   * <p>
   * Setter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setIndexedValue(index1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), lppvReturn);
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param lppvReturn Mandatory java.lang.Object parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void setIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @MarshalAs(NativeType.VARIANT) java.lang.Object lppvReturn);

  /**
   * <p>
   * Setter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setIndexedValue(index1, index2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), lppvReturn);
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lppvReturn Mandatory java.lang.Object parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void setIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @MarshalAs(NativeType.VARIANT) java.lang.Object lppvReturn);

  /**
   * <p>
   * Setter method for the COM property "IndexedValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index4 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setIndexedValue(index1, index2, index3, com4j.Variant.getMissing(), lppvReturn);
   * </code>
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lppvReturn Mandatory java.lang.Object parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index3,
    @MarshalAs(NativeType.VARIANT) java.lang.Object lppvReturn);

  /**
   * <p>
   * Setter method for the COM property "IndexedValue"
   * </p>
   * @param index1 Mandatory java.lang.Object parameter.
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index4 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lppvReturn Mandatory java.lang.Object parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(10)
  void setIndexedValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index4,
    @MarshalAs(NativeType.VARIANT) java.lang.Object lppvReturn);


  /**
   * <p>
   * Getter method for the COM property "NumIndices"
   * </p>
   * @return  Returns a value of type short
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(11)
  short getNumIndices();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Application
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.vbide.Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._Properties
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.vbide._Properties getParent();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(40) //= 0x28. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(41) //= 0x29. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide.VBE getVBE();


  /**
   * <p>
   * Getter method for the COM property "Collection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._Properties
   */

  @DISPID(42) //= 0x2a. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.vbide._Properties getCollection();


  /**
   * <p>
   * Getter method for the COM property "Object"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(45) //= 0x2d. The runtime will prefer the VTID if present
  @VTID(17)
  com4j.Com4jObject getObject();


  /**
   * <p>
   * Setter method for the COM property "Object"
   * </p>
   * @param lppunk Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(45) //= 0x2d. The runtime will prefer the VTID if present
  @VTID(18)
  void setObject(
    com4j.Com4jObject lppunk);


  // Properties:
}
