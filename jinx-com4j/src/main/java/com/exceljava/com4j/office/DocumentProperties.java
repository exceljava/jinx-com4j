package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{2DF8D04D-5BFA-101B-BDE5-00AA0044DE52}")
public interface DocumentProperties extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @VTID(7)
  void getParent();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getItem(index, 1033);
   * </code>
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(8)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.DocumentProperty getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(8)
  @DefaultMethod
  com.exceljava.com4j.office.DocumentProperty getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(9)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter value is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, linkToContent, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkToContent Mandatory boolean parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.DocumentProperty add(
    java.lang.String name,
    boolean linkToContent);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter value is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, linkToContent, type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkToContent Mandatory boolean parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.DocumentProperty add(
    java.lang.String name,
    boolean linkToContent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, linkToContent, type, value, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkToContent Mandatory boolean parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.DocumentProperty add(
    java.lang.String name,
    boolean linkToContent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, linkToContent, type, value, linkSource, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param linkToContent Mandatory boolean parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=6)
  com.exceljava.com4j.office.DocumentProperty add(
    java.lang.String name,
    boolean linkToContent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param linkToContent Mandatory boolean parameter.
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param value Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DocumentProperty
   */

  @VTID(10)
  com.exceljava.com4j.office.DocumentProperty add(
    java.lang.String name,
    boolean linkToContent,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getCreator();


  // Properties:
}
