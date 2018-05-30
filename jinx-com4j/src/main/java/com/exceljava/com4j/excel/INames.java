package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208B8-0001-0000-C000-000000000046}")
public interface INames extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter macroType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {11}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter macroType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 11}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter macroType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 11}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter macroType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 11}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter shortcutKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 11}, optParamIndex = {4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter category is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 11}, optParamIndex = {5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter nameLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, category, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 11}, optParamIndex = {6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersToLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 11}, optParamIndex = {7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameLocal);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 11}, optParamIndex = {8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToLocal);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersToR1C1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11}, optParamIndex = {9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLocal);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersToR1C1Local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToR1C1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11}, optParamIndex = {10}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToR1C1);

  /**
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param macroType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param shortcutKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param category Optional parameter. Default value is com4j.Variant.getMissing()
   * @param nameLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToR1C1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersToR1C1Local Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(10)
  com.exceljava.com4j.excel.Name add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object macroType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object shortcutKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object category,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object nameLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToR1C1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersToR1C1Local);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name item();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(index, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name item(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(index, indexLocal, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name item(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(index, indexLocal, refersTo, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name item(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(11)
  com.exceljava.com4j.excel.Name item(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(12)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name _Default();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(index, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(12)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name _Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(index, indexLocal, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(12)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name _Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(index, indexLocal, refersTo, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(12)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.Name _Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Name
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.Name _Default(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object indexLocal,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object refersTo,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
