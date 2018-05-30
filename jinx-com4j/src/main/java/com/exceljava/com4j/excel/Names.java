package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Names extends Com4jObject,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category,
    @Optional java.lang.Object nameLocal);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category,
    @Optional java.lang.Object nameLocal,
    @Optional java.lang.Object refersToLocal);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category,
    @Optional java.lang.Object nameLocal,
    @Optional java.lang.Object refersToLocal,
    @Optional java.lang.Object categoryLocal);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category,
    @Optional java.lang.Object nameLocal,
    @Optional java.lang.Object refersToLocal,
    @Optional java.lang.Object categoryLocal,
    @Optional java.lang.Object refersToR1C1);

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
   */

  @DISPID(181)
  com.exceljava.com4j.excel.Name add(
    @Optional java.lang.Object name,
    @Optional java.lang.Object refersTo,
    @Optional java.lang.Object visible,
    @Optional java.lang.Object macroType,
    @Optional java.lang.Object shortcutKey,
    @Optional java.lang.Object category,
    @Optional java.lang.Object nameLocal,
    @Optional java.lang.Object refersToLocal,
    @Optional java.lang.Object categoryLocal,
    @Optional java.lang.Object refersToR1C1,
    @Optional java.lang.Object refersToR1C1Local);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(170)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name item();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(index, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(170)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name item(
    @Optional java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * item(index, indexLocal, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(170)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name item(
    @Optional java.lang.Object index,
    @Optional java.lang.Object indexLocal);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(170)
  com.exceljava.com4j.excel.Name item(
    @Optional java.lang.Object index,
    @Optional java.lang.Object indexLocal,
    @Optional java.lang.Object refersTo);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(0)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name _Default();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter indexLocal is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(index, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name _Default(
    @Optional java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter refersTo is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Default(index, indexLocal, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @DefaultMethod
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.Name _Default(
    @Optional java.lang.Object index,
    @Optional java.lang.Object indexLocal);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param indexLocal Optional parameter. Default value is com4j.Variant.getMissing()
   * @param refersTo Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(0)
  @DefaultMethod
  com.exceljava.com4j.excel.Name _Default(
    @Optional java.lang.Object index,
    @Optional java.lang.Object indexLocal,
    @Optional java.lang.Object refersTo);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
