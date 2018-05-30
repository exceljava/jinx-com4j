package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ListObjects extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>com.exceljava.com4j.excel.XlListObjectSourceType parameter sourceType is set to 1</li><li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlListObjectSourceType.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "0", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject _Add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "0", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(sourceType, source, com4j.Variant.getMissing(), 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "0", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(sourceType, source, linkSource, 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(sourceType, source, linkSource, xlListObjectHasHeaders, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders);

  /**
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2085)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional java.lang.Object destination);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.ListObject get_Default(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4)
  @PropGet
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.ListObject getItem(
    java.lang.Object index);


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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlListObjectSourceType parameter sourceType is set to 1</li><li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 0, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlListObjectSourceType.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 0, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, source, com4j.Variant.getMissing(), 0, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, source, linkSource, 0, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"0", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, source, linkSource, xlListObjectHasHeaders, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, source, linkSource, xlListObjectHasHeaders, destination, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional java.lang.Object destination);

  /**
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableStyleName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional java.lang.Object source,
    @Optional java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional java.lang.Object destination,
    @Optional java.lang.Object tableStyleName);


  // Properties:
}
