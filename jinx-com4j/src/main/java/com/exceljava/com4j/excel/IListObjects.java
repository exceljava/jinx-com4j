package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024470-0001-0000-C000-000000000046}")
public interface IListObjects extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>com.exceljava.com4j.excel.XlListObjectSourceType parameter sourceType is set to 1</li><li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter linkSource is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlYesNoGuess parameter xlListObjectHasHeaders is set to 0</li><li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 0, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlListObjectSourceType.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "0", "80020004"})
  @ReturnValue(index=5)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "0", "80020004"})
  @ReturnValue(index=5)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"80020004", "0", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR}, literal = {"0", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders);

  /**
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(10)
  com.exceljava.com4j.excel.ListObject _Add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.excel.ListObject get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(13)
  com.exceljava.com4j.excel.ListObject getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {6}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlListObjectSourceType.class, java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"1", "80020004", "80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=6)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 6}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=6)
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 6}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "0", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 6}, optParamIndex = {3, 4, 5}, javaType = {com.exceljava.com4j.excel.XlYesNoGuess.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"0", "80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource);

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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 6}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 6}, optParamIndex = {5}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=6)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);

  /**
   * @param sourceType Optional parameter. Default value is 1
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param linkSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlListObjectHasHeaders Optional parameter. Default value is 0
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableStyleName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(15)
  com.exceljava.com4j.excel.ListObject add(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlListObjectSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object linkSource,
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlYesNoGuess xlListObjectHasHeaders,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableStyleName);


  // Properties:
}
