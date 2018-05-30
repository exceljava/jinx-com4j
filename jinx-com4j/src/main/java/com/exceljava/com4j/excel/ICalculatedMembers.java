package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024454-0001-0000-C000-000000000046}")
public interface ICalculatedMembers extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(11)
  com.exceljava.com4j.excel.CalculatedMember getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.CalculatedMember get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter solveOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(name, formula, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.CalculatedMember _Add(
    java.lang.String name,
    java.lang.String formula);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(name, formula, solveOrder, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.CalculatedMember _Add(
    java.lang.String name,
    java.lang.String formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(14)
  com.exceljava.com4j.excel.CalculatedMember _Add(
    java.lang.String name,
    java.lang.String formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter solveOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dynamic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hierarchizeDistinct is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dynamic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hierarchizeDistinct is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, solveOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dynamic is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hierarchizeDistinct is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, solveOrder, type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hierarchizeDistinct is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, solveOrder, type, dynamic, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dynamic Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dynamic);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hierarchizeDistinct is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, solveOrder, type, dynamic, displayFolder, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dynamic Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dynamic,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dynamic Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hierarchizeDistinct Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(15)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dynamic,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hierarchizeDistinct);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter solveOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter measureGroup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentHierarchy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 9}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter measureGroup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentHierarchy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 9}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter displayFolder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter measureGroup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentHierarchy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 9}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter measureGroup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentHierarchy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, type, displayFolder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 9}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter parentHierarchy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, type, displayFolder, measureGroup, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param measureGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 9}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object measureGroup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter parentMember is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param measureGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param parentHierarchy Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 9}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object measureGroup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentHierarchy);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter numberFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addCalculatedMember(name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, parentMember, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param measureGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param parentHierarchy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param parentMember Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 9}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=9)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object measureGroup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentHierarchy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentMember);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param measureGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param parentHierarchy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param parentMember Optional parameter. Default value is com4j.Variant.getMissing()
   * @param numberFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.CalculatedMember
   */

  @VTID(16)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    @MarshalAs(NativeType.VARIANT) java.lang.Object formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object solveOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayFolder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object measureGroup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentHierarchy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object parentMember,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object numberFormat);


  // Properties:
}
