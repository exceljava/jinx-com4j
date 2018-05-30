package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface CalculatedMembers extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   */

  @DISPID(118)
  @PropGet
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(170)
  @PropGet
  com.exceljava.com4j.excel.CalculatedMember getItem(
    java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  com.exceljava.com4j.excel.CalculatedMember get_Default(
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
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
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
   */

  @DISPID(2085)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember _Add(
    java.lang.String name,
    java.lang.String formula,
    @Optional java.lang.Object solveOrder);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2085)
  com.exceljava.com4j.excel.CalculatedMember _Add(
    java.lang.String name,
    java.lang.String formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type);


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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object dynamic);

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
   */

  @DISPID(181)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object dynamic,
    @Optional java.lang.Object displayFolder);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.Object parameter.
   * @param solveOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dynamic Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayFolder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hierarchizeDistinct Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(181)
  com.exceljava.com4j.excel.CalculatedMember add(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object dynamic,
    @Optional java.lang.Object displayFolder,
    @Optional java.lang.Object hierarchizeDistinct);


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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object displayFolder);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object displayFolder,
    @Optional java.lang.Object measureGroup);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object displayFolder,
    @Optional java.lang.Object measureGroup,
    @Optional java.lang.Object parentHierarchy);

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
   */

  @DISPID(3091)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object displayFolder,
    @Optional java.lang.Object measureGroup,
    @Optional java.lang.Object parentHierarchy,
    @Optional java.lang.Object parentMember);

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
   */

  @DISPID(3091)
  com.exceljava.com4j.excel.CalculatedMember addCalculatedMember(
    java.lang.String name,
    java.lang.Object formula,
    @Optional java.lang.Object solveOrder,
    @Optional java.lang.Object type,
    @Optional java.lang.Object displayFolder,
    @Optional java.lang.Object measureGroup,
    @Optional java.lang.Object parentHierarchy,
    @Optional java.lang.Object parentMember,
    @Optional java.lang.Object numberFormat);


  // Properties:
}
