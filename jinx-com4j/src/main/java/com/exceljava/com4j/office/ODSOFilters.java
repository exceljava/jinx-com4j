package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1534-0000-0000-C000-000000000046}")
public interface ODSOFilters extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject item(
    int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter bstrCompareTo is set to ""</li><li>boolean parameter deferUpdate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(column, comparison, conjunction, "", false);
   * </code>
   * </p>
   * @param column Mandatory java.lang.String parameter.
   * @param comparison Mandatory com.exceljava.com4j.office.MsoFilterComparison parameter.
   * @param conjunction Mandatory com.exceljava.com4j.office.MsoFilterConjunction parameter.
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "false"})
  void add(
    java.lang.String column,
    com.exceljava.com4j.office.MsoFilterComparison comparison,
    com.exceljava.com4j.office.MsoFilterConjunction conjunction);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter deferUpdate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(column, comparison, conjunction, bstrCompareTo, false);
   * </code>
   * </p>
   * @param column Mandatory java.lang.String parameter.
   * @param comparison Mandatory com.exceljava.com4j.office.MsoFilterComparison parameter.
   * @param conjunction Mandatory com.exceljava.com4j.office.MsoFilterConjunction parameter.
   * @param bstrCompareTo Optional parameter. Default value is ""
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void add(
    java.lang.String column,
    com.exceljava.com4j.office.MsoFilterComparison comparison,
    com.exceljava.com4j.office.MsoFilterConjunction conjunction,
    @Optional @DefaultValue("") java.lang.String bstrCompareTo);

  /**
   * @param column Mandatory java.lang.String parameter.
   * @param comparison Mandatory com.exceljava.com4j.office.MsoFilterComparison parameter.
   * @param conjunction Mandatory com.exceljava.com4j.office.MsoFilterConjunction parameter.
   * @param bstrCompareTo Optional parameter. Default value is ""
   * @param deferUpdate Optional parameter. Default value is false
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  void add(
    java.lang.String column,
    com.exceljava.com4j.office.MsoFilterComparison comparison,
    com.exceljava.com4j.office.MsoFilterConjunction conjunction,
    @Optional @DefaultValue("") java.lang.String bstrCompareTo,
    @Optional @DefaultValue("0") boolean deferUpdate);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter deferUpdate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(index, false);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void delete(
    int index);

  /**
   * @param index Mandatory int parameter.
   * @param deferUpdate Optional parameter. Default value is false
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  void delete(
    int index,
    @Optional @DefaultValue("0") boolean deferUpdate);


  // Properties:
}
