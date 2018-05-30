package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1530-0000-0000-C000-000000000046}")
public interface OfficeDataSourceObject extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "ConnectString"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  java.lang.String getConnectString();


  /**
   * <p>
   * Setter method for the COM property "ConnectString"
   * </p>
   * @param pbstrConnect Mandatory java.lang.String parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(8)
  void setConnectString(
    java.lang.String pbstrConnect);


  /**
   * <p>
   * Getter method for the COM property "Table"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getTable();


  /**
   * <p>
   * Setter method for the COM property "Table"
   * </p>
   * @param pbstrTable Mandatory java.lang.String parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  void setTable(
    java.lang.String pbstrTable);


  /**
   * <p>
   * Getter method for the COM property "DataSource"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getDataSource();


  /**
   * <p>
   * Setter method for the COM property "DataSource"
   * </p>
   * @param pbstrSrc Mandatory java.lang.String parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  void setDataSource(
    java.lang.String pbstrSrc);


  /**
   * <p>
   * Getter method for the COM property "Columns"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getColumns();


  /**
   * <p>
   * Getter method for the COM property "RowCount"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  int getRowCount();


  /**
   * <p>
   * Getter method for the COM property "Filters"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(15)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getFilters();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter rowNbr is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * move(msoMoveRow, 1);
   * </code>
   * </p>
   * @param msoMoveRow Mandatory com.exceljava.com4j.office.MsoMoveRow parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  int move(
    com.exceljava.com4j.office.MsoMoveRow msoMoveRow);

  /**
   * @param msoMoveRow Mandatory com.exceljava.com4j.office.MsoMoveRow parameter.
   * @param rowNbr Optional parameter. Default value is 1
   * @return  Returns a value of type int
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  int move(
    com.exceljava.com4j.office.MsoMoveRow msoMoveRow,
    @Optional @DefaultValue("1") int rowNbr);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter bstrSrc is set to ""</li><li>java.lang.String parameter bstrConnect is set to ""</li><li>java.lang.String parameter bstrTable is set to ""</li><li>int parameter fOpenExclusive is set to 0</li><li>int parameter fNeverPrompt is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open("", "", "", 0, 1);
   * </code>
   * </p>
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.String.class, java.lang.String.class, java.lang.String.class, int.class, int.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.BSTR, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"", "", "", "0", "1"})
  void open();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter bstrConnect is set to ""</li><li>java.lang.String parameter bstrTable is set to ""</li><li>int parameter fOpenExclusive is set to 0</li><li>int parameter fNeverPrompt is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(bstrSrc, "", "", 0, 1);
   * </code>
   * </p>
   * @param bstrSrc Optional parameter. Default value is ""
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.String.class, java.lang.String.class, int.class, int.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"", "", "0", "1"})
  void open(
    @Optional @DefaultValue("") java.lang.String bstrSrc);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter bstrTable is set to ""</li><li>int parameter fOpenExclusive is set to 0</li><li>int parameter fNeverPrompt is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(bstrSrc, bstrConnect, "", 0, 1);
   * </code>
   * </p>
   * @param bstrSrc Optional parameter. Default value is ""
   * @param bstrConnect Optional parameter. Default value is ""
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.String.class, int.class, int.class}, nativeType = {NativeType.BSTR, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"", "0", "1"})
  void open(
    @Optional @DefaultValue("") java.lang.String bstrSrc,
    @Optional @DefaultValue("") java.lang.String bstrConnect);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter fOpenExclusive is set to 0</li><li>int parameter fNeverPrompt is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(bstrSrc, bstrConnect, bstrTable, 0, 1);
   * </code>
   * </p>
   * @param bstrSrc Optional parameter. Default value is ""
   * @param bstrConnect Optional parameter. Default value is ""
   * @param bstrTable Optional parameter. Default value is ""
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"0", "1"})
  void open(
    @Optional @DefaultValue("") java.lang.String bstrSrc,
    @Optional @DefaultValue("") java.lang.String bstrConnect,
    @Optional @DefaultValue("") java.lang.String bstrTable);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter fNeverPrompt is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * open(bstrSrc, bstrConnect, bstrTable, fOpenExclusive, 1);
   * </code>
   * </p>
   * @param bstrSrc Optional parameter. Default value is ""
   * @param bstrConnect Optional parameter. Default value is ""
   * @param bstrTable Optional parameter. Default value is ""
   * @param fOpenExclusive Optional parameter. Default value is 0
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  void open(
    @Optional @DefaultValue("") java.lang.String bstrSrc,
    @Optional @DefaultValue("") java.lang.String bstrConnect,
    @Optional @DefaultValue("") java.lang.String bstrTable,
    @Optional @DefaultValue("0") int fOpenExclusive);

  /**
   * @param bstrSrc Optional parameter. Default value is ""
   * @param bstrConnect Optional parameter. Default value is ""
   * @param bstrTable Optional parameter. Default value is ""
   * @param fOpenExclusive Optional parameter. Default value is 0
   * @param fNeverPrompt Optional parameter. Default value is 1
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  void open(
    @Optional @DefaultValue("") java.lang.String bstrSrc,
    @Optional @DefaultValue("") java.lang.String bstrConnect,
    @Optional @DefaultValue("") java.lang.String bstrTable,
    @Optional @DefaultValue("0") int fOpenExclusive,
    @Optional @DefaultValue("1") int fNeverPrompt);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter sortAscending1 is set to false</li><li>java.lang.String parameter sortField2 is set to ""</li><li>boolean parameter sortAscending2 is set to false</li><li>java.lang.String parameter sortField3 is set to ""</li><li>boolean parameter sortAscending3 is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSortOrder(sortField1, false, "", false, "", false);
   * </code>
   * </p>
   * @param sortField1 Mandatory java.lang.String parameter.
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {boolean.class, java.lang.String.class, boolean.class, java.lang.String.class, boolean.class}, nativeType = {NativeType.VariantBool, NativeType.BSTR, NativeType.VariantBool, NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"false", "", "false", "", "false"})
  void setSortOrder(
    java.lang.String sortField1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter sortField2 is set to ""</li><li>boolean parameter sortAscending2 is set to false</li><li>java.lang.String parameter sortField3 is set to ""</li><li>boolean parameter sortAscending3 is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSortOrder(sortField1, sortAscending1, "", false, "", false);
   * </code>
   * </p>
   * @param sortField1 Mandatory java.lang.String parameter.
   * @param sortAscending1 Optional parameter. Default value is false
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.String.class, boolean.class, java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.VariantBool, NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BOOL, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "false", "", "false"})
  void setSortOrder(
    java.lang.String sortField1,
    @Optional @DefaultValue("-1") boolean sortAscending1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter sortAscending2 is set to false</li><li>java.lang.String parameter sortField3 is set to ""</li><li>boolean parameter sortAscending3 is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSortOrder(sortField1, sortAscending1, sortField2, false, "", false);
   * </code>
   * </p>
   * @param sortField1 Mandatory java.lang.String parameter.
   * @param sortAscending1 Optional parameter. Default value is false
   * @param sortField2 Optional parameter. Default value is ""
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {boolean.class, java.lang.String.class, boolean.class}, nativeType = {NativeType.VariantBool, NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"false", "", "false"})
  void setSortOrder(
    java.lang.String sortField1,
    @Optional @DefaultValue("-1") boolean sortAscending1,
    @Optional @DefaultValue("") java.lang.String sortField2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter sortField3 is set to ""</li><li>boolean parameter sortAscending3 is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSortOrder(sortField1, sortAscending1, sortField2, sortAscending2, "", false);
   * </code>
   * </p>
   * @param sortField1 Mandatory java.lang.String parameter.
   * @param sortAscending1 Optional parameter. Default value is false
   * @param sortField2 Optional parameter. Default value is ""
   * @param sortAscending2 Optional parameter. Default value is false
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "false"})
  void setSortOrder(
    java.lang.String sortField1,
    @Optional @DefaultValue("-1") boolean sortAscending1,
    @Optional @DefaultValue("") java.lang.String sortField2,
    @Optional @DefaultValue("-1") boolean sortAscending2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter sortAscending3 is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSortOrder(sortField1, sortAscending1, sortField2, sortAscending2, sortField3, false);
   * </code>
   * </p>
   * @param sortField1 Mandatory java.lang.String parameter.
   * @param sortAscending1 Optional parameter. Default value is false
   * @param sortField2 Optional parameter. Default value is ""
   * @param sortAscending2 Optional parameter. Default value is false
   * @param sortField3 Optional parameter. Default value is ""
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void setSortOrder(
    java.lang.String sortField1,
    @Optional @DefaultValue("-1") boolean sortAscending1,
    @Optional @DefaultValue("") java.lang.String sortField2,
    @Optional @DefaultValue("-1") boolean sortAscending2,
    @Optional @DefaultValue("") java.lang.String sortField3);

  /**
   * @param sortField1 Mandatory java.lang.String parameter.
   * @param sortAscending1 Optional parameter. Default value is false
   * @param sortField2 Optional parameter. Default value is ""
   * @param sortAscending2 Optional parameter. Default value is false
   * @param sortField3 Optional parameter. Default value is ""
   * @param sortAscending3 Optional parameter. Default value is false
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  void setSortOrder(
    java.lang.String sortField1,
    @Optional @DefaultValue("-1") boolean sortAscending1,
    @Optional @DefaultValue("") java.lang.String sortField2,
    @Optional @DefaultValue("-1") boolean sortAscending2,
    @Optional @DefaultValue("") java.lang.String sortField3,
    @Optional @DefaultValue("-1") boolean sortAscending3);


  /**
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(19)
  void applyFilter();


  // Properties:
}
