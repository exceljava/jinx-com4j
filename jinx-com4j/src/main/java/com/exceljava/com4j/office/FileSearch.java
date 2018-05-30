package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0332-0000-0000-C000-000000000046}")
public interface FileSearch extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "SearchSubFolders"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  boolean getSearchSubFolders();


  /**
   * <p>
   * Setter method for the COM property "SearchSubFolders"
   * </p>
   * @param searchSubFoldersRetVal Mandatory boolean parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  void setSearchSubFolders(
    boolean searchSubFoldersRetVal);


  /**
   * <p>
   * Getter method for the COM property "MatchTextExactly"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  boolean getMatchTextExactly();


  /**
   * <p>
   * Setter method for the COM property "MatchTextExactly"
   * </p>
   * @param matchTextRetVal Mandatory boolean parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  void setMatchTextExactly(
    boolean matchTextRetVal);


  /**
   * <p>
   * Getter method for the COM property "MatchAllWordForms"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  boolean getMatchAllWordForms();


  /**
   * <p>
   * Setter method for the COM property "MatchAllWordForms"
   * </p>
   * @param matchAllWordFormsRetVal Mandatory boolean parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(14)
  void setMatchAllWordForms(
    boolean matchAllWordFormsRetVal);


  /**
   * <p>
   * Getter method for the COM property "FileName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getFileName();


  /**
   * <p>
   * Setter method for the COM property "FileName"
   * </p>
   * @param fileNameRetVal Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(16)
  void setFileName(
    java.lang.String fileNameRetVal);


  /**
   * <p>
   * Getter method for the COM property "FileType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileType
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.MsoFileType getFileType();


  /**
   * <p>
   * Setter method for the COM property "FileType"
   * </p>
   * @param fileTypeRetVal Mandatory com.exceljava.com4j.office.MsoFileType parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(18)
  void setFileType(
    com.exceljava.com4j.office.MsoFileType fileTypeRetVal);


  /**
   * <p>
   * Getter method for the COM property "LastModified"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoLastModified
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoLastModified getLastModified();


  /**
   * <p>
   * Setter method for the COM property "LastModified"
   * </p>
   * @param lastModifiedRetVal Mandatory com.exceljava.com4j.office.MsoLastModified parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(20)
  void setLastModified(
    com.exceljava.com4j.office.MsoLastModified lastModifiedRetVal);


  /**
   * <p>
   * Getter method for the COM property "TextOrProperty"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getTextOrProperty();


  /**
   * <p>
   * Setter method for the COM property "TextOrProperty"
   * </p>
   * @param textOrProperty Mandatory java.lang.String parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(22)
  void setTextOrProperty(
    java.lang.String textOrProperty);


  /**
   * <p>
   * Getter method for the COM property "LookIn"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(23)
  java.lang.String getLookIn();


  /**
   * <p>
   * Setter method for the COM property "LookIn"
   * </p>
   * @param lookInRetVal Mandatory java.lang.String parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(24)
  void setLookIn(
    java.lang.String lookInRetVal);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoSortBy parameter sortBy is set to 1</li><li>com.exceljava.com4j.office.MsoSortOrder parameter sortOrder is set to 1</li><li>boolean parameter alwaysAccurate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * execute(1, 1, false);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {com.exceljava.com4j.office.MsoSortBy.class, com.exceljava.com4j.office.MsoSortOrder.class, boolean.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VariantBool}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_BOOL}, literal = {"1", "1", "false"})
  @ReturnValue(index=3)
  int execute();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoSortOrder parameter sortOrder is set to 1</li><li>boolean parameter alwaysAccurate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * execute(sortBy, 1, false);
   * </code>
   * </p>
   * @param sortBy Optional parameter. Default value is 1
   * @return  Returns a value of type int
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {com.exceljava.com4j.office.MsoSortOrder.class, boolean.class}, nativeType = {NativeType.Int32, NativeType.VariantBool}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BOOL}, literal = {"1", "false"})
  @ReturnValue(index=3)
  int execute(
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSortBy sortBy);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter alwaysAccurate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * execute(sortBy, sortOrder, false);
   * </code>
   * </p>
   * @param sortBy Optional parameter. Default value is 1
   * @param sortOrder Optional parameter. Default value is 1
   * @return  Returns a value of type int
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  @ReturnValue(index=3)
  int execute(
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSortBy sortBy,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSortOrder sortOrder);

  /**
   * @param sortBy Optional parameter. Default value is 1
   * @param sortOrder Optional parameter. Default value is 1
   * @param alwaysAccurate Optional parameter. Default value is false
   * @return  Returns a value of type int
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(25)
  int execute(
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSortBy sortBy,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSortOrder sortOrder,
    @Optional @DefaultValue("-1") boolean alwaysAccurate);


  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(26)
  void newSearch();


  /**
   * <p>
   * Getter method for the COM property "FoundFiles"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FoundFiles
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.FoundFiles getFoundFiles();


  @VTID(27)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.FoundFiles.class})
  java.lang.String getFoundFiles(
    int index,
    @LCID int lcid);

  /**
   * <p>
   * Getter method for the COM property "PropertyTests"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PropertyTests
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.PropertyTests getPropertyTests();


  @VTID(28)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.PropertyTests.class})
  com.exceljava.com4j.office.PropertyTest getPropertyTests(
    int index,
    @LCID int lcid);

  /**
   * <p>
   * Getter method for the COM property "SearchScopes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SearchScopes
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(29)
  com.exceljava.com4j.office.SearchScopes getSearchScopes();


  @VTID(29)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SearchScopes.class})
  com.exceljava.com4j.office.SearchScope getSearchScopes(
    int index);

  /**
   * <p>
   * Getter method for the COM property "SearchFolders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SearchFolders
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.SearchFolders getSearchFolders();


  @VTID(30)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SearchFolders.class})
  com.exceljava.com4j.office.ScopeFolder getSearchFolders(
    int index);

  /**
   * <p>
   * Getter method for the COM property "FileTypes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FileTypes
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(31)
  com.exceljava.com4j.office.FileTypes getFileTypes();


  @VTID(31)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.FileTypes.class})
  com.exceljava.com4j.office.MsoFileType getFileTypes(
    int index);

  /**
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(32)
  void refreshScopes();


  // Properties:
}
