package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002446B-0001-0000-C000-000000000046}")
public interface IAllowEditRange extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(7)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(8)
  void setTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(9)
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Setter method for the COM property "Range"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(10)
  void setRange(
    com.exceljava.com4j.excel.Range rhs);


  /**
   * @param password Mandatory java.lang.String parameter.
   */

  @VTID(11)
  void changePassword(
    java.lang.String password);


  /**
   */

  @VTID(12)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotect(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void unprotect();

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(13)
  void unprotect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);


  /**
   * <p>
   * Getter method for the COM property "Users"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.UserAccessList
   */

  @VTID(14)
  com.exceljava.com4j.excel.UserAccessList getUsers();


  // Properties:
}
