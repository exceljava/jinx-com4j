package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002446D-0001-0000-C000-000000000046}")
public interface IUserAccess extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(7)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "AllowEdit"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(8)
  boolean getAllowEdit();


  /**
   * <p>
   * Setter method for the COM property "AllowEdit"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(9)
  void setAllowEdit(
    boolean rhs);


  /**
   */

  @VTID(10)
  void delete();


  // Properties:
}
