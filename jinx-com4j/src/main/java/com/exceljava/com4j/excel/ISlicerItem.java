package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C9-0001-0000-C000-000000000046}")
public interface ISlicerItem extends Com4jObject {
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
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(9)
  com.exceljava.com4j.excel.SlicerCache getParent();


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getCaption();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSourceName();


  /**
   * <p>
   * Getter method for the COM property "SourceNameStandard"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(13)
  java.lang.String getSourceNameStandard();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(14)
  java.lang.String getValue();


  /**
   * <p>
   * Getter method for the COM property "Selected"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getSelected();


  /**
   * <p>
   * Setter method for the COM property "Selected"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(16)
  void setSelected(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getHasData();


  // Properties:
}
