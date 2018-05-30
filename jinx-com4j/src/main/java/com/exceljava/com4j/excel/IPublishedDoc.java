package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244F8-0001-0000-C000-000000000046}")
public interface IPublishedDoc extends Com4jObject {
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
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getTitle();


  /**
   * <p>
   * Getter method for the COM property "DisclosureScope"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getDisclosureScope();


  /**
   * <p>
   * Getter method for the COM property "Url"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  java.lang.String getUrl();


  /**
   * <p>
   * Getter method for the COM property "PublishedDate"
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @VTID(13)
  java.util.Date getPublishedDate();


  // Properties:
}
