package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0337-0000-0000-C000-000000000046}")
public interface IFind extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "SearchPath"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(7)
  @DefaultMethod
  java.lang.String getSearchPath();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "SubDir"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  boolean getSubDir();


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getTitle();


  /**
   * <p>
   * Getter method for the COM property "Author"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getAuthor();


  /**
   * <p>
   * Getter method for the COM property "Keywords"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getKeywords();


  /**
   * <p>
   * Getter method for the COM property "Subject"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String getSubject();


  /**
   * <p>
   * Getter method for the COM property "Options"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileFindOptions
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoFileFindOptions getOptions();


  /**
   * <p>
   * Getter method for the COM property "MatchCase"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  boolean getMatchCase();


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(16)
  java.lang.String getText();


  /**
   * <p>
   * Getter method for the COM property "PatternMatch"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  boolean getPatternMatch();


  /**
   * <p>
   * Getter method for the COM property "DateSavedFrom"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDateSavedFrom();


  /**
   * <p>
   * Getter method for the COM property "DateSavedTo"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(19)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDateSavedTo();


  /**
   * <p>
   * Getter method for the COM property "SavedBy"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610743821) //= 0x6002000d. The runtime will prefer the VTID if present
  @VTID(20)
  java.lang.String getSavedBy();


  /**
   * <p>
   * Getter method for the COM property "DateCreatedFrom"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDateCreatedFrom();


  /**
   * <p>
   * Getter method for the COM property "DateCreatedTo"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743823) //= 0x6002000f. The runtime will prefer the VTID if present
  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDateCreatedTo();


  /**
   * <p>
   * Getter method for the COM property "View"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileFindView
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.MsoFileFindView getView();


  /**
   * <p>
   * Getter method for the COM property "SortBy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileFindSortBy
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoFileFindSortBy getSortBy();


  /**
   * <p>
   * Getter method for the COM property "ListBy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFileFindListBy
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.MsoFileFindListBy getListBy();


  /**
   * <p>
   * Getter method for the COM property "SelectedFile"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743827) //= 0x60020013. The runtime will prefer the VTID if present
  @VTID(26)
  int getSelectedFile();


  /**
   * <p>
   * Getter method for the COM property "Results"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IFoundFiles
   */

  @DISPID(1610743828) //= 0x60020014. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.IFoundFiles getResults();


  @VTID(27)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.IFoundFiles.class})
  java.lang.String getResults(
    int index);

  /**
   * @return  Returns a value of type int
   */

  @DISPID(1610743829) //= 0x60020015. The runtime will prefer the VTID if present
  @VTID(28)
  int show();


  /**
   * <p>
   * Setter method for the COM property "SearchPath"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(29)
  @DefaultMethod
  void setSearchPath(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(30)
  void setName(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "SubDir"
   * </p>
   * @param retval Mandatory boolean parameter.
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(31)
  void setSubDir(
    boolean retval);


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(32)
  void setTitle(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "Author"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(33)
  void setAuthor(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "Keywords"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(34)
  void setKeywords(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "Subject"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(35)
  void setSubject(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "Options"
   * </p>
   * @param penmOptions Mandatory com.exceljava.com4j.office.MsoFileFindOptions parameter.
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(36)
  void setOptions(
    com.exceljava.com4j.office.MsoFileFindOptions penmOptions);


  /**
   * <p>
   * Setter method for the COM property "MatchCase"
   * </p>
   * @param retval Mandatory boolean parameter.
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(37)
  void setMatchCase(
    boolean retval);


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743817) //= 0x60020009. The runtime will prefer the VTID if present
  @VTID(38)
  void setText(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "PatternMatch"
   * </p>
   * @param retval Mandatory boolean parameter.
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(39)
  void setPatternMatch(
    boolean retval);


  /**
   * <p>
   * Setter method for the COM property "DateSavedFrom"
   * </p>
   * @param pdatSavedFrom Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743819) //= 0x6002000b. The runtime will prefer the VTID if present
  @VTID(40)
  void setDateSavedFrom(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pdatSavedFrom);


  /**
   * <p>
   * Setter method for the COM property "DateSavedTo"
   * </p>
   * @param pdatSavedTo Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(41)
  void setDateSavedTo(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pdatSavedTo);


  /**
   * <p>
   * Setter method for the COM property "SavedBy"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743821) //= 0x6002000d. The runtime will prefer the VTID if present
  @VTID(42)
  void setSavedBy(
    java.lang.String pbstr);


  /**
   * <p>
   * Setter method for the COM property "DateCreatedFrom"
   * </p>
   * @param pdatCreatedFrom Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(43)
  void setDateCreatedFrom(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pdatCreatedFrom);


  /**
   * <p>
   * Setter method for the COM property "DateCreatedTo"
   * </p>
   * @param pdatCreatedTo Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743823) //= 0x6002000f. The runtime will prefer the VTID if present
  @VTID(44)
  void setDateCreatedTo(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pdatCreatedTo);


  /**
   * <p>
   * Setter method for the COM property "View"
   * </p>
   * @param penmView Mandatory com.exceljava.com4j.office.MsoFileFindView parameter.
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(45)
  void setView(
    com.exceljava.com4j.office.MsoFileFindView penmView);


  /**
   * <p>
   * Setter method for the COM property "SortBy"
   * </p>
   * @param penmSortBy Mandatory com.exceljava.com4j.office.MsoFileFindSortBy parameter.
   */

  @DISPID(1610743825) //= 0x60020011. The runtime will prefer the VTID if present
  @VTID(46)
  void setSortBy(
    com.exceljava.com4j.office.MsoFileFindSortBy penmSortBy);


  /**
   * <p>
   * Setter method for the COM property "ListBy"
   * </p>
   * @param penmListBy Mandatory com.exceljava.com4j.office.MsoFileFindListBy parameter.
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(47)
  void setListBy(
    com.exceljava.com4j.office.MsoFileFindListBy penmListBy);


  /**
   * <p>
   * Setter method for the COM property "SelectedFile"
   * </p>
   * @param pintSelectedFile Mandatory int parameter.
   */

  @DISPID(1610743827) //= 0x60020013. The runtime will prefer the VTID if present
  @VTID(48)
  void setSelectedFile(
    int pintSelectedFile);


  /**
   */

  @DISPID(1610743850) //= 0x6002002a. The runtime will prefer the VTID if present
  @VTID(49)
  void execute();


  /**
   * @param bstrQueryName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743851) //= 0x6002002b. The runtime will prefer the VTID if present
  @VTID(50)
  void load(
    java.lang.String bstrQueryName);


  /**
   * @param bstrQueryName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743852) //= 0x6002002c. The runtime will prefer the VTID if present
  @VTID(51)
  void save(
    java.lang.String bstrQueryName);


  /**
   * @param bstrQueryName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743853) //= 0x6002002d. The runtime will prefer the VTID if present
  @VTID(52)
  void delete(
    java.lang.String bstrQueryName);


  /**
   * <p>
   * Getter method for the COM property "FileType"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743854) //= 0x6002002e. The runtime will prefer the VTID if present
  @VTID(53)
  int getFileType();


  /**
   * <p>
   * Setter method for the COM property "FileType"
   * </p>
   * @param plFileType Mandatory int parameter.
   */

  @DISPID(1610743854) //= 0x6002002e. The runtime will prefer the VTID if present
  @VTID(54)
  void setFileType(
    int plFileType);


  // Properties:
}
