
<%
ignorehack = True
allowanything = False
%>
<!-- #include file="a_server-checks.asp"-->

<!-- #include file="include_meta.asp" -->

</head>
<body class="webstep2">

  <%
            Session("RCM273_Currency")="EUR"
           Session("RCM273_CompanyCurrency")="&#8364;"
           Session("RCM273_CompanyColour")="#0080C0" '---the booking form color is blue, web designer can change it here
           CompanyCode="PtWestCoastCampers273"
           RCMURL="https://secure2.rentalcarmanager.com"

           Dim webConn
           Set webConn = Server.CreateObject("ADODB.Connection")
            DatabaseFile="PtWestCoastCampers273"
            webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"

         Set s_sy=webconn.Execute("SELECT * FROM SystemTable WHERE Code='AvgRate'  ")
         Session("RCM273_useEverageRate")="No" '---flag to change the code between use  everage Rate or seasonal rate
        If not s_sy.EOF then
              Session("RCM273_useEverageRate") = s_sy("syValue")
         END IF
        s_sy.CLOSE
         SET s_sy=NOTHING
FUNCTION tidyup(thisString)

      thisString=LTrim(RTrim(Replace(thisString, "'" , "")))
      thisString=Replace(thisString, "”" , "")
      thisString=Replace(thisString, "<" , "")
      thisString=Replace(thisString, "http" , "")
      thisString=Replace(thisString, "://" , "")
      thisString=Replace(thisString, "=" , "")
       thisString=Replace(thisString, "?" , " ")
      thisString=Replace(thisString, "/" , "")
      thisString=Replace(thisString, "+" , "")
      thisString=Replace(thisString, "%" , "")
      tidyup=Replace(thisString, ">" , "")
END FUNCTION
SUB SelectPickupLocation
      Response.Write("<select name='PickupLocationID' class='form-control input-sm'>")

               Set l_s = webConn.Execute("select * FROM Location WHERE PickupAvailable=1 order by location")
        While Not l_s.EOF
                if CStr(l_s("ID"))=Session("RCM273_PickupLocationID")  then
                           Response.Write("<option value='"&l_s("ID")&"' selected=""selected"">"&l_s("Location")&" &nbsp;</option>")
               else
                  Response.Write("<option value='"&l_s("ID")&"' >"&l_s("Location")&" &nbsp;</option>")
               end if
        l_s.MoveNext
         Wend
         l_s.Close
         Set l_s=nothing
        Response.Write("</select>")
END SUB

SUB SelectDropoffLocation
   Set l_s = webConn.Execute("select * FROM Location WHERE DropoffAvailable=1  order by location")
            Response.Write("<select name='DropoffLocationID' class='form-control input-sm'  >")

                   While Not l_s.EOF
                     if CStr(l_s("ID"))=Session("RCM273_DropoffLocationID")  then
                              Response.Write("<option value='"&l_s("ID")&"' selected=""selected"">"&l_s("Location")&" &nbsp;</option>")
                     else
                              Response.Write("<option value='"&l_s("ID")&"' >"&l_s("Location")&" &nbsp;</option>")
                     end if
           l_s.MoveNext
            Wend
            l_s.Close
      Set l_s=nothing
         Response.Write("</select>")
END SUB


SUB TheQutoationForm
          if Session("RCM273_RequestPickDate")="" then
                Session("RCM273_RequestPickDate")=Day(Date+5)&"/"&MonthName(Month(Date+5))&"/"&Year(Date+5)
                Session("RCM273_RequestDropDate")=Day(Date+15)&"/"&MonthName(Month(Date+15))&"/"&Year(Date+15)
                Session("RCM273_RequestPickTime")="11:00"
                 Session("RCM273_RequestDropTime")="10:00"
         end if

         RequestPickDate=Session("RCM273_RequestPickDate")
         RequestDropDate= Session("RCM273_RequestDropDate")


         defaultPickupTime= Session("RCM273_RequestPickTime")
         defaultDrpoffTime= Session("RCM273_RequestDropTime")
         defaultPickupDate = Session("RCM273_RequestPickDate")

         defaultDropoffDate = Session("RCM273_RequestDropDate")



  %>

<!-- #include file="include_header.asp" -->

     
<div class="container">
<h4 class="smallm_title centered bigger"><span>Change Booking enquiry</span></h4>

<form method="post" name="theform" action="webstep2.asp" id="theform" class="form-inline" >
<div class="jumbotron">
<div class="row">
<div class="form-group">
    
      <div class="col-xs-4">
      <label class="control-label">Pickup Location:</label>
                        <%    SelectPickupLocation  %>
                      </div>

                        <div class="col-xs-5">
      <label class="control-label" for="PickupDay">Pickup Date:</label>
                              <select name="PickupDay" id="PickupDay" class="form-control input-sm">
                                <%     for pd=1 to 9
                     zero = "0"
                     if  pd=Day(DefaultPickupDate)   then
                                  Response.Write("<option value='")
                              Response.Write(""& zero & pd &"' selected=""selected"">"&pd&"</option>")
                           else
                                    Response.Write("<option value='"& zero & pd &"' >"&pd&"</option>")
                           end if
                     zero = ""
                  next
                     for pd=10 to 31
                           if  pd=Day(DefaultPickupDate) then
                                 Response.Write("<option value='")
                                 Response.Write(""& pd &"' selected=""selected"">"&pd&"</option>")
                           else
                                 Response.Write("<option value='"& pd &"' >"&pd&"</option>")
                           end if
                           zero = ""
                     next
                                %>
                              </select>


                            <label class="sr-only" for="PickupMonth">Month</label>
                              <select name="PickupMonth" id="PickupMonth" class="form-control input-sm">
                                <% for pm = 1 to 9
               theMonth=Left(MonthName(pm),3)
               monthValue="0"&pm
                           if   pm=Month(defaultPickupDate)    then
                                 Response.Write("<option value='"&monthValue&"' selected=""selected"">"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&monthValue&"' >"&theMonth&"</option>")
                                end if
                  next
                  for pm = 10 to 12
               theMonth=Left(MonthName(pm),3)
                              if   pm=Month(defaultPickupDate)    then
                                 Response.Write("<option value='"&pm&"' selected=""selected"">"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&pm&"' >"&theMonth&"</option>")
                                end if

             next %>
                              </select>

                            <label class="sr-only" for="PickupYear">Year</label>
                              <select name="PickupYear" id="PickupYear" class="form-control input-sm">
                                <%   for theYear=(Year(Date)-1) to (Year(Date)+2)
                if   theYear=Year(defaultPickupDate)  then
                               Response.Write("<option value='"&theYear&"' selected=""selected"">"&theYear&"</option>")
                     else
                           Response.Write("<option value='"&theYear&"'>"&theYear&"</option>")
                     end if
                next
                                %>
                              </select>
                            

                              <input type="hidden" value="<% if len(month(defaultPickupDate))<2 then Response.Write("0"&month(defaultPickupDate)) else Response.Write(month(defaultPickupDate)) end if%><%="/" & day(defaultPickupDate) & "/" & year(defaultPickupDate)%>" id="txtStartDate" size="12" />

                            </div>

       <div class="col-xs-3">
           <label class="control-label" for="pickupTime">Pickup Time:</label>
                        <select name="pickupTime" id="pickupTime" class="form-control input-sm">
                          <option value='00:00'>midnight</option>
                          <option value='00:30'>00:30 AM</option>
                          <%
                      for i=1 to 23
                          for j=1 to 2
                              if j=1  then
                                 min="00"
                               else
                                 min="30"
                               end if
                               if  i<12 then
                                 test=i+1&":"&min&" PM"
                                 if i<10 then
                                      theTime="0"&i&":"&min
                                 else
                                          theTime=i&":"&min
                                 end if
                               else
                                 test=i&":"&min&" PM"
                                 theTime=i&":"&min
                               end if

                               'if   Hour(Time)= (i-1) AND  min="00" then
                                if   theTime = defaultPickupTime then
                                 Response.Write("<option value='"&theTime&"' selected=""selected"">"&test&"</option>")
                               else
                                 Response.Write("<option value='"&theTime&"' >"&test&"</option>")

                               end if
                           next
                      next
                          %>
                        </select>
                     </div>

</div><!-- form-group -->
</div>


<div class="row">                   
<div class="form-group"> 
      <div class="col-xs-4">
      <label class="control-label">Return Location:</label>

                        <%   SelectDropoffLocation    '--Subs         %>
                            </div>
      
      <div class="col-xs-5">
      <label class="control-label" for="DropoffDay">Return Date:</label>

                              <select name="DropoffDay" id="DropoffDay"  class="form-control input-sm">
                                <%  for pd=1 to 9
                     zero = "0"
                     if  pd=Day(DefaultDropoffDate)   then
                           Response.Write("<option value='")
                              Response.Write(""& zero & pd &"' selected=""selected"">"&pd&"</option>")
                           else
                                    Response.Write("<option value='"& zero & pd &"' >"&pd&"</option>")
                           end if
                     zero = ""
                  next
                     for pd=10 to 31
                           if  pd=Day(DefaultDropoffDate) then
                        Response.Write("<option value='")
                        Response.Write(""& pd &"' selected=""selected"">"&pd&"</option>")
                     else
                              Response.Write("<option value='"& pd &"' >"&pd&"</option>")
                           end if
               zero = ""
            next
                                %>
                              </select>

                            <label class="sr-only" for="DropoffMonth">Month</label>
                              <select name="DropoffMonth" id="DropoffMonth"  class="form-control input-sm">
                                <%

          for pm = 1 to 9
               theMonth=Left(MonthName(pm),3)
               monthValue="0"&pm
                          if  pm=Month(defaultDropoffDate)   then
                        Response.Write("<option value='"&monthValue&"' selected=""selected"">"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&monthValue&"' >"&theMonth&"</option>")
                                end if
                  next
                  for pm = 10 to 12
               theMonth=Left(MonthName(pm),3)
                         if  pm=Month(defaultDropoffDate)   then
                        Response.Write("<option value='"&pm&"' selected=""selected"">"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&pm&"' >"&theMonth&"</option>")
                                end if

             next
                                %>
                              </select>

                            <label class="sr-only" for="DropoffYear">Year</label>
                              <select name="DropoffYear" id="DropoffYear"  class="form-control input-sm">
                                <%    for theYear=(Year(Date)-1) to (Year(Date)+5)
                     if   theYear=Year(defaultDropoffDate)  then
                               Response.Write("<option value='"&theYear&"' selected=""selected"">"&theYear&"</option>")
                     else
                           Response.Write("<option value='"&theYear&"'>"&theYear&"</option>")
                     end if
                  next    %>
                              </select>
                            


                              <input type="hidden" value="<% if len(month(defaultDropoffDate))<2 then Response.Write("0"&month(defaultDropoffDate)) else Response.Write(month(defaultDropoffDate)) end if%>/<%if len(day(defaultDropoffDate)) < 2 then Response.Write("0"&day(defaultDropoffDate)) else Response.Write(day(defaultDropoffDate)) end if%>/<%=year(defaultDropoffDate)%>" id="txtEndDate" size="12" />
                       
                    </div>




                     <div class="col-xs-3">
                        <label class="control-label" for="DropoffTime" >Return Time:</label>
                        <select name="DropoffTime" id="DropoffTime" class="form-control input-sm" >
                          <option value='00:00'>midnight</option>
                          <option value='00:30'>00:30 AM</option>
                          <%
                     for i=1 to 23
                           for j=1 to 2
                              if j=1  then
                                 min="00"
                               else
                                 min="30"
                               end if
                               if  i<12 then
                                 test=i+2&":"&min&" PM"
                                 if i<10 then
                                      theTime="0"&i&":"&min
                                 else
                                          theTime=i&":"&min
                                 end if
                               else
                                 test=i&":"&min&" PM"
                                 theTime=i&":"&min
                               end if

                               'if   Hour(Time)= (i-1) AND  min="00"  then
                               if   theTime = defaultDrpoffTime then
                                 Response.Write("<option value='"&theTime&"' selected=""selected"">"&test&"</option>")
                               else
                                 Response.Write("<option value='"&theTime&"' >"&test&"</option>")

                               end if
                           next
                      next
                          %>
                       


                        <%   Set s_c = webConn.Execute("select count(*) as SizeCount FROM CategoryType  WHERE RentingType=1  ")

               if s_c("SizeCount") >1 then
                       Set l_s = webConn.Execute("SELECT * FROM CategoryType WHERE RentingType=1  order by Ordering") %>

                      <label class="control-label" for="CategoryTypeID" >Vehicle Type:</label>
                        
                        <select name="CategoryTypeID" class="form-control input-sm" >

                       <option value="0" >Choose Vehcile Type</option>
                      <%  While Not l_s.EOF
                              if CStr(l_s("ID"))=Session("RCM273_CategoryTypeID")  then
                                    Response.Write("<option value='"&l_s("ID")&"' selected=""selected"">"&l_s("CategoryType")&" &nbsp;</option>")
                              else
                                    Response.Write("<option value='"&l_s("ID")&"' >"&l_s("CategoryType")&" &nbsp;</option>")
                              end if
                        l_s.MoveNext
                        Wend
                        l_s.Close
                        Set l_s=nothing
                     Response.Write("</select>")
               else
                     s_c.close
                     set s_c=nothing
                     Set s_c = webConn.Execute("SELECT ID  FROM CategoryType WHERE RentingType=1 ")
                     Response.Write("<input name=""CategoryTypeID"" type=""hidden""  value='"&s_c("ID")&"'  /></td>")
                     Response.Write("<td align=""left"" >")
               end if
               s_c.close
               set s_c=nothing
                        %>
                      

</div>
</div><!-- form-group -->

       


</div><!-- row -->

<div class="row">
<div class="form-group"> 


<div class="col-xs-4">


     <label class="control-label" for="YoungestDriver">Youngest Driver:</label>
                        <%
      Response.Write("<select name=""driverage"" id=""YoungestDriver"" class=""form-control input-sm""><option value=""0"" selected=""selected"">Please select</option>")
        SET s_c=webConn.Execute("SELECT Count(*) as Records from DriverAge")
        NoRecords=0
        if ISNULL(s_c("Records"))<>True then
         NoRecords=s_c("Records")
         end if
         s_c.close
         SET s_c=nothing

        SET s_m=webConn.Execute("SELECT * from DriverAge  order by DriverAge")
        k=0
        While Not s_m.EOF
               k=k+1
                DriverAge=s_m("DriverAge")&" Years"
               if k=NoRecords then
                  DriverAge=s_m("DriverAge")&" + Years"
               end if
                  if CStr(s_m("DriverAge"))=Session("RCM273_driverage")  then
                           Response.Write("<option value='"&s_m("DriverAge")&"' selected=""selected"">"&DriverAge&"</option>")
                  elseif Session("RCM273_driverage")="" and s_m("DefaultAge")="True" then
                           Response.Write("<option value='"&s_m("DriverAge")&"' selected=""selected"">"&DriverAge&"</option>")
                  else
                           Response.Write("<option value='"&s_m("DriverAge")&"' >"&DriverAge&"</option>")
                  end if
        s_m.MoveNext
        Wend
        s_m.Close
        Set s_m=nothing     %>
                        </select>

</div>

<div class="col-xs-5">&nbsp;
</div>

<div class="col-xs-3">


                          <input name="submit" type="submit" class="rcmbutton btn btn-info btn-block btn-lg" value="GET A QUOTE" />
  </div>
  </div>
</div>
      </div><!-- jumbotron -->
                  </form>

              <%



   END SUB
 SUB  GetEachSeasonRateStructureID
          '-----check system set up if Calculate Seasonal Rates using total rental days (long hire rate)
         TotalDays= Session("RCM273_TotalRentalDays")

         longHireRate="No"
        Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='LONGR'  ")
         If not RG.EOF then
            longHireRate= RG("syValue")'--TOTAL booking days
         END IF
         RG.CLOSE
         SET RG=NOTHING
         DaysForRate=NoOfDaysEachSeason
          '========new code========
          if  theExtraHour<>0 and actureNoOfDaysEachSeason >=1 and HourRate<>0 then      '---if 1.4 days should use a 1 day rate  not 2 day rate., need add condition "theExtraHour=0" here for add 1 day up
                DaysForRate=Fix(actureNoOfDaysEachSeason)                                '--HourRate<>0, if hourRate has set up, otherwise charge 1.4 days as 2 day rate
         end if
         '========end new code========

        if longHireRate="Yes" then    
               DaysForRate= TotalDays '--TOTAL booking days
                '========new code========
                if  theExtraHour<>0 and actureTotalDays >=1 and HourRate<>0 then  '---if 1.4 days should use a 1 day rate but not 2 day rate, need add condition "theExtraHour=0" here for add 1 day up
                  DaysForRate=Fix(actureTotalDays)
                  end if
               '========end new code======== 
        end if
       'Response.Write("<br>test"&TotalDays&"</br>")
       

        '--for a 3.5 days booking should use 3 days rate(the  higher rate) not the 4 days rate
        if  DaysForRate > Fix(DaysForRate) then
         DaysForRate=Fix(DaysForRate)
         end if

        Session("RCM273_DaysForRate")=DaysForRate
          RateName="Rate"
          Rate=0
          RateStructureID=1
          Session("RCM273_RateStructureID"&i&"")=1
         '----------get rateStructureID ------------------
         '---weekend  booking

         '---only bookings <=3 days and between 12:00 Friday to 12:00 Monday (AU Coutesy Cars)
   '---check systemtable if use weekend rate
       Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='WKEND'  ")
       UseWeekendRate="No"
       If not RG.EOF then
       UseWeekendRate=RG("syValue")
       end if
       RG.CLOSE
       SET RG=NOTHING


       WeekendRate="No"
   IF UseWeekendRate="Yes" THEN
      if  DaysForRate<=3 and VBFriday=Weekday(Session("RCM273_RequestPickDate")) then
          WeekendStard=Session("RCM273_RequestPickDate")&" 12:00"
          WeekendEnd=(CDate(Session("RCM273_RequestPickDate"))+3)&" 12:00"
          if   CDate(Session("RCM273_RequestPickDateTime"))=>CDate(WeekendStard) and CDate(Session("RCM273_RequestDropDateTime"))<=CDate(WeekendEnd) then
                  WeekendRate="Yes"
                  'Response.Write(WeekendStard)
                   'Response.Write("---5Weekend----")
                   'Response.Write(WeekendEnd)
          end if
      end if

      if  DaysForRate<=2 and VBSaturday=Weekday(Session("RCM273_RequestPickDate")) then
               WeekendStard=(CDate(Session("RCM273_RequestPickDate"))-1)&" 12:00"
               WeekendEnd=(CDate(Session("RCM273_RequestPickDate"))+2)&" 12:00"
               if   CDate(Session("RCM273_RequestDropDateTime"))<=CDate(WeekendEnd) then
                   WeekendRate="Yes"
                  'Response.Write(WeekendStard)
                   'Response.Write("---6Weekend----")
                   'Response.Write(WeekendEnd)
               end if
      end if
          '---added Dec 2009,   for 1 day booking pickup sunday 
          if  DaysForRate<=1 and  VBSunday=Weekday(Session("RCM273_RequestPickDate")) then
                   WeekendStard=(CDate(Session("RCM273_RequestPickDate"))-1)&" 12:00"
                  WeekendEnd=(CDate(Session("RCM273_RequestPickDate"))+1)&" 12:00"
                  if   CDate(Session("RCM273_RequestDropDateTime"))<=CDate(WeekendEnd) then
                      WeekendRate="Yes"
                      'Response.Write(WeekendStard)
                        'Response.Write("---7Weekend----")
                         'Response.Write(WeekendEnd)
                  end if
            end if
   END IF   '---end UseWeekendRate="Yes"

       if WeekendRate="Yes"   then
                 'Session("RCM273_RateStructureID"&i&"")=1
                 RateStructureID=1
                 Session("RCM273_RateStructureID"&i&"")=1
                  RateName="Weekend"
                  FromDay=2
                 ToDay=2
                'Response.Write("<br>1. Weekend 2 - 2")
      else
               '--get RateStructure 1st--------

               SQL= "SELECT Top 1 (FromDay), ID, RateName FROM CarRateStructure where ID<>1 "
               SQL=SQL&" AND FromDay<="&Cint(DaysForRate)&" order by FromDay desc "
               Set s_rs=webConn.Execute(SQL)
               if NOT s_rs.EOF   then  '---  not weekend (more then 2 days)
                     'Response.Write("<br>"&s_rs("ID")&". "&s_rs("RateName")&":&nbsp;"&s_rs("FromDay")&" - "&s_rs("ToDay")&"&nbsp;")
                     RateName=s_rs("RateName")
                     RateStructureID=s_rs("ID")
                     Session("RCM273_RateStructureID"&i&"")=s_rs("ID")
               end if
               s_rs.Close
               SET s_rs=nothing
       end if
       'Response.Write(RateStructureID)
       'keep rateStructureID here

END SUB

 SUB  findTheRate2  '---Sample Web Quotation1       '---get the rate Discount-----------



         Response.Write("<h4 class=""smallm_title centered bigger""><a name='book'></a><span>Book your campervan</span></h4>")

        '---allow muti-records for Unavailable booking perioed, only select distinc record 1st
       SQL=" SELECT  Distinct  CategoryID, ordering,CategoryType,CarSize.*  FROM WebLocationCategory,CarSize ,CategoryType "
       SQL=SQL&" WHERE (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' "
       SQL=SQL&" AND CategoryID=CarSize.ID and CategoryTypeID=CategoryType.ID "
       SQL=SQL&" AND CategoryTypeID='"&CINT(Session("RCM273_CategoryTypeID"))&"'  "
       SQL=SQL&" AND WebAvaliable=1 "
       SQL=SQL&" ORDER By ordering,CategoryType, Orders,Size "
       '---mar 2012, new category available table with status (s_WebCategoryAvailability.asp), seach categoryStatus for code changes for other steps
      '--=====check step3====== findTheRate2, code changed, select the category from carSize Table only, not WebLocationCategory --
      SQL=" SELECT  Distinct  CategoryID, ordering,CategoryType,CarSize.*  FROM WebCategoryAvailability,CarSize ,CategoryType "
       SQL=SQL&" WHERE (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' "
       SQL=SQL&" AND CategoryID=CarSize.ID and CategoryTypeID=CategoryType.ID "
       SQL=SQL&" AND CategoryTypeID='"&CINT(Session("RCM273_CategoryTypeID"))&"'  "
       SQL=SQL&" AND WebAvaliable=1 "
      'SQL=SQL&" AND WebCategoryAvailability.CompanyID=1 "
       SQL=SQL&" ORDER By ordering,CategoryType, Orders,Size "
      'Response.Write SQL
      Set s_cs=webConn.Execute(SQL)

      j=0

      DO WHILE NOT s_cs.EOF    'for each car Category
         categoryStatus=""
         '---need default to not available, if there is only unavailable record, with out available record,
         categoryAvaliable="No"
         RelocationMinBookingDay=0
          Unavailable="True"
         ' ++++++++++++
            '---Jan 2013, added back the min day check from the relocation table, new flag added
             categoryAvaliable4Relocation="Yes"
          '--10/Mar/09, nwe code, check MinBookingDay for each category  ---
                 SQL = "SELECT  top 1 MinBookingDay,WebAvaliable FROM WebRelocationFees where CarSizeID="&s_cs("ID")&" "
                SQL=SQL&" and (PickupLocationID) ='"&CINT(Session("RCM273_PickupLocationID"))&"' and (DropoffLocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"' "
                SQL=SQL&" AND ((PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"') "
                 SQL=SQL&" or  PickupDateFrom='2100-01-01') "
                SQL=SQL&" order by PickupDateFrom "
                Set s_re=webConn.Execute(SQL)
                if Not s_re.EOF then
                     if  s_re("WebAvaliable") =False then
                          categoryAvaliable="No"  '--with message category unavailiable
                          categoryAvaliable4Relocation="No"
                           'Response.Write(categoryAvaliable)
                     else
                           if s_re("MinBookingDay") > TotalRentalDays then
                                 RelocationMinBookingDay=s_re("MinBookingDay") '--with message minibooking day
                                 categoryAvaliable="No"
                                 categoryAvaliable4Relocation="No"
                              end if
                     end if


               else
                        s_re.close
                       set s_re=nothing
                       SQL = "SELECT  top 1 MinBookingDay,WebAvaliable FROM WebRelocationFees where CarSizeID=0 "
                        SQL=SQL&" and (PickupLocationID) ='"&CINT(Session("RCM273_PickupLocationID"))&"' and (DropoffLocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"' "
                        SQL=SQL&" AND ((PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"') "
                        SQL=SQL&" or  PickupDateFrom='2100-01-01') "
                        SQL=SQL&" order by PickupDateFrom "
                         Set s_re=webConn.Execute(SQL)
                          if Not s_re.EOF then
                              if  s_re("WebAvaliable") =False then
                                    categoryAvaliable="No"
                                    categoryAvaliable4Relocation="No"
                                   'Response.Write(categoryAvaliable)
                              else
                                    if s_re("MinBookingDay") > TotalRentalDays then
                                       RelocationMinBookingDay=s_re("MinBookingDay")
                                       categoryAvaliable="No"
                                       categoryAvaliable4Relocation="No"
                                    end if
                              end if
                         end if


               end if
               'Response.Write(SQL)
               s_re.close

            '--for each location category check all muti unavailabe record  ....
                  SQL=" SELECT UnavaliableFrom,UnavaliableTo,MinBookingDay  FROM WebLocationCategory "
                  SQL=SQL&" WHERE CategoryID="&s_cs("CategoryID")&" and (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
                  ' SQL=SQL&" AND WebLocationCategory.CompanyID=12 "
                  '---Mar 2012, new changes
                   SQL=" SELECT top 1 *   FROM WebCategoryAvailability "
                  SQL=SQL&" WHERE CategoryID="&s_cs("CategoryID")&" and (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
                  SQL=SQL&" AND (DropoffLocationID=0 or DropoffLocationID = '"&CINT(Session("RCM273_DropoffLocationID"))&"' ) "
                ' SQL=SQL&" and (DateFrom<='"&RequestPickDate&"' and DateTo>='"&RequestPickDate&"') "
                  '---Oct 2012, new code,  will not available if  any renting date is in the unavailable date range, not just the pickup dates in the date range
                  SQL=SQL&" and ( ('"&RequestPickDate&"'>=Datefrom and '"&RequestPickDate&"' <=DateTo) or ('"&RequestDropDate&"'>=DateFrom and '"&RequestDropDate&"'<=DateTo) or ('"&RequestPickDate&"'<=DateFrom and '"&RequestDropDate&"'>=DateTo) ) "
                ' SQL=SQL&" AND WebLocationCategory.CompanyID=12 "
                   '---Aug 2014,  changed the order by, get the unallvailabe 1st
                      'SQL=SQL&" order by DropoffLocationID DESC, StatusID DESC "
                           SQL=SQL&" order by StatusID DESC, DropoffLocationID DESC "
                            ' Response.Write SQL
                Set s_min=webConn.Execute(SQL)
                Unavailable="False"
                MinBookingDay=0
                  ' StatusID=1 Available for booking (free sale)
                  ' StatusID=2 Might be available (place a request)
                  ' StatusID=3 Not available
                if NOT s_min.EOF  then
                       categoryStatus=s_min("StatusID")
                       categoryAvaliable="Yes" '--if find a record then set to available
                        Unavailable="False"
                       if s_min("MinBookingDay") > TotalRentalDays then
                                 MinBookingDay=s_min("MinBookingDay") '--with message minibooking day
                                 categoryAvaliable="No" '--with message minibooking day
                       end if
                       if s_min("StatusID")=3 then
                                    Unavailable="True" '---booked out message
                       end if
                       s_min.Close
                       set s_min=nothing
                else
                        s_min.Close
                       set s_min=nothing
                       SQL=" SELECT top 1 *   FROM WebCategoryAvailability "
                        SQL=SQL&" WHERE CategoryID="&s_cs("CategoryID")&" and (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
                         SQL=SQL&" and Year(DateFrom)=2100 "
                           SQL=SQL&" AND (DropoffLocationID=0 or DropoffLocationID = '"&CINT(Session("RCM273_DropoffLocationID"))&"' ) "
                           '---Aug 2014,  changed the order by, get the unallvailabe 1st
                         'SQL=SQL&" order by DropoffLocationID DESC, StatusID DESC "
                           SQL=SQL&" order by StatusID DESC, DropoffLocationID DESC "
                         Set s_min=webConn.Execute(SQL)
                        if NOT s_min.EOF  then
                              'Response.Write s_min("MinBookingDay")
                              categoryStatus=s_min("StatusID")
                              categoryAvaliable="Yes"
                              Unavailable="False"
                              if s_min("MinBookingDay") > TotalRentalDays then
                                    MinBookingDay=s_min("MinBookingDay") '--with message minibooking day
                                    categoryAvaliable="No" '--with message minibooking day
                              end if
                              if s_min("StatusID")=3 then
                                    Unavailable="True" '---booked out message
                              end if

                        end if
                        s_min.Close
                        set s_min=nothing
               end if
                'Response.Write SQL
                 'Response.Write "<br>categoryStatus="
               'Response.Write categoryStatus




      if  categoryAvaliable="Yes" and categoryAvaliable4Relocation="Yes" then


           '----- SUB subGETHourlyRate
            Less1DayHireHourRate=0
            TotalDays4HourlyRate=Round((RequestDropDateTime - RequestPickDateTime),2) '--do not include Grace period
            'Response.Write TotalDays4HourlyRate
            if TotalDays4HourlyRate<1 then
                  CategoryID=s_cs("CategoryID")
                  NumberOfHours= Round((RequestDropDateTime - RequestPickDateTime)*1440/60,2)
                  if NumberOfHours>Fix(NumberOfHours) then
                     NumberOfHours=Fix(NumberOfHours) + 1
                  end if
                  SQL= "SELECT Top 1 (ToHour), ID, RateName,NumberDays FROM CarRateStructureHour   "
                  SQL=SQL&" where FromHour<"&Cint(NumberOfHours)&"  and ToHour>="&Cint(NumberOfHours)&" order by Fromhour desc "
                  'Response.Write SQL
                  Set s_rs=webConn.Execute(SQL)
                  if NOT s_rs.EOF   then
                                 RateName=s_rs("RateName")
                                 RateStructureID=s_rs("ID")
                                 NumberDays=s_rs("NumberDays") '---converted NumberDays for save rate and reports
                                 SET s_m=webConn.Execute("SELECT * from Season where (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' and notActive=0 and (Season='Default' or  (EndDate>='"&Session("RCM273_RequestPickDate")&"' and StartDate<='"&Session("RCM273_RequestDropDate")&"')) order by StartDate DESC ")
                                 if  NOT s_m.EOF then
                                       Season=s_m("Season")
                                       SeasonID = s_m("ID")
                                       SQL="SELECT * FROM QCarRateDetailsHour  where CarRateStructureID="&s_rs("ID")&" and  SeasonID="&s_m("ID")&"   and CarSizeID="&Cint(CategoryID)&"  "
                                       'Response.Write SQL
                                       Set s_r=webConn.Execute(SQL)

                                       IF NOT s_r.EOF THEN
                                          if ISNull(s_r("Rate"))<>True then
                                             Less1DayHireHourRate=s_r("Rate")
                                             '---convert hour rate to stadard rate by using the NumberDays saved in CarRateStructureHour
                                             '---make it easy to save the rates and for reports
                                             StandardRate1=s_r("Rate")/s_rs("NumberDays")
                                          end if
                                       else
                                               '--dec 2012, changed Season='Default' to  SeasonID="&s_m("ID")&", SeasonID is location based

                                             SQL="SELECT * FROM QCarRateDetailsHour  where CarRateStructureID="&s_rs("ID")&" and  SeasonID="&s_m("ID")&"   and CarSizeID="&Cint(CategoryID)&"  "
                                             'Response.Write SQL
                                             Set s_s=webConn.Execute(SQL)

                                             IF NOT s_s.EOF THEN
                                                   if s_s("Rate")<>"" then
                                                   Less1DayHireHourRate=s_s("Rate")
                                                   '---convert hour rate to stadard rate by using the NumberDays saved in CarRateStructureHour
                                                   '---make it easy to save the rates and for reports
                                                   StandardRate1=s_s("Rate")/s_rs("NumberDays")
                                                end if
                                             END IF
                                             s_s.close
                                             set s_s=nothing

                                       END IF
                                       s_r.close
                                       set s_r=nothing

                                 end if
                                 s_m.close
                                 SET s_m=nothing
                     end if
                     s_rs.Close
                     SET s_rs=nothing

            end if
          


       '  if Less1DayHireHourRate=0 then'--------------






             '---get discount Rate
                '--- check if the rental period in the Discount Date range
                '--- then check if there is a category discount rate
                '--- if not then use the location discount rate
               DiscountRate=0
            DiscountID=0
            DiscountType="p"
            '---Rita, Jul 2012, new code for discount, same as step3
            BookingDate=Year(Date)&"-"&Month(Date)&"-"&Day(Date)
            SQL=" SELECT TOP 1 LocationID,DropoffLocationID, CarSizeID,  DateRange from Discount left join CampaignCode on CampaignCodeID=CampaignCode.ID  "
            SQL=SQL&" WHERE WebItems=1   "
            if Session("RCM273_PromoCode")<>"" then
            SQL=SQL&" AND CampaignCode='"&Left(tidyup(Session("RCM273_PromoCode")),30)&"'  "
            else
            SQL=SQL&" AND CampaignCodeID=0  "
            end if
            SQL=SQL&" AND (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or LocationID=0 )  "
            SQL=SQL&" AND (DropoffLocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"' or DropoffLocationID=0 )  "
            SQL=SQL&" AND (CarSizeID="&s_cs("ID")&" or CarSizeID=0)  "
            SQL=SQL&" AND (DateRange=0 or ( DateRange=1 "
            SQL=SQL&" AND DateFrom<='"&RequestPickDate&"' and DateTo>='"&RequestDropDate&"' "
            SQL=SQL&" AND ( (BookingDateFrom<='"&BookingDate&"' and BookingDateTo>='"&BookingDate&"') "
            SQL=SQL&" or (BookingDateFrom is Null or YEAR(BookingdateFrom)=2100) )"
             SQL=SQL&" ))  "
            SQL=SQL&" Group By LocationID,DropoffLocationID, CarSizeID,  DateRange "
            SQL=SQL&" order by LocationID desc, DropoffLocationID Desc, CarSizeID desc,  DateRange desc "
            'TotalResponse.Write(SQL)
            Set s_disDate=webConn.Execute(SQL)
            if  NOT s_disDate.EOF then    '---rental period in the discount date range and specific category apply
                     SQL="SELECT Discount.*,CampaignCode  from Discount left join CampaignCode on CampaignCodeID=CampaignCode.ID "
                     SQL=SQL&"  WHERE  WebItems=1  "
                     SQL=SQL&" AND (LocationID)='"&s_disDate("LocationID")&"' "
                     SQL=SQL&" AND (DropoffLocationID)='"&s_disDate("DropoffLocationID")&"' "
                     SQL=SQL&" AND CarSizeID="&s_disDate("CarSizeID")&" "
                     if Session("RCM273_PromoCode")<>"" then
                     SQL=SQL&" AND CampaignCode='"&Left(Session("RCM273_PromoCode"),30)&"'  "
                     else
                     SQL=SQL&" AND CampaignCodeID=0  "
                     end if
                     if s_disDate("DateRange")=0 then
                           SQL=SQL&" and DateRange=0 "
                     else
                           SQL=SQL&" AND DateRange=1 "
                           SQL=SQL&" AND DateFrom<='"&RequestPickDate&"' and DateTo>='"&RequestDropDate&"' "
                           SQL=SQL&" AND ( (BookingDateFrom<='"&BookingDate&"' and BookingDateTo>='"&BookingDate&"') "
                           SQL=SQL&" or (BookingDateFrom is Null or YEAR(BookingdateFrom)=2100) )"
                     end if
                     'Response.Write(SQL)
                     Set s_disCat=webConn.Execute(SQL)
                     if NOT s_disCat.EOF then
                           if s_disCat("DiscountRate")<>0 then
                                 DiscountType=s_disCat("DiscountType")
                                 DiscountRate=s_disCat("DiscountRate")
                                 DiscountID=s_disCat("ID")

                                 'Session("RCM273_DiscountName")=s_disCat("DiscountName")
                                 'Session("RCM273_DiscountType")=s_disCat("DiscountType")
                                 'Session("RCM273_DiscountRate")=s_disCat("DiscountRate")
                                 'Session("RCM273_DiscountID")=s_disCat("ID")
                           end if
                     end if
                     s_disCat.close
                     set s_disCat=nothing
             end if
             s_disDate.close
             set s_disDate=nothing

        '---check fixed rate discunt------
              FixedDiscountName=""
              FixedDiscountRate=0
              FixedDiscountID=0
         if Session("RCM273_PromoCode")<>"" then
            SQL=" select df_extraFeeID,ExtraFees.Name,ExtraFees.Fees  FROM DiscoutFixedRate, CampaignCode, ExtraFees "
            SQL=SQL&" WHERE df_CampaignCodeID=CampaignCode.ID and df_extraFeeID=ExtraFees.ID "
            SQL=SQL&" AND CampaignCode='"&Left(tidyup(Session("RCM273_PromoCode")),30)&"' "
            SQL=SQL&" and (df_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or df_LocationID=0) "
            SQL=SQL&" and (df_CategoryID="&s_cs("ID")&" or df_CategoryID=0) "
            SQL=SQL&" and df_NoDays<="&Cint(Session("RCM273_TotalRentalDays"))&" "
            SQL=SQL&" and df_DiscoutStart<='"&Session("RCM273_RequestPickDate")&"' and df_DiscoutEnd>='"&Session("RCM273_RequestPickDate")&"'  "
       else
            SQL=" select df_extraFeeID,ExtraFees.Name,ExtraFees.Fees  FROM DiscoutFixedRate,  ExtraFees "
            SQL=SQL&" WHERE df_CampaignCodeID=0 and df_extraFeeID=ExtraFees.ID "
             SQL=SQL&" and (df_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or df_LocationID=0) "
            SQL=SQL&" and (df_CategoryID="&s_cs("ID")&" or df_CategoryID=0) "
            SQL=SQL&" and df_NoDays<="&Cint(Session("RCM273_TotalRentalDays"))&" "
            SQL=SQL&" and df_DiscoutStart<='"&Session("RCM273_RequestPickDate")&"' and df_DiscoutEnd>='"&Session("RCM273_RequestPickDate")&"'  "
      end if
            ' Response.Write SQL
             SET s_f=webConn.Execute(SQL)
             if NOT s_f.EOF then
                    FixedDiscountName=s_f("Name")
                    FixedDiscountRate=s_f("Fees")
                    FixedDiscountID=s_f("df_extraFeeID")
                 'Response.Write s_f("Fees")
             end if
             s_f.close
            set s_f=nothing

 end if '---------if  categoryAvaliable="Yes" and categoryAvaliable4Relocation="Yes" then

            '--for each season, get the rate
                  j=j+1
                '--check if image exists-------------
            Dim fs
          theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_cs("ImageName")
           '---check location images
             CategoryDesc=s_cs("Size")
             if s_cs("WebDesc") <>"" then
                        CategoryDesc=s_cs("WebDesc")
            end if

                   '---check location images
            SQL="SELECT * FROM CarSizeLocation WHERE  (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' and CategoryID="&s_cs("ID")&"  "
            Set s_li=webConn.Execute(SQL)
            if NOT s_li.EOF then
                       if s_li("ImageName")<>"" then
                       theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_li("ImageName")
                       end if
                       if s_li("WebDesc")<>"" then
                       CategoryDesc =s_li("WebDesc")
                       end if
               end if
               s_li.close
               Set s_li=nothing
             SQL="SELECT * FROM CarSizeLocation WHERE  (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' and CategoryID="&s_cs("ID")&"  "
            Set s_li=webConn.Execute(SQL)
            if NOT s_li.EOF then
                       if s_li("ImageName")<>"" then
                       theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_li("ImageName")
                       end if
                       if s_li("WebDesc")<>"" then
                       VehicleDesc =s_li("WebDesc")
                       end if
               end if
               s_li.close
               Set s_li=nothing




'---table for each vehciel
Response.Write("<div class=""row"">")
Response.Write("<div class=""col-xs-7"">")

                  '--left side table with images----
                  Response.Write("<p class=""text-center""><img src='"&theimage&"' class=""img-rounded vehicle-image"" /></p>")
        
                  if s_cs("noSmallCase")<>"0" then
                     Response.Write("<img src='SmallCase.gif' alt='No. of Small Case'  border=""0"" />x"&s_cs("noSmallCase")&" ")
                  else
                     Response.Write(" ")
                  end if
                  if s_cs("noLargeCase")<>"0" then
                     Response.Write("<img src='LargeCase.gif' alt='No. of Large Case' border=""0"" />x"&s_cs("noLargeCase")&" ")
                  else
                     Response.Write(" ")
                  end if
                  if s_cs("noChildren")<>"0" then
                     Response.Write("<img src='Children.gif' alt='No. of Children' border=""0"" />x"&s_cs("noChildren")&" ")
                  else
                     Response.Write(" ")
                  end if
                  if s_cs("noAdults")<>"0" then
                     Response.Write("<img src='Adult.gif' alt='No. of Adults' border=""0"" />x"&s_cs("noAdults")&" ")
                  else
                     Response.Write(" ")
                  end if

Response.Write("</div>")

            '---table for the right side category desc, rate and booking button
                  Response.Write("<div class=""col-xs-5"">")
                  Response.Write("<h2> "&CategoryDesc&" </h2>")
                  Response.Write("<hr />")

                  Response.Write(" "&s_cs("VehicleDesc")&" ")
                  if      s_cs("Categoryspecial")<>"" then
                     Response.Write(" "&s_cs("Categoryspecial")&" ")
                end if



   if Less1DayHireHourRate=0 then'--------------
          '---the rates-----------
          costEachSeason=0
                totalRate=0
               TotalRentalDays=0
                totalDays4Rate=0
                LocationRate=1 '--default this as rate available
                  for i=1 to SeasonCount
                         DiscountRate=0
                         DiscountID=0
                         DiscountType="p"
                        'Get the Season Start Date to use for Discount
                        SQL="select DiscountRate, DiscountType from Discount	"
                        SQL=SQL&"WHERE WebItems=1 AND (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' "
                        SQL=SQL&"AND (DropoffLocationID)='0' AND CarSizeID=0 AND CampaignCodeID=0 AND DateRange=1 "
                        SQL=SQL&"AND (DateFrom<=(select StartDate from Season where ID = '"&Cint(Session("RCM273_SeasonID"&i&""))&"') "
                        SQL=SQL&"and DateTo>=(select StartDate from Season where ID = '"&Cint(Session("RCM273_SeasonID"&i&""))&"') ) "
                        SQL=SQL&"AND ( (BookingDateFrom<='"&BookingDate&"' and BookingDateTo>='"&BookingDate&"') "
                        SQL=SQL&"or (BookingDateFrom is Null or YEAR(BookingdateFrom)=2100) )"
                        'Response.Write(SQL)
                        Set s_DisRate=webConn.Execute(SQL)
                        if NOT s_DisRate.EOF   then
                          DiscountType=s_DisRate("DiscountType")
                          DiscountRate=s_DisRate("DiscountRate")
                        end if

                         '--for each Season Rate
                           SQL="SELECT * FROM QCarRateDetails "
                           SQL=SQL&"WHERE CarSizeID="&s_cs("ID")&" "
                           SQL=SQL&"AND CarRateStructureID="&Cint(Session("RCM273_RateStructureID"&i&""))&" "
                           SQL=SQL&"AND (SeasonID)='"&Cint(Session("RCM273_SeasonID"&i&""))&"' "
                          'Response.Write(SQL)
                           Set s_m=webConn.Execute(SQL)
                           Rate=0
                          if NOT s_m.EOF   then
                                  if ISNull(s_m("Rate"))<>True and s_m("Rate")<>0 then

                                    if DiscountType="p" then
                                       Rate=s_m("Rate")*(1-DiscountRate/100)
                                    else
                                       Rate=s_m("Rate")-DiscountRate
                                    end if
                                    NoOfDaysEachSeason=Session("RCM273_NoOfDaysEachSeason"&i&"")
                                    costEachSeason=Rate*NoOfDaysEachSeason
                                    totalRate=totalRate+costEachSeason
                                  end if
                         end if

                          '----Aug 2012, rita , if rate=0  and use weekendRate but no weekendRate set up, then use the normarl rate
                         if  Rate=0 and WeekendRate="Yes"   then
                                    SQL= "SELECT Top 1 (FromDay),ToDay, ID, RateName FROM CarRateStructure where ID<>1 "
                                    SQL=SQL&" AND FromDay<="&Cint(Session("RCM273_DaysForRate"))&" order by FromDay desc "
                                    Set s_rs=webConn.Execute(SQL)
                                    if NOT s_rs.EOF   then  '---  not weekend (more then 2 days)
                                             RateName=s_rs("RateName")
                                             RateStructureID=s_rs("ID")
                                             'Session("RCM273_RateStructureID"&i&"")=s_rs("ID")
                                    end if
                                    s_rs.Close
                                    SET s_rs=nothing
                                    SQL="SELECT * FROM QCarRateDetails  where CarRateStructureID="&Cint(RateStructureID)&" and  SeasonID='"&Cint(Session("RCM273_SeasonID1"))&"'  and LocationID="&CINT(Session("RCM273_PickupLocationID"))&"  and CarSizeID="&s_cs("ID")&"  "
                                    Set s_r=webConn.Execute(SQL)
                                    IF NOT s_r.EOF THEN
                                          if ISNull(s_r("Rate"))<>True then
                                                if DiscountType="p" then
                                                   Rate=s_m("Rate")*(1-DiscountRate/100)
                                                else
                                                   Rate=s_m("Rate")-DiscountRate
                                                end if
                                                NoOfDaysEachSeason=Session("RCM273_NoOfDaysEachSeason"&i&"")
                                                costEachSeason=Rate*NoOfDaysEachSeason
                                                totalRate=totalRate+costEachSeason
                                          end if
                                    END IF
                                    s_r.close
                                    set s_r=nothing

                        end if
                        if Rate>0 then
                                 if Session("RCM273_useEverageRate")<>"Yes" then

                                  Response.Write("<div class=""row"">")
                                  Response.Write("<div class=""col-xs-9"">")
                                 
                                 Response.Write("<p class=""text-left"">"&NoOfDaysEachSeason&" Days at ") 

                                    if  DiscountRate >0 then
                                          
                                    Response.Write("<span class=""text-muted""><s>"&Session("RCM273_CompanyCurrency")&" "&s_m("Rate")&"</s></span> <span class=""text-danger"">")
                                        
                                    end if

                                 Response.Write(""&Session("RCM273_CompanyCurrency")&" "&FormatNumber(Rate,2)&"" )

                                  if  DiscountRate >0 then
                                  Response.Write("</span>")
                                  end if

                                  Response.Write(" per day")

                                    if  DiscountRate >0 then
                                          if DiscountType="p"  then
                                                Response.Write("<br /><span class=""label label-danger"">-"&DiscountRate&"% Discount</span>")
                                          else
                                                Response.Write("<br /><span class=""label label-danger"">-"&Session("RCM273_CompanyCurrency")&" "&DiscountRate&" Discount</span>")
                                          end if
                                    end if

                                Response.Write("</p></div>")

                                  Response.Write("<div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(costEachSeason,2)&"</p>")
                                  Response.Write("</div></div>")

                                    
                                    
                                 end if
                        else ' ----rate =0

                                    LocationRate=0

                                    NoOfDaysEachSeason=Session("RCM273_NoOfDaysEachSeason"&i&"")

                                   if Session("RCM273_useEverageRate")<>"Yes" then
                                    Response.Write("<div class=""row"">")
                                  Response.Write("<div class=""col-xs-9"">")
                                    Response.Write("<p class=""text-left"">"&NoOfDaysEachSeason&" Days </p><div class=""col-xs-3""><p class=""text-right""><span class=""text-danger"">No Rate available</span></p></div></div>")
                                    end if

                        end if
                        s_m.close
                        set s_m=nothing
                        TotalRentalDays=TotalRentalDays+NoOfDaysEachSeason
                        totalDays4Rate=totalDays4Rate+NoOfDaysEachSeason
                 next
           '--for display Everage rate not the seasonal rate
           '===should use the totalDays4Rate  (added up  days of each season ) to calculate the everayge rate, not the Session("actureTotalDays"),
           '--eg with pickup 25/Dec/2012 11:00 to  15/Jan/2013 12:00 , total min 30300,  total days 21.04,Session("actureTotalDays")=21.04
           '---but if they have set Part Day Rate 0-5 hours as 20% rate, the 1 hour should be charged as 0.2 days
           '---the number of days  will be 21.2 days , totalDays4Rate=21.2 days
             EverageRate=FormatNumber(totalRate/totalDays4Rate,2) '--used for free day special,
             totalRate=EverageRate*totalDays4Rate
             TotalRentalDays=totalDays4Rate
             Session("RCM273_totalDays4Rate")=totalDays4Rate '--used in step3 to calculate average rate


            TotalCost=totalRate
            if Session("RCM273_useEverageRate")="Yes" then
                           if  DiscountRate >0 then
                                                 Response.Write("<tr><td class='text' align='left'  nowrap='nowrap'>Rate &nbsp;"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(EverageRate,2)&" x "&TotalRentalDays&" days<font color=red>(includes a "&DiscountRate&"% Discount)</font></td><td class='text' align=right >"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(totalRate,2)&"</td></tr>")
                           else
                                                 Response.Write("<tr><td class='text' align='left'  nowrap='nowrap'>Rate &nbsp;"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(EverageRate,2)&" x "&TotalRentalDays&" days</td><td class='text' align=right >"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(totalRate,2)&"</td></tr>")

                           end if
             end if

           if FixedDiscountRate<>0 then
               totalRate=totalRate+FixedDiscountRate
               Response.Write("<tr><td class=""red"" align=""left""  nowrap=""nowrap"">"&FixedDiscountName&" <td class=""red"" align=""right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(FixedDiscountRate,2)&" off</td></tr>")
            end if
                 ' if  DiscountRate>0 then
              ' if DiscountType="p" then
              ' Response.Write(""&DiscountRate&"% Discount included <br />")
              ' else
              ' Response.Write("$"&DiscountRate&" Discount included <br />")
              ' end if
            'end if
            NoFreeDays=0
             SQL=" select  fd_NoofFreeDays,fd_apply2EveryNoDays,fd_extraFeeID,max(fd_NoDays) as MinDay "
            SQL=SQL&" FROM FreeDaySpecial left join CampaignCode on fd_CampaignCodeID=CampaignCode.ID  "
            SQL=SQL&" WHERE fd_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
            if Session("RCM273_PromoCode")<>"" then
               SQL=SQL&" AND CampaignCode='"&Left(tidyup(Session("RCM273_PromoCode")),30)&"'  "
            else
               SQL=SQL&" AND fd_CampaignCodeID=0  "
         end if
         SQL=SQL&" and (fd_CategoryID='"&s_cs("ID")&"' or fd_CategoryID=0) "
         SQL=SQL&" and fd_NoDays<="&CINT(Session("RCM273_TotalRentalDays"))&" "
         SQL=SQL&" and fd_FreeDayStart<='"&Session("RCM273_RequestPickDate")&"' and fd_FreeDayEnd>='"&Session("RCM273_RequestPickDate")&"' "
         SQL=SQL&" AND fd_NoDays in (select max(fd_NoDays)   "
         SQL=SQL&" FROM  FreeDaySpecial left join CampaignCode on fd_CampaignCodeID=CampaignCode.ID  "
         SQL=SQL&" where fd_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
         if Session("RCM273_PromoCode")<>"" then
            SQL=SQL&" AND CampaignCode='"&Left(tidyup(Session("RCM273_PromoCode")),30)&"'  "
         else
            SQL=SQL&" AND fd_CampaignCodeID=0  "
      end if
      SQL=SQL&" and (fd_CategoryID='"&s_cs("ID")&"' or fd_CategoryID=0) "
      SQL=SQL&" and fd_NoDays<="&CINT(Session("RCM273_TotalRentalDays"))&"   "
      SQL=SQL&" and fd_FreeDayStart<='"&Session("RCM273_RequestPickDate")&"' and fd_FreeDayEnd>='"&Session("RCM273_RequestPickDate")&"'  ) "
      SQL=SQL&" GROUP by fd_NoofFreeDays, fd_apply2EveryNoDays,fd_extraFeeID, fd_NoDays "
            'Response.Write(sql)

            SET s_l=webConn.Execute(SQL)
            if NOT s_l.EOF then
                     ExtraFeeID=s_l("fd_extraFeeID")
                     NoFreeDays=s_l("fd_NoofFreeDays")
                     if  s_l("fd_apply2EveryNoDays") =True then
                           NoFreeDays=NoFreeDays * Fix(Session("RCM273_TotalRentalDays")/s_l("MinDay"))
                     end if
                     FreeDayValue=Round((EverageRate*NoFreeDays),2)
                   TotalCost=TotalCost - FreeDayValue
                     Response.Write("<tr><td class=""redtext"">You qualify for a "&NoFreeDays&" Free Day Special!<td class=""text"" align=""right""> "&Session("RCM273_CompanyCurrency")&"-"&FreeDayValue&"</td></tr>")

             END IF

 else '--  if  Less1DayHireHourRate>0 then


                             rentalTotalDays=1  '---is the no of days rate charged, used in kms free
                           LocationRate=Less1DayHireHourRate
                           Response.Write("<tr><td class=""text"" align=""left""  nowrap=""nowrap"">&nbsp;"&RateName&" rate @ <td class=""text"" align=""right""  nowrap=""nowrap""> "&Session("RCM273_CompanyCurrency")&" "&Less1DayHireHourRate&"</td></tr>")
                        '---code was missing---
                        totalRate=Less1DayHireHourRate
                        TotalCost=Less1DayHireHourRate

  end if '---------if Less1DayHireHourRate=0 then'--------------


Response.Write("<strong>Additional fees:</strong>")


Session("RCM273_RelocationFee")=0
 RelocationFeeGST=0
         RelocationFeeStampDuty=0

              '--1. check Relocation record (with caterory, date range)
         SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
         SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID and (CarSizeID)='"&s_cs("ID")&"'  and Mandatory=0 "
         SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"'  "
         SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
                  '--next line of code will return the max Minbookingday if there are 2 records entered for same conditin
         SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
         'Response.Write(SQL)
          Set s_o=webConn.Execute(SQL)
          if  s_o.EOF THEN
                  s_o.close
                  Set s_o=Nothing
                  '--2. if no vehicle category Relocation fee found, check Relocation record (with  date range only, CarSizeID=0)
                  SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                  SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and Mandatory=0  "
                  SQL=SQL&" AND CarSizeID=0 "
                  SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"'  "
                  SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
                  SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                  'Response.Write(SQL)
                   Set s_o=webConn.Execute(SQL)
                   if  s_o.EOF THEN
                        s_o.close
                        Set s_o=Nothing
                        '--3. if no vehicle category Relocation fee found, check Relocation record (with  category, no date ragne)
                        SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                        SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and (CarSizeID)='"&s_cs("ID")&"' "
                        SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"'  "
                        SQL=SQL&"  AND  Year(PickupDateFrom)=2100  and Mandatory=0 "
                        SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                        'Response.Write(SQL)
                        Set s_o=webConn.Execute(SQL)
                        if  s_o.EOF THEN
                              s_o.close
                              Set s_o=Nothing
                                    '--4. if no vehicle category Relocation fee found, check Relocation record (with no category, no date ragne)
                              SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                              SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID   and Mandatory=0 "
                              SQL=SQL&" AND CarSizeID=0 "
                              SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"'  "
                              SQL=SQL&"  AND  Year(PickupDateFrom)=2100 "
                              SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                              'Response.Write(SQL)
                              Set s_o=webConn.Execute(SQL)
                        end if
                  end if

         end if

         if NOT s_o.EOF THEN
               DO WHILE NOT s_o.EOF
               'Response.Write s_o("DaysNocharge")
                  if  s_o("DaysNocharge")=0 or (s_o("DaysNocharge")>0 and Session("RCM273_TotalRentalDays")<s_o("DaysNocharge"))  then
                     if s_o("GST")="True" then
                        RelocationFeeGST=RelocationFeeGST+s_o("Fees")
                     end if
                     if s_o("StampDuty")="True" then
                        RelocationFeeStampDuty=RelocationFeeStampDuty+s_o("StampDuty")
                     end if
                     Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_o("Fees")
                     Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_o("Name")&" </div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_o("Fees"),2)&"</div></div>")



                 end if
                s_o.MoveNext
                Loop
         end if
         s_o.CLOSE
         SET s_o=NOTHING

'Response.Write("<Br>")
      '---Aug 2014, new WebExtralAdditionalFees, added for wicked AU, allow to add additional fee for date range
            SQL="SELECT DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebExtralAdditionalFees,ExtraFees "
            SQL=SQL&" WHERE  PickuplocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' "
            SQL=SQL&"  AND (DropofflocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"' or DropofflocationID=0) "
             SQL=SQL&" and (CarSizeID='"&s_cs("ID")&"' or WebExtralAdditionalFees.CategoryTypeID="&s_cs("CategoryTypeID")&" or  (CarSizeID=0 and WebExtralAdditionalFees.CategoryTypeID=0)   ) "
            SQL=SQL&" and Mandatory=0 AND ExtraFeeID=ExtraFees.ID "
          SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
            SQL=SQL&"  AND (DaysNocharge>"&CINT(Session("RCM273_TotalRentalDays"))&" or DaysNocharge=0 )"
            'Response.Write SQL
           ' Response.Write("<Br><br>")
            Set s_o=webConn.Execute(SQL)
            DO WHILE NOT s_o.EOF
                     if s_o("GST")="True" then
                        RelocationFeeGST=RelocationFeeGST+s_o("Fees")
                     end if
                     if s_o("StampDuty")="True" then
                        RelocationFeeStampDuty=RelocationFeeStampDuty+s_o("StampDuty")
                     end if
                     Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_o("Fees")
                     Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_o("Name")&" </div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_o("Fees"),2)&"</div></div>")
            s_o.MoveNext
            Loop
            s_o.CLOSE
            SET s_o=NOTHING


 


 '-----Pickup Location After hours and befores fee, check if the pickup time is between the office time-------------
        Session("RCM273_AfterHoursFee")=0
        Session("RCM273_PickupAfterHoursFee")=0
        AfterHoursFeeGST=0
        AfterHoursFeeStampDuty=0
        if Session("RCM273_PickupAfterHourFeeID")<>0  then 'check the time
               AfterHoursFee=0
               Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&CINT(Session("RCM273_PickupAfterHourFeeID"))&"' ")
               if s_st("Fees")  <>0 then
                       '---Pickup Location After Hour fees   'do not convert to CDate
                     PickupOpeningTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeOpeningTime"))
                     PickupClosingTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeClosingTime"))
                     if (Session("RCM273_RequestPickDateTime") < PickupOpeningTime) or (Session("RCM273_RequestPickDateTime") > PickupClosingTime) THEN
                            Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_st("Name")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_st("Fees"),2)&"</div></div>")

                           Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_st("Fees")
                            if AfterHoursGST="True" then
                         AfterHoursFeeGST=s_st("Fees")
                      end if
                      if AfterHoursStampDuty="True" then
                         AfterHoursFeeStampDuty=s_st("Fees")
                      end if
                     end if
            end if
            s_st.close
            SET s_st=nothing

      end if

    '------------Dropoff Location After Hour Fees ------------------
    Session("RCM273_DropoffAfterHoursFee")=0
    if Session("RCM273_DropoffAfterHourFeeID")<>0  then
                  Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&CINT(Session("RCM273_DropoffAfterHourFeeID"))&"' ")
                 if s_st("Fees")  <>0 then
                         DropoffOpeningTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeOpeningTime"))
                        DropoffClosingTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeClosingTime"))
                        if (Session("RCM273_RequestDropDateTime") < DropoffOpeningTime) or (Session("RCM273_RequestDropDateTime") > DropoffClosingTime) THEN
                               Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_st("Fees")
                                'Response.Write("<tr><td class=""text"" align=""right"" nowrap=""nowrap"">"&s_st("Name")&"  "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_st("Fees"),2)&"</td></tr>")
                              Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_st("Name")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_st("Fees"),2)&"</div></div>")
                         if AfterHoursGST="True" then
                         AfterHoursFeeGST= AfterHoursFeeGST+s_st("Fees")
                      end if
                      if AfterHoursStampDuty="True" then
                         AfterHoursFeeStampDuty=AfterHoursFeeStampDuty + s_st("Fees")
                      end if

                        end if
            end if

            s_st.close
            SET s_st=nothing
         end if
         TotalCost=TotalCost+Session("RCM273_RelocationFee")


    '----------------check Holiday Charges for pickup  from table HolidayextraFees  ---------------------
         SQL="SELECT Fees,eh_HolidayName FROM ExtraFees4Holiday,ExtraFees WHERE eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestPickDate")&"' and  eh_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"')  "
        ' Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)

       if NOT s_ex.EOF THEN
               Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_ex("eh_HolidayName")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_ex("Fees"),2)&"</div></div>")
               TotalCost=TotalCost+ s_ex("Fees")
      end if
      s_ex.close
      Set s_ex=Nothing

      '--check Holiday Charges for dropoff  from table HolidayextraFees
       SQL="SELECT Fees,eh_HolidayName  FROM ExtraFees4Holiday,ExtraFees WHERE   eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestdropDate")&"' and  eh_LocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"') "
       'Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)
      if NOT s_ex.EOF THEN
               Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_ex("eh_HolidayName")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_ex("Fees"),2)&"</div></div>")
              TotalCost=TotalCost+ s_ex("Fees")
      end if
      s_ex.close
      Set s_ex=Nothing
          '----------------END check Holiday Charges for pickup  from table HolidayextraFees  ---------------------

   '---------------Mandatory Insurance-------------------

EachInsuranceFees=0
 TotalCost=TotalCost+EachInsuranceFees

      '----Mandatory Feees-------
          SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
         SQL=SQL&" AND (VehicleSizeID='"&s_cs("ID")&"' or VehicleSizeID=0) "
         SQL=SQL&" AND (VehicleSizeID='"&Session("RCM273_CarSizeID")&"' or VehicleSizeID=0) "
         SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
       SQL=SQL&" and WebItems=1 and InsuranceExtra=0  and Mandatory=1 AND inUse=1 and PercentageTotalCost=0 and MerchantFee=0 order by Type,Name"
     Set s_ex=webConn.Execute(SQL)
     j=0
     EachMandatoryExtraFees=0
     TotalMandatoryExtraFees=0
     MandatoryExtraStamp=0
     MandatoryExtraGST=0
     MandatoryExtraFeesNoDays=1
     DO WHILE NOT s_ex.EOF
         'Response.Write(SQL)
            MandatoryExtraFeesNoDays=1

             if s_ex("Type")="Daily" then

                  '---daily fees always use   Session("RCM273_TotalRentalDays") (fixed number)
                        MandatoryExtraFeesNoDays=Session("RCM273_TotalRentalDays")
                        EachMandatoryExtraFees=MandatoryExtraFeesNoDays*s_ex("Fees")
             elseif s_ex("Type")="Fixed" then
                        EachMandatoryExtraFees = s_ex("Fees")
             else
                       EachMandatoryExtraFees = (s_ex("Fees")/100)*totalRate
             end if
                     '--27/Jun/2009 for Daily and % fees allow check max price in step4
                if s_ex("MaxPrice") >0  and  EachMandatoryExtraFees>s_ex("MaxPrice") then
                                 EachMandatoryExtraFees=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and EachMandatoryExtraFees<s_ex("MaxPrice") then
                                  EachMandatoryExtraFees=s_ex("MaxPrice")
                end if

                     TotalMandatoryExtraFees=EachMandatoryExtraFees+ TotalMandatoryExtraFees

                   Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_ex("Name")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(EachMandatoryExtraFees,2)&"</div></div>")



      s_ex.MoveNext
     j=j+1
     Loop
     s_ex.close
     SET s_ex=nothing
     TotalCost=TotalCost+TotalMandatoryExtraFees


            SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
          SQL=SQL&" AND (VehicleSizeID='"&s_cs("ID")&"' or VehicleSizeID=0) "
          'SQL=SQL&" and (CategoryTypeID ="&s_cs("CategoryTypeID")&" or CategoryTypeID =0 ) "
          SQL=SQL&" and WebItems=1 and Mandatory=1 AND inUse=1 and "
           SQL=SQL&" (PercentageTotalCost=1 or MerchantFee=1) "
            SQL=SQL&" and Type='Percentage' order by percentageTotalCost,MerchantFee,Name"
         'Response.Write(SQL)
         Set s_ex=webConn.Execute(SQL)
         j=0
         EachMerchantFee=0
         TotalMerchantFee=0
         MerchantFeeStamp=0
         MerchantFeeGST=0
         MerchantFeeNoDays=1
         BankBaseCalculationFee= TotalCost
         percentageBaseCalculationFee= TotalCost

         DO WHILE NOT s_ex.EOF
         'Response.Write(SQL)
                  EachMerchantFee =(s_ex("Fees")/100)*totalCost

                  '--27/Jun/2009 for % fees allow check max price in step4
                  if s_ex("MaxPrice") >0  and  EachMerchantFee>s_ex("MaxPrice") then
                         EachMerchantFee=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and  EachMerchantFee<s_ex("MaxPrice") then
                           EachMerchantFee=s_ex("MaxPrice")
                end if
                 EachMerchantFee =Round(EachMerchantFee,2)

                  TotalCost=TotalCost+EachMerchantFee
                  BankBaseCalculationFee=BankBaseCalculationFee+EachMerchantFee
               Response.Write("<div class=""row""><div class=""col-xs-9""><p class=""text-left"">"&s_ex("Name")&"</div><div class=""col-xs-3""><p class=""text-right"">"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(EachMerchantFee,2)&"</div></div>")

          s_ex.MoveNext
         j=j+1
         Loop
         s_ex.close
         SET s_ex=nothing

         GSTRate=0.10
          GSTRate= 0
          Set s_st=webConn.Execute("SELECT * FROM QStampDuty WHERE ID="&CINT(Session("RCM273_PickupLocationID"))&"  ")
          If not s_st.EOF then
             if ISNull(s_st("StampDuty"))<>True then
                StampDutyRate=s_st("StampDuty")/100
                end if
                    if ISNull(s_st("GSTRate"))<>True then
                       GSTRate= s_st("GSTRate")/100
                    end if
                    Session("RCM273_TaxName1") = s_st("SalesTax")
                    Session("RCM273_TaxName2") = s_st("LocalTaxDesc")
          end if
          s_st.close
          set s_st=nothing
           Session("RCM273_GSTInclusive") = "Yes"
         Set s_st=webConn.Execute("SELECT * FROM SystemTable WHERE Code='GSINC'  ")
         If not s_st.EOF then
               Session("RCM273_GSTInclusive") = s_st("syValue")
        END IF
         s_st.CLOSE
        SET s_st=NOTHING

        ExtraGST=0
        GSTInsurance=0
        if Session("RCM273_GSTInclusive") = "Yes" then
         GST=(TotalCost)*(1-100/(100+GSTRate*100))
         else
          GST=( totalRate +  AfterHoursFeeGST + RelocationFeeGST + ExtraGST+ MandatoryExtraGST + GSTInsurance + GSTEachMerchantFee)*GSTRate

         'GST=(TotalCost)*GSTRate
         end if

          GST=Round(GST,2)

   '===============###### if using hourly rates and the TotalCost display as 0 #####==================
   '===============###### those code were missing, do a search on this and added in #####==================
   '===============###### totalRate=Less1DayHireHourRate #####==================
   '===============###### TotalCost=Less1DayHireHourRate #####==================
     Response.Write("<hr /><div class=""row total-estimate"">")

        Response.Write("<div class=""col-xs-6""><p class=""lead""><strong>Total estimate:</strong></p></div><div class=""col-xs-6""><p class=""text-right ""><span class=""lead ""><strong ><mark>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(TotalCost,2)&"</mark></strong></span>")
     if Session("RCM273_GSTInclusive") = "Yes" then
     Response.Write("<br /><span class=""text-muted""><em><small>includes "&Session("RCM273_TaxName1")&" "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(GST,2)&"</small></em></span>")
      end if

      Response.Write("</p></div></div>")


         '=============================================




               if   categoryAvaliable="No"  then
                     if MinBookingDay=0 then
                           Response.Write("<div class=""alert alert-danger"" role=""alert"">This vehicle category is Unavailable</div>")
                     else
                           Response.Write("<div class=""alert alert-danger"" role=""alert"">This vehicle category has a  "&MinBookingDay&" day minimum hire period</div>")
                     end if
            elseif    categoryAvaliable4Relocation="No" then
                   if MinBookingDay=0 then
                           Response.Write("<div class=""alert alert-danger"" role=""alert"">Sorry, this vehicle category is Unavailable for one way hire between your selected location</div>")
                     else
                           Response.Write("<div class=""alert alert-danger"" role=""alert"">This vehicle category has a  "&RelocationMinBookingDay&" day minimum hire period</div>")
                     end if

                elseif  Unavailable="False"  then
                        if CInt(Session("RCM273_driverage"))- s_cs("AgeYoungestDriver")>=0 then
                            if   LocationRate>0 then
                                 if categoryStatus=2 then
                                       Response.Write("<div class=""alert alert-danger"" role=""alert"">LIMITED AVAILABILITY.</div>")
                                 end if
                                 'Response.Write("<tr><td><td    align=""center""  class=""text"">")
                                  '---Jul 2012, new changes by Rita, need pass all the booking details by QueryString, when user open 2 windows and the sessions will be maxed up and cuase lots of problems
                                  'Response.Write("<a href='webstep3.asp?categoryStatus="&categoryStatus&"&CarSizeID="&s_cs("ID")&"&TotalRentalDays="&TotalDays&"&PickupLocationID="&CINT(Request.Form("PickupLocationID"))&"&DLocationID="&CINT(Session("RCM273_DropoffLocationID"))&"&PickDateTime="&RequestPickDate&"&DropoffDateTime="&RequestDropDate&"&PickupTime="&RequestPickTime&"&DropoffTime="&RequestDropTime&"' class='rcmbutton btn btn-info btn-block btn-lg text-uppercase' >")

                                  Response.Write("<a href='webstep3.asp?categoryStatus=2&CarSizeID="&s_cs("ID")&"&TotalRentalDays="&TotalDays&"&PickupLocationID="&CINT(Request.Form("PickupLocationID"))&"&DLocationID="&CINT(Session("RCM273_DropoffLocationID"))&"&PickDateTime="&RequestPickDate&"&DropoffDateTime="&RequestDropDate&"&PickupTime="&RequestPickTime&"&DropoffTime="&RequestDropTime&"' class='rcmbutton btn btn-info btn-block btn-lg text-uppercase' >")

                                 Response.Write("Book now <span class=""glyphicon glyphicon-chevron-right"" aria-hidden=""true""></span></a></div>")
                            else
                                 'Response.Write("<tr><td><td    align=""center""  class=""text"">")
                                 Response.Write("<div class=""alert alert-danger"" role=""alert"">No Rate Available</div>")
                            end if

                        else
                           Response.Write("<div class=""alert alert-danger"" role=""alert"">Not available for hire to drivers under "&s_cs("AgeYoungestDriver")&" years of age</div>")
                        end if
                 else
                        if Unavailable="True"  then
                        Response.Write("<div class=""alert alert-danger"" role=""alert"">BOOKED OUT<br>"&CategoryDesc&"</div>")
                        else
                           if CInt(Session("RCM273_driverage"))- s_cs("AgeYoungestDriver")>=0 then

                                 if LocationRate>0 then
                                       if categoryStatus=2 then
                                                Response.Write("<div class=""alert alert-danger"" role=""alert"">LIMITED AVAILABILITY.</div>")
                                       end if
                                       'Response.Write("<tr><td><td   align=""center"" class=""text"">")
                                        '---Jul 2012, new changes by Rita, need pass all the booking details by QueryString, when user open 2 windows and the sessions will be maxed up and cuase lots of problems
                                       Response.Write("<a href='webstep3.asp?categoryStatus="&categoryStatus&"&CarSizeID="&s_cs("ID")&"&TotalRentalDays="&TotalDays&"&PickupLocationID="&CLNG(Request.Form("PickupLocationID"))&"&DLocationID="&CINT(Session("RCM273_DropoffLocationID"))&"&PickDateTime="&RequestPickDate&"&DropoffDateTime="&RequestDropDate&"&PickupTime="&RequestPickTime&"&DropoffTime="&RequestDropTime&"' class='rcmbutton btn btn-success btn-block btn-lg text-uppercase' >")

                                       Response.Write("Choose <span class=""glyphicon glyphicon-chevron-right"" aria-hidden=""true""></span></a></div>")
                                       'Response.Write("</a></td></tr>")
                                  else
                                       'Response.Write("<tr><td><td  align=""center""  class=""text"">")
                                       Response.Write("<div class=""alert alert-danger"" role=""alert"">No Rate Available</div>")
                                 end if

                           else
                                    Response.Write("<div class=""alert alert-danger"" role=""alert"">NOT available for hire to drivers under "&s_cs("AgeYoungestDriver")&" years of age</div")

                           end if
                        end if
                end if

               Response.Write("</div>") '---end right side table
            Response.Write("<hr /></div>")
  'Response.Write("</td></tr></table>")




         s_cs.MoveNext
          Loop
          s_cs.Close
         set s_cs=nothing

         'Response.Write("</td></tr></table>")
         'Response.Write("</td></tr>")


END SUB
              %>
             
     
                      <%

   if Session("RCM273_ErrorMesage")<>"" then

          if Request.QueryString("CatID")<>"" then
          '--for passing the carsizeID from other web page, if they select a vehciel type and with a booking link

            Set s_c = webConn.Execute("SELECT CategoryTypeID FROM CarSize WHERE ID='"&Left(CINT(Request.QueryString("CatID")),3)&"'  ")
            if Not s_c.EOF then
            Session("RCM273_CategoryTypeID")=s_c("CategoryTypeID")
            end if
            s_c.close
            set s_c=nothing
         end if

             TheQutoationForm

             Response.Write("<div class=""alert alert-danger"" role=""alert""> "&Session("RCM273_ErrorMesage")&"</div>")
             'Response.Write("</table></td></tr></table></td></tr></table>")

            '--abandon sessions, if 2 windows open...
            Session.Abandon
            
            Session("RCM273_ErrorMesage")=""
             defaultPickupTime="11:00"
            defaultDrpoffTime="10:00"
            defaultPickupDate=Date+5
           defaultDropoffDate=Date+15

      elseif Request.Form("PickupLocationID")="" then  'end if if Session("RCM273_ErrorMesage")<>"" then
             TheQutoationForm
              Session("RCM273_ErrorMesage")=""
             defaultPickupTime="11:00"
            defaultDrpoffTime="10:00"
            defaultPickupDate=Date+5
           defaultDropoffDate=Date+15
      else  'end if if Session("RCM273_ErrorMesage")<>"" then






        '---15/Apr/08, the following code works when go back to step1
         if Request.QueryString("refid")<>"" and Session("RCM273_referral")="" then
          Session("RCM273_referral")=Left(Request.QueryString("refid"),3)
         end if
     



     '--from webstep2.asp

      '-------------the Pickup and Dropoff date time, No. of booking days----------------

      Session("RCM273_driverage")=0
      if Request.Form("driverage")<>"0" then
      Session("RCM273_driverage")=Left(tidyup(Request.Form("driverage")),3)
      end if
         '-------------if calander used----------------
      'RequestPickDate = Request.Form("Pickupdate")
      'RequestDropDate = Request.Form("dropoffdate")


 if Request.Form("PickupMonth")="" then
         RequestPickDate = Session("RCM273_RequestPickDate")
         RequestDropDate = Session("RCM273_RequestDropDate")
         RequestPickDateTime = Session("RCM273_RequestPickDateTime")
         RequestDropDateTime =Session("RCM273_RequestDropDateTime")


 else
         if Request.Form("PromoCode")<>"" then
                  Session("RCM273_PromoCode")=Left(tidyup(Request.Form("PromoCode")),30)

                 '-- valide the PromoCode

                   Set s_c=webConn.Execute(" select * from CampaignCode where CampaignCode='"&Left(tidyup(Request.Form("PromoCode")),20)&"'")
                      IF s_c.EOF  THEN
                           Session("RCM273_PromoCode")=""
                      end if
                   s_c.close
                   Set s_c=nothing


         end if
         Session("RCM273_CategoryTypeID")=Left(tidyup(Request.Form("CategoryTypeID")),3)
         RequestPickYear = CInt(Request.Form("PickupYear"))
         RequestPickMonth =  Left(MonthName(Request.Form("PickupMonth")),3)
         RequestPickDay  = CInt(Request.Form("PickupDay"))
         RequestPickDate = RequestPickDay&"/"&RequestPickMonth&"/"&RequestPickYear

         RequestDropYear = CInt(Request.Form("DropoffYear"))
         RequestDropMonth =  Left(MonthName(Request.Form("DropoffMonth")),3)
         RequestDropDay  = CInt(Request.Form("DropoffDay"))
         RequestDropDate = RequestDropDay&"/"&RequestDropMonth&"/"&RequestDropYear


         '-------------if no Calander used----------------
         ' RequestPickYear = CInt(Request.Form("PickupYear"))
         ' RequestPickMonth = (Request.Form("PickupMonth"))
         ' RequestPickDay  = CInt(Request.Form("PickupDay"))
         ' RequestPickDate = RequestPickDay&"/"&RequestPickMonth&"/"&RequestPickYear

         ' RequestDropYear = CInt(Request.Form("DropoffYear"))
         ' RequestDropMonth = (Request.Form("DropoffMonth"))
         ' RequestDropDay  = CInt(Request.Form("DropoffDay"))
         ' RequestDropDate = RequestDropDay&"/"&RequestDropMonth&"/"&RequestDropYear

         RequestPickTime=Request.Form("PickupTime")
         RequestDropTime=Request.Form("DropoffTime")

         Session("RCM273_RequestPickTime")=Request.Form("PickupTime")
         Session("RCM273_RequestDropTime")=Request.Form("DropoffTime")
         Session("RCM273_PickupLocationID") = Left(tidyup(Request.Form("PickupLocationID")),3)
         Session("RCM273_DropoffLocationID")= Left(tidyup(Request.Form("DropoffLocationID")),3)

         if Request.Form("DropoffLocationID")="Same" then
                  Session("RCM273_DropoffLocationID") = Left(tidyup(Request.Form("PickupLocationID")),3)
         end if

         if Request.Form("PickupLocationID")="0" or  Request.Form("PickupLocationID")="" then
               Session("RCM273_ErrorMesage")="Please select your Pickup Location."
                 Response.Redirect "webstep2.asp"
         end if
         if Request.Form("DropoffLocationID")="0" or Request.Form("DropoffLocationID")="" then
               Session("RCM273_ErrorMesage")="Please select your Dropoff Location."
                 Response.Redirect "webstep2.asp"
         end if


        if IsDate(RequestPickDate)=True then
            Session("RCM273_RequestPickDate")=RequestPickDate
        end if
        if IsDate(RequestDropDate)=True then
            Session("RCM273_RequestDropDate")=RequestDropDate
        end if

        '----check the pickup and Return date--------------------------
        if IsDate(RequestPickDate)<>True then
                  Session("RCM273_ErrorMesage")="PLEASE CHECK YOUR BOOKING DETAILS."
                    Response.Redirect "webstep2.asp"
        end if

        if IsDate(RequestDropDate)<>True then
                  Session("RCM273_ErrorMesage")="The Return Date does not exist."
                    Response.Redirect "webstep2.asp"
         end if

         if  DateDiff("d", RequestDropDate, RequestPickDate)> 0 then
                  Session("RCM273_ErrorMesage")="Return Date is earlier then Pick up date."
                    Response.Redirect "webstep2.asp"
         end if

          if  RequestDropDate=RequestPickDate and RequestPickTime=RequestDropTime then
            Session("RCM273_ErrorMesage")="Return Date time is same as Pick up date time."
              Response.Redirect "webstep2.asp"

   end if
        if Request.Form("driverage")="0" then
            Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='AGERQ'  ")
            AGERequired="No"
            If not RG.EOF then
               AGERequired=  RG("syValue") '--Track Youngest Driver
            END IF
            RG.CLOSE
            SET RG=NOTHING
            Response.Write("<input type=""hidden"" name=""AGERequired"" size=""10"" value="&AGERequired&">")

            if AGERequired="Yes" then
               Session("RCM273_ErrorMesage")="  Age of Youngest Driver required."
                 Response.Redirect "webstep2.asp"
            end if

   end if
        RequestPickDateTime = CDate(RequestPickDate&" "& RequestPickTime)
        Session("RCM273_RequestPickDateTime")=RequestPickDate&" "& RequestPickTime
        RequestDropDateTime =CDate(RequestDropDate&" "& RequestDropTime)
        Session("RCM273_RequestDropDateTime")=RequestDropDate&" "& RequestDropTime
end if  '--end if Request.Form("PickupMonth")="" then





TheQutoationForm

if Session("RCM273_PickupLocationID")="" then
         Session("RCM273_ErrorMesage")="Please select  requested Pickup and dropoff dates  "
              Response.Redirect "webstep2.asp"
end if

if RequestDropDateTime - RequestPickDateTime<=0 then
           Session("RCM273_ErrorMesage")="The Drop Off Date/Time entered is the same or earlier than the Pick Up Date/Time."
           Response.Redirect "webstep2.asp"

   end if

 '---check Holidays with no pickup and Dropoff -----------
        SQL="SELECT * from syHolidays where (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND Type='P' and StartDate<= '"&RequestPickDate&"' and EndDate>='"&RequestPickDate&"' "
        SET s_hod=webConn.Execute(SQL)
         DO WHILE NOT  s_hod.EOF    '--need the loop

               if s_hod("WeekDays")=0 then
                        if s_hod("ClosingTime")<>"" then '--check the closing time
                              ClosingDateTime= RequestPickDate&" "&s_hod("ClosingTime")
                              TimeDiff= Round((CDate(ClosingDateTime) - CDate(Session("RCM273_RequestPickDateTime")))*1440,2)
                                        'Response.Write(ClosingDateTime)
                            if TimeDiff<0 then
                                          Session("RCM273_ErrorMesage")="Pickup time unavailable, this location closes at "&s_hod("ClosingTime")&"<br>on the requested Pickup date "&RequestPickDate&" - "&s_hod("HolidayName")&". "
                                         Response.Redirect "webstep2.asp"
                               end if
                   else
                     Session("RCM273_ErrorMesage")="The request Pickup date "&RequestPickDate&" is Unavailable. "
                       Response.Redirect "webstep2.asp"
                        end if

                elseif WeekDay(RequestPickDate)=s_hod("WeekDays") then
                 if s_hod("ClosingTime")<>"" then '--check the closing time
                              ClosingDateTime= RequestPickDate&" "&s_hod("ClosingTime")
                              TimeDiff= Round((CDate(ClosingDateTime) - CDate(Session("RCM273_RequestPickDateTime")))*1440,2)
                                        'Response.Write(ClosingDateTime)
                            if TimeDiff<0 then
                                          Session("RCM273_ErrorMesage")="Pickup time unavailable, this location closes at "&s_hod("ClosingTime")&"<Br>on the requested Pickup date "&RequestPickDate&" - "&s_hod("HolidayName")&". "
                                            Response.Redirect "webstep2.asp"
                               end if
                   else
                        TheDay=WeekDayName(s_hod("WeekDays"))
                        Session("RCM273_ErrorMesage")="The request Pickup date "&TheDay&" "&RequestPickDate&" is Unavailable"
                          Response.Redirect "webstep2.asp"
                         end if

            end if
         s_hod.MoveNext
   Loop
        s_hod.close
        set s_hod=nothing

        SQL="SELECT * from syHolidays where (LocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"' AND Type='D' and StartDate<= '"&RequestDropDate&"' and EndDate>='"&RequestDropDate&"' "
        SET s_hod=webConn.Execute(SQL)
        DO WHILE NOT  s_hod.EOF    '--need the loop
               if s_hod("WeekDays")=0 then
                        if s_hod("ClosingTime")<>"" then '--check the closing time
                              ClosingDateTime= RequestDropDate&" "&s_hod("ClosingTime")
                              TimeDiff= Round((CDate(ClosingDateTime) - CDate(Session("RCM273_RequestDropDateTime")))*1440,2)
                                        'Response.Write(TimeDiff)
                            if TimeDiff<0 then
                                              Session("RCM273_ErrorMesage")="The request Dropoff date "&RequestDropDate&" is closing at "&s_hod("ClosingTime")&" - "&s_hod("HolidayName")&". "
                                               Response.Redirect "webstep2.asp"
                                        end if
                   else
                     Session("RCM273_ErrorMesage")="The request Dropoff date "&RequestDropDate&" is Unavailable. "
                       Response.Redirect "webstep2.asp"
                         end if

                   elseif WeekDay(RequestDropDate)=s_hod("WeekDays") then
                     if s_hod("ClosingTime")<>"" then '--check the closing time
                              ClosingDateTime= RequestDropDate&" "&s_hod("ClosingTime")
                              TimeDiff= Round((CDate(ClosingDateTime) - CDate(Session("RCM273_RequestDropDateTime")))*1440,2)
                                        'Response.Write(TimeDiff)
                            if TimeDiff<0 then
                                              Session("RCM273_ErrorMesage")="The request Dropoff date "&RequestDropDate&" is closing at "&s_hod("ClosingTime")&" - "&s_hod("HolidayName")&". "
                                               Response.Redirect "webstep2.asp"
                                        end if
                   else
                     TheDay=WeekDayName(s_hod("WeekDays"))
                     Session("RCM273_ErrorMesage")="The request Dropoff date "&TheDay&" "&RequestDropDate&" is Unavailable. "
                       Response.Redirect "webstep2.asp"
                         end if
            end if
         s_hod.MoveNext
   Loop
        s_hod.close
        set s_hod=nothing


       '---get the pickup location informations----------------------------------------------
        MinBookingDay=1
        OfficeOpeningTime="8:00"
        OfficeClosingTime="17:30"
         PickupOfficeOpeningDateTime =(RequestPickDate&" "&OfficeOpeningTime)
        PickupOfficeClosingDateTime=(RequestPickDate&" "&OfficeClosingTime)
        DropoffOfficeClosingDateTime=(RequestDropDate&" "&Session("RCM273_OfficeClosingTime"))
        DropoffOfficeOpeningDateTime=(RequestDropDate&" "&OfficeOpeningTime)


        Session("RCM273_PickupLocation")=""
        Session("RCM273_PickupLocationCode")=""
        Session("RCM273_LocationEmail")=""
        Session("RCM273_DocPrFix")=""
        Session("RCM273_PickupAfterHourFeeID")=0

        Session("RCM273_MinimunAge")=0
        NoticeRequired=0
         Set s_pl = webConn.Execute("SELECT * FROM QLocationState where (ID) = '"&CINT(Session("RCM273_PickupLocationID"))&"' ")
        if Not s_pl.EOF then
            'PickupAfterHourBooking= s_pl("AfterHourBooking")
              Session("RCM273_PickupLocation")=s_pl("Location")
                Session("RCM273_PickupLocationCode")=s_pl("CityCode")
                Session("RCM273_LocationEmail")=s_pl("Email")
                Session("RCM273_DocPrFix")=s_pl("DocPrFix")
                '--Jul 2012, do not use the minbookingdays from web location, use the min booking days from web cagegory available record

                'MinBookingDay=s_pl("MinBookingDay")
                  NoticeRequired=s_pl("NoticeRequired")
            Session("RCM273_MinimunAge")=s_pl("MinimunAge")

         PickupOfficeOpeningTime=s_pl("OfficeOpeningTime")
         PickupOfficeClosingTime=s_pl("OfficeClosingTime")
              PickupOfficeOpeningDateTime=(RequestPickDate&" "&s_pl("OfficeOpeningTime"))
              PickupOfficeClosingDateTime=(RequestPickDate&" "&s_pl("OfficeClosingTime"))

            AmAfterHoursStart=s_pl("OfficeOpeningTime")
                PmAfterHoursEnd=s_pl("OfficeClosingTime")

                   '--20/Jun/2008, Opening and closing time should use the OfficeOpeningTime record for each weekday (not the location record)
                SQL="SELECT * from OfficeOpeningTime where ot_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' AND ot_WeekDay='"&WeekDay(RequestPickDate)&"'  "
               SET s_oo=webConn.Execute(SQL)
              'Response.Write(SQL)
                OfficeOpeningTimeSet=False
           if Not s_oo.EOF then
                               PickupOfficeOpeningDateTime=(RequestPickDate&" "&s_oo("ot_OfficeOpeningTime"))
                            PickupOfficeClosingDateTime=(RequestPickDate&" "&s_oo("ot_OfficeClosingTime"))
                            PickupOfficeOpeningTime=s_oo("ot_OfficeOpeningTime")
                               PickupOfficeClosingTime=s_oo("ot_OfficeClosingTime")

                               AmAfterHoursStart=s_oo("ot_AmAfterHoursStart")
                               PmAfterHoursEnd=s_oo("ot_PmAfterHoursEnd")
                                   OfficeOpeningTimeSet=True
                end if
                s_oo.close
                set  s_oo=nothing
               'Response.Write(SQL)
                   ' Response.Write(RequestPickDateTime)
                     '          Response.Write("--")
                      '       Response.Write(PickupOfficeOpeningDateTime)
            Session("RCM273_PickupOfficeOpeningTime")=PickupOfficeOpeningTime
            Session("RCM273_PickupOfficeClosingTime")=PickupOfficeClosingTime


            if  RequestPickDateTime < CDate(PickupOfficeOpeningDateTime)  or RequestPickDateTime > CDate(PickupOfficeClosingDateTime)  then
                'if  RequestPickDateTime > CDate(PickupOfficeOpeningDateTime)  or RequestPickDateTime < CDate(PickupOfficeClosingDateTime)  then
                    '-- if pickup is outside office time then Check if location will take bookings outside office hours, Unattended Dropoffs
               if s_pl("AfterHourBooking") = "False" then '--if not allow after hour booking
                              Session("RCM273_ErrorMesage")="Pickup Location "&Session("RCM273_PickupLocation")&"   will not take bookings outside office hours ("&WeekdayName(WeekDay(RequestPickDate))&" "&PickupOfficeOpeningTime&" - "&PickupOfficeClosingTime&").<Br>Please contact the bookings office directly."
                                Response.Redirect "webstep2.asp"
               else '--allow after hour booking, need check if pickup and Return Time inside the  after hours
                         if OfficeOpeningTimeSet=False  then '--if no records set for location hours
                              Session("RCM273_PickupAfterHourFeeID")=s_pl("AfterHourFeeID")
                         else
                          '--if allow after hour booking, reset office time to After hour start and End time
                           if PmAfterHoursEnd="00:00" then
                                    PmAfterHoursEnd="23:55"
                              end if
                               PickupOfficeOpeningDateTime=(RequestPickDate&" "&AmAfterHoursStart)
                           PickupOfficeClosingDateTime=(RequestPickDate&" "&PmAfterHoursEnd)
                            ' if  RequestPickDateTime >CDate(PickupOfficeOpeningDateTime)  or RequestPickDateTime < CDate(PickupOfficeClosingDateTime)  then
                              if  CDate(RequestPickDateTime) <CDate(PickupOfficeOpeningDateTime)  or RequestPickDateTime > CDate(PickupOfficeClosingDateTime)  then
                                       'Session("RCM273_ErrorMesage")="Pickup Location "&Session("RCM273_PickupLocation")&"   will not take bookings outside office hours ("&WeekDayName(WeekDay(RequestPickDate))&" "&AmAfterHoursStart&" - "&PmAfterHoursEnd&").<Br>Please contact the bookings office directly."
                                   Session("RCM273_ErrorMesage")="Pickup Location "&Session("RCM273_PickupLocation")&"   will not take bookings outside office hours ("&WeekdayName(WeekDay(RequestPickDate))&" "&PickupOfficeOpeningTime&" - "&PickupOfficeClosingTime&").<Br>Please contact the bookings office directly."
                                   Response.Redirect "webstep2.asp"
                           else
                                    '--charge after hour fee if pickup is later then AmHurs Start or in the PM after hours period
                                     Session("RCM273_PickupAfterHourFeeID")=s_pl("AfterHourFeeID")
                                 'Response.Write("PickupAfterHourfee--")

                        end if
                  end if
               end if
            end if

        end if
        s_pl.close
        set s_pl=nothing


                        '---get the dropoff location informations-------------------------
      '--1. if allow unattended dropoff, do not need check if is a after hour booking
      '--2. if not allow unattended dropoff, chech if is a after hour booking

   Session("RCM273_DropoffAfterHourFeeID")=0
      Session("RCM273_UnattendedDropoffFeesID")=0
      Session("RCM273_DropoffLocation")=""
      Session("RCM273_DropoffLocationCode")=""
      Set l_s = webConn.Execute("SELECT * FROM QLocationState where (ID) = '"&CINT(Session("RCM273_DropoffLocationID"))&"' ")
      if Not l_s.EOF then
         Session("RCM273_DropoffLocation")=l_s("Location")
         Session("RCM273_DropoffLocationCode")=l_s("CityCode")
         DropoffAfterHourBooking= l_s("AfterHourBooking")
         '--dropoff location Office openingDateTime,ClosingDateTime
            DropoffOfficeOpeningTime =l_s("OfficeOpeningTime")
            DropoffOfficeClosingTime=l_s("OfficeClosingTime")
            DropoffOfficeOpeningDateTime=(RequestDropDate&" "&l_s("OfficeOpeningTime"))
            DropoffOfficeClosingDateTime=(RequestDropDate&" "&l_s("OfficeClosingTime"))

           '--20/Jun/2008, Opening and closing time should use the OfficeOpeningTime record for each weekday (not the location record)
               SQL="SELECT * from OfficeOpeningTime where ot_LocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"' AND ot_WeekDay='"&WeekDay(RequestDropDate)&"'  "
            SET s_oo=webConn.Execute(SQL)
            'Response.Write(SQL)
            '--set the AfterHour time same as office hour, if no record be set up in tale OfficeOpeningTime then would not have a null data
            DropoffOfficeAfterHourStart=DropoffOfficeClosingDateTime
            DropoffOfficeAfterHourEnd=DropoffOfficeClosingDateTime
               SetDropOFfLocationHour=False
            if Not s_oo.EOF then
                     DropoffOfficeOpeningDateTime=(RequestDropDate&" "&s_oo("ot_OfficeOpeningTime"))
                     DropoffOfficeClosingDateTime=(RequestDropDate&" "&s_oo("ot_OfficeClosingTime"))

                     DropoffOfficeAfterHourStart=(RequestDropDate&" "&s_oo("ot_AmAfterHoursStart"))
                     DropoffOfficeAfterHourEnd=(RequestDropDate&" "&s_oo("ot_PmAfterHoursEnd"))
                     '--over write the office TIME for the errormessage to use
                     DropoffOfficeOpeningTime=s_oo("ot_OfficeOpeningTime")
                     DropoffOfficeClosingTime=s_oo("ot_OfficeClosingTime")
                     SetDropOFfLocationHour=True
         end if
      s_oo.close
         set s_oo=nothing
         Session("RCM273_DropoffOfficeOpeningTime")=DropoffOfficeOpeningTime
         Session("RCM273_DropoffOfficeClosingTime")=DropoffOfficeClosingTime
         '---if is  out office hours---------
      if  ( RequestDropDateTime > CDate(DropoffOfficeClosingDateTime) or RequestDropDateTime <CDate(DropoffOfficeOpeningDateTime) ) then
                  if l_s("UnattendedDropoffs") = "True"  then '--check if allow Unattended Dropoff
               '--check if charge fees for unattended dropoff
               'Session("RCM273_UnattendedDropoffFeesID")=l_s("UnattendedDropoffFeesID")
               Session("RCM273_DropoffAfterHourFeeID") =  l_s("UnattendedDropoffFeesID")
           ' Response.Write("<font color='red'>--UnattendedDropoff--</font>")
                  else  '--if not allow Unattended Dropoff
            if DropoffAfterHourBooking <> "True" then
               '--if not allow after hour booking -------------
               Session("RCM273_ErrorMesage")="Dropoff Location "&Session("RCM273_DropoffLocation")&"  will not take bookings outside office hours ("&WeekDayName(WeekDay(RequestDropDate))&"  "&DropoffOfficeOpeningTime&" - "&DropoffOfficeClosingTime&").<Br>Please contact the bookings office directly."
                 Response.Redirect "webstep2.asp"
               else    '--- if allow afterhours booking

                     if SetDropOFfLocationHour=False then '--if no location hours record set
                        Session("RCM273_DropoffAfterHourFeeID")=l_s("AfterHourFeeID")

                       else
                             if  ( RequestDropDateTime > CDate(DropoffOfficeAfterHourEnd) or RequestDropDateTime <CDate(DropoffOfficeAfterHourStart) ) then
                              Session("RCM273_ErrorMesage")="Dropoff Location "&Session("RCM273_DropoffLocation")&"  will not take bookings outside office hours ("&WeekDayName(WeekDay(RequestDropDate))&" "&DropoffOfficeOpeningTime&" - "&DropoffOfficeClosingTime&").<Br>Please contact the bookings office directly."
                                Response.Redirect "webstep2.asp"
                     else '---if is not out the after hours time, check if charge after hour booking
                              if  ( RequestDropDateTime > CDate(DropoffOfficeClosingDateTime) or RequestDropDateTime <CDate(DropoffOfficeOpeningDateTime) ) then
                                    Session("RCM273_DropoffAfterHourFeeID")=l_s("AfterHourFeeID")
                     'Response.Write("afterHourFee--")
                        end if
                    end if
                end if
              end if
         end if
         end if
        end if
         l_s.close
         set l_s=nothing



         if  CStr(Session("RCM273_MinimunAge"))<>"0" and  CInt(Session("RCM273_driverage"))< CInt(Session("RCM273_MinimunAge"))  then
                 Session("RCM273_ErrorMesage")="Vehicles are not available for hire to drivers under "&Session("RCM273_MinimunAge")&" years of age"
                  Response.Redirect "webstep2.asp"

       end if

     Session("RCM273_NoOfDays")=1

   '---calculat the total min - Grace period  (mins)
          Totalmin=0
         Session("RCM273_Graceperiod")=0
      Graceperiod=0
          Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='GP'  ")
    If not RG.EOF then
            Session("RCM273_Graceperiod")=RG("syValue") '-720 for 12 Hour rate
            Graceperiod=RG("syValue")
   END IF
         RG.CLOSE
   SET RG=NOTHING
   '---May 2013, new code, if less then a day do not take the system grace period out
   if RequestDropDate= RequestPickDate then
      Graceperiod=0
      Session("RCM273_Graceperiod")=0
      end if
     ' SQL=" SELECT  top 1 Rate12Hour  FROM CarSize  "
      SQL=" SELECT   Rate12Hour  FROM CategoryType  WHERE CategoryType.ID='"&CINT(Session("RCM273_CategoryTypeID"))&"' "

      Set s_r=webConn.Execute(SQL)
      if NOT s_r.EOF then
              if s_r("Rate12Hour")=True then
                Session("RCM273_Graceperiod")= Graceperiod -720
               end if

      end if
      s_r.close
      set s_r=nothing

   Session("RCM273_TotalRentalDays")=0
   Totalmin= Round((RequestDropDateTime - RequestPickDateTime)*1440-Session("RCM273_Graceperiod"),2)
        Session("RCM273_TotalRentalDays")=Totalmin/1440
        Session("RCM273_TotalRentalDays24Hour")=Totalmin/1440
        Session("RCM273_KmsTotaldays")=Round(Totalmin/1440,2)  '--used in step3 calculate Kms
        actureTotalDays=Round(Totalmin/1440,2) '---for find out the extra hours in which season

        if  Session("RCM273_TotalRentalDays") > Fix(Session("RCM273_TotalRentalDays")) then
            Session("RCM273_TotalRentalDays")=Fix(Session("RCM273_TotalRentalDays"))+1
         end if

         if Session("RCM273_TotalRentalDays")<1 then
             Session("RCM273_TotalRentalDays")=1
      end if
      '---daily Extrafees always use   Session("RCM273_TotalRentalDays") (fixed number) ----
       TotalRentalDays=Session("RCM273_TotalRentalDays")
         TotalDays= Session("RCM273_TotalRentalDays")

        NoticeRequiredDate=(Now+NoticeRequired)

         '------- check No. of Notice Required for online booking --------------------
        if RequestPickDateTime =< (NoticeRequiredDate) then
            Session("RCM273_ErrorMesage")="  Reservation requests made for "&Session("RCM273_PickupLocation")&" must be made "&NoticeRequired&" days or<br>  more prior to vehicle pick up."
              Response.Redirect "webstep2.asp"

   end if


      '---if relocating check if avaliable 1st, then the MinBookingDay --------------------
                  SQL = "SELECT  top 1 MinBookingDay,WebAvaliable,PickupDateTo,PickupDateFrom FROM WebRelocationFees where  "
                SQL=SQL&"  (PickupLocationID) ='"&CINT(Session("RCM273_PickupLocationID"))&"' and (DropoffLocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"' "
                SQL=SQL&" AND ((PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"') "
                 SQL=SQL&" or  Year(PickupDateFrom)=2100 ) "
                SQL=SQL&" AND CarSizeID=0 "
                SQL=SQL&" order by PickupDateFrom "
                'Response.Write(SQL)
                  Set s_re = webConn.Execute(SQL)
                   if Not s_re.EOF then
                    if s_re("WebAvaliable")=True then
                  if s_re("MinBookingDay") >1 and s_re("MinBookingDay") > TotalRentalDays then
                     if Year(s_re("PickupDateFrom"))<>2100 then
                     Session("RCM273_ErrorMesage")="The minimum Rental period between "&Session("RCM273_PickupLocation")&" and "&Session("RCM273_DropoffLocation")&" <br>for an Internet booking is "&s_re("MinBookingDay")&" days from "&Day(s_re("PickupDateFrom"))&"/"&Left(MonthName(Month(s_re("PickupDateFrom"))),3)&"/"&Year(s_re("PickupDateFrom"))&" to "&Day(s_re("PickupDateTo"))&"/"&Left(MonthName(Month(s_re("PickupDateTo"))),3)&"/"&Year(s_re("PickupDateTo"))&". "
                               Response.Redirect "webstep2.asp"
                                 else
                                 Session("RCM273_ErrorMesage")="The minimum Rental period between "&Session("RCM273_PickupLocation")&" and "&Session("RCM273_DropoffLocation")&" <br>for an Internet booking is "&s_re("MinBookingDay")&" days. "
                                Response.Redirect "webstep2.asp"
                                end if

                    end if
              else
                           Session("RCM273_ErrorMesage")="The one way booking between "&Session("RCM273_PickupLocation")&" and "&Session("RCM273_DropoffLocation")&" is not available "
                             Response.Redirect "webstep2.asp"
                     end if
            end if
            s_re.close
            set s_re=nothing



   '--Check if location will take bookings outside office hours, Unattended Dropoffs
   if Session("RCM273_PickupAfterHourBooking") = "False" then
                    if  RequestPickDateTime < CDate(PickupOfficeOpeningDateTime)  or RequestPickDateTime > CDate(PickupOfficeClosingDateTime)  then
                        Session("RCM273_ErrorMesage")="Pickup Location "&Session("RCM273_PickupLocation")&"   will not take bookings outside office hours ("&Session("RCM273_PickupOfficeOpeningTime")&" - "&Session("RCM273_PickupOfficeClosingTime")&").<Br>Please contact the bookings office directly."
                          Response.Redirect "webstep2.asp"

            end if
       end if
       if Session("RCM273_DropoffAfterHourBooking") = "False" then
                 if  UnattendedDropoffs = "False" and ( RequestDropDateTime > CDate(DropoffOfficeClosingDateTime) or RequestDropDateTime <CDate(DropoffOfficeOpeningDateTime) ) then
                        Session("RCM273_ErrorMesage")="Dropoff Location "&Session("RCM273_DropoffLocation")&"  will not take bookings outside office hours ("&Session("RCM273_DropoffOfficeOpeningTime")&" - "&Session("RCM273_DropoffOfficeClosingTime")&").<Br>Please contact the bookings office directly."
                       Response.Redirect "webstep2.asp"

                 end if
       end if





       '------the form-----------------------------------------------



        '-----Start Hourly Rate ------------
        '----1st need find the extra hours in which season, (use the extra hours as a day to find the season)
        '--- need a acture No of Days (actureTotalDays) and actureNoOfDaysEachSeason to find out the extra hours in which season
        '---if there are hourly rate set up then use it for calculattion
       '----if the total cost for the no of hours > = one day rate, then do not calculate hourly rate

        theExtraHour=0
        HourRate=0
        if Session("RCM273_RequestPickTime")<> Session("RCM273_RequestDropTime") then
            theExtraHour=Round(((1440*(CDate(Session("RCM273_RequestDropDate")&" "& Session("RCM273_RequestDropTime")) - CDate(Session("RCM273_RequestDropDate")&" "& Session("RCM273_RequestPickTime")))-Session("RCM273_Graceperiod")) /60),2)
                if  theExtraHour > Fix(theExtraHour)  then
                  theExtraHour=Fix(theExtraHour)+1
            end if
                theExtraHour=Round(theExtraHour,2)
                'Response.Write(theExtraHour)
            'Response.Write("<br>")
            SQL="SELECT * from CarRateHourly WHERE (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"'  and HourFrom<"&CINT(theExtraHour)&" and HourTo>="&CINT(theExtraHour)&" "
                'Response.Write(SQL)
                SET s_hr=webConn.Execute(SQL)
               if Not s_hr.EOF then
                  HourRate=s_hr("HourlyRate")
                  if   theExtraHour*s_hr("HourlyRate")>=100  then
                  '----if the total cost for the no of hours > = one day rate, then do not calculate hourly rate
                        '--reset to 0
                        theExtraHour=0
                        HourRate=0
                  end if

                end if
                s_hr.close
                set s_hr=nothing
        end  if
        '-----END Hourly Rate ------------


    '---1. get the no of days for each season----
    '---2. For each car type, get each season rate, then calculate the total cost for each car type---


   TotalDays = Session("RCM273_TotalRentalDays")
          SQL="SELECT * from Season where LocationID="&CINT(Session("RCM273_PickupLocationID"))&" and notActive=0 and (Season='Default' or  (EndDate>='"&RequestPickDate&"' and StartDate<='"&RequestDropDate&"')  )  order by StartDate DESC  "
          'Response.Write(SQL)
          SET s_m=webConn.Execute(SQL)

    HeightSeasonDays=0
    TotalCost=0
         i=0

             DO WHILE NOT s_m.EOF
             '_____________ calculate the no of days for each season
               NoOfDaysEachSeason=0
               actureNoOfDaysEachSeason=0 ' to find out extra Hours in which season
                  '--for special seasons, the calculate no of days
               if s_m("Season") <>"Default" and Session("RCM273_TotalRentalDays")-HeightSeasonDays>0 then          ' for not default season
                 '--need use Season start and End Datetime for calculating the No of days (with PickupDateTime, dorpoffDateTime)
                  SeasonStartTime="00:00:00"
            SeasonStartingDateTime=CDate(s_m("StartDate")&" "& SeasonStartTime)
                 SeasonEndingDateTime=s_m("EndDate")
                        '---if the time diff >12 hours then
                       'IF  (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)=>12  and Round((RequestDropDateTime-SeasonStartingDateTime)*1440,2)/60>12 then
                      IF  (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)=>12  and Round((RequestDropDateTime-SeasonStartingDateTime)*1440,2)/60>0 then
                           ' _____________ calculate the no of days in each season
                               'if Round((SeasonEndingDateTime - RequestDropDateTime )*1440,2)/60=>12 then    '4--Dropoffdate > seasonEnd
                          if Round((SeasonEndingDateTime - RequestDropDateTime )*1440,2)/60>=0  then
                               '--the season strat time should be same as Pickup time
                                 SStartingDateTime=CDate(s_m("StartDate")&" "& Session("RCM273_RequestPickTime"))
                                 themin= Round((RequestDropDateTime - SStartingDateTime)*1440-Session("RCM273_Graceperiod"),2)
                                 if  themin<0 then
                                    themin=0
                                 end if
                                 NoOfDaysEachSeason=themin/1440
                                 actureNoOfDaysEachSeason=themin/1440
                                 if  NoOfDaysEachSeason > Fix(NoOfDaysEachSeason) then
                                       NoOfDaysEachSeason=Fix(NoOfDaysEachSeason)+1
                                 end if
                                 'Response.Write("<p>2 "&s_m("Season")&" "&NoOfDaysEachSeason&" ")
                          else
                                 NoOfDaysEachSeason=DateDiff("d", s_m("StartDate"),s_m("EndDate"))
                                 actureNoOfDaysEachSeason=DateDiff("d", s_m("StartDate"),s_m("EndDate"))
                                 'Response.Write("<p>6 "&s_m("Season")&" "&NoOfDaysEachSeason&" ")
                           end if
                         'ELSEIF   (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)<12  and Round((SeasonEndingDateTime - RequestPickDateTime )*1440,2)/60=>12 then
                        ELSEIF   (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)<12  and Round((SeasonEndingDateTime - RequestPickDateTime )*1440,2)/60=>0 then
                               'if Round((SeasonEndingDateTime - RequestDropDateTime )*1440,2)/60>12 then    '4--Dropoffdate > seasonEnd
                            if Round((SeasonEndingDateTime - RequestDropDateTime )*1440,2)/60>=0  then
                                  themin= Round((RequestDropDateTime - RequestPickDateTime)*1440-Session("RCM273_Graceperiod"),2)
                   if  themin<0 then
                           themin=0
                   end if
                   NoOfDaysEachSeason=themin/1440
                   actureNoOfDaysEachSeason=themin/1440
                   if  NoOfDaysEachSeason > Fix(NoOfDaysEachSeason) and (HeightSeasonDays+NoOfDaysEachSeason)<TotalDays  then
                        NoOfDaysEachSeason=Fix(NoOfDaysEachSeason)+1
                      else
                   NoOfDaysEachSeason=Fix(NoOfDaysEachSeason)
                   end if
                                 'Response.Write("<p>4 "&s_m("Season")&" "&NoOfDaysEachSeason&" ---"&theDays&"")
                         else
                             '--the season end time should be same as Return Time
                                 SEndingDateTime=CDate(s_m("EndDate")&" "& Session("RCM273_RequestPickTime"))
                         themin= Round((SEndingDateTime - RequestPickDateTime)*1440,2)
                  if  themin<0 then
                           themin=0
                   end if
                  NoOfDaysEachSeason=0
                   NoOfDaysEachSeason=themin/1440
                   actureNoOfDaysEachSeason=themin/1440
                   if  NoOfDaysEachSeason > Fix(NoOfDaysEachSeason) then
                        NoOfDaysEachSeason=Fix(NoOfDaysEachSeason)
                   end if
                   SeasonLong  = (s_m("EndDate")-s_m("StartDate"))
                   if NoOfDaysEachSeason >SeasonLong then
                        NoOfDaysEachSeason  = SeasonLong
                   end if

                     '---Dec 2013, added, the following code, when calculate the 2nd season, add 1 day for the condition below
                        if  Session("RCM273_RequestPickDate")>=s_m("StartDate") and Session("RCM273_RequestDropDate")>s_m("EndDate")  and (HeightSeasonDays+NoOfDaysEachSeason)<TotalDays then
                                             NoOfDaysEachSeason=Fix(NoOfDaysEachSeason)+1
                        end if
                   if (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)<0 and TotalDays=1 then
                        NoOfDaysEachSeason=NoOfDaysEachSeason-1
                   end if
                   if TotalDays=1 then    '---added 29/Nov/06,had 0 rate if the booking day is end of season---
                         NoOfDaysEachSeason=1
                   end if
                   'Response.Write("<p>3 "&s_m("Season")&" "&NoOfDaysEachSeason&" ")
               end if
            END IF
            HeightSeasonDays=HeightSeasonDays+NoOfDaysEachSeason  ' add the height seasons days up
         elseif s_m("Season") ="Default" and Session("RCM273_TotalRentalDays")-HeightSeasonDays>0  then
                '___________the default Season, days = total days -provices seasons days
                NoOfDaysEachSeason=Session("RCM273_TotalRentalDays")-HeightSeasonDays
                actureNoOfDaysEachSeason=actureTotalDays-HeightSeasonDays
         end if
            'Response.Write(NoOfDaysEachSeason)
             'Response.Write(Round(actureNoOfDaysEachSeason,2))
             'Response.Write("<br>")
             actureNoOfDaysEachSeason=Round(actureNoOfDaysEachSeason,2)
            '   Response.Write theExtraHour
            'Response.Write("<br>")
             SeasonCost=0
             if NoOfDaysEachSeason>0 then
                  i=i+1
                  if s_m("Season")="Default" then
                        Season="Rates"
                        Session("RCM273_Season"&i&"")="Rates"

                  else
                        Season=s_m("Season")
                        Session("RCM273_Season"&i&"")=s_m("Season")
                  end if

                  '--pass the seasonID as array
                  Session("RCM273_SeasonID"&i&"")=s_m("ID")
                  ' Session("RCM273_NoOfDaysEachSeason"&i&"")=NoOfDaysEachSeason

                    GetEachSeasonRateStructureID
                    '---the code has to be here after get the RateStructureID-------------------
                    '---need find the last season
                     '----if there are extra Hours and hour rate <>0 , reset the noOfDays for each season

                     '   '-- add the extra hour to the  1st season number of days, Mar 2012 made the code changes
                        if  i=1 and theExtraHour<>0 and  HourRate<>0 then
                            NoOfDaysEachSeason=NoOfDaysEachSeason-1+ theExtraHour*HourRate/100
                          ' Response.Write("<br>theExtraHour=")
                           'Response.Write(theExtraHour)
                          ' Response.Write("-----HourRate=")
                          ' Response.Write(HourRate)
                          ' Response.Write("---noOfDays=")
                          ' Response.Write(Round(NoOfDaysEachSeason,2))
                          '  Response.Write s_m("Season")

                        end if

                     Session("RCM273_NoOfDaysEachSeason"&i&"")=NoOfDaysEachSeason
                     if Session("RCM273_TotalRentalDays")=1 then
                        Session("RCM273_NoOfDaysEachSeason"&i&"")=1
                        NoOfDaysEachSeason=1
                     end if
                     '--pass all following to Step3.asp for get the rate again

            end if
            s_m.MoveNext
            Loop
            s_m.Close
            Response.Write("<input type=""hidden"" name=""SeasonCount"" value='"&i&"' />")                  'Count
          SeasonCount=i

          '---get car type, for each car type, get each season rate, then calculate total cost for each type

          findTheRate2

          Response.Write("<tr><td style=""height:2px""  bgcolor='"&Session("RCM273_CompanyColour")&"'  ></td></tr>")
          Response.Write("<tr><td style=""height:5px""   bgcolor='"&Session("RCM273_CompanyColour")&"'  ></td></tr>")

          Response.Write("</table>")
          Response.Write("</td></tr>")
          Response.Write("</table>")

          webConn.CLOSE
          SET webConn=nothing
   end if 'if Session("RCM273_ErrorMesage")<>"" then
%>
  
 </div>
<!-- #include file="include_footer.asp" -->

</body>
</html>
