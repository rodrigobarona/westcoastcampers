<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->

<!-- #include file="include_meta.asp" -->

</head>
<body  class="webstep3" >

<!-- #include file="include_header.asp" -->

<%
FUNCTION tidyup(thisString)
    thisString=LTrim(RTrim(Replace(thisString, "'" , "")))
      thisString=Replace(thisString, "�" , "")
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



      if Request.QueryString("PickupLocationID")="" then
            Response.Redirect "webstep2.asp"
      end if



   if Request.QueryString("CarSizeID")="" then
          Session("RCM273_ErrorMesage")="Please select the vehicle."
                  Response.Redirect "webstep2.asp"
     end if
      if Session("RCM273_CategoryTypeID")="" then
            Response.Redirect "webstep2.asp"
     end if
     if Session("RCM273_driverage")="" then
         Response.Redirect "webstep2.asp"
     end if



      Session("RCM273_Mileage")="Miles"


     CompanyCode="PtWestCoastCampers273"
           RCMURL="https://secure2.rentalcarmanager.com"

           Dim webConn
           Set webConn = Server.CreateObject("ADODB.Connection")
            DatabaseFile="PtWestCoastCampers273"
            webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"



   Session("RCM273_CustomerID")=0 '--used for existing customer
   Session("RCM273_BookingBufferNo")=0
   '---reset the session in the begging of the page, when user open more than one windows to compare rates, all the session will be maxed up, and cause problems
   Session("RCM273_CarSizeID")=Left(Request.QueryString("CarSizeID"),3)
   Session("RCM273_categoryStatus") = Left(Request.QueryString("categoryStatus"),1)
   Session("RCM273_PickupLocationID") = Left(Request.QueryString("PickupLocationID"),3)
   Session("RCM273_DropoffLocationID")= Left(Request.QueryString("DLocationID"),3)


   Session("RCM273_RequestPickDate")=Day(Request.QueryString("PickDateTime"))&"/"&MonthName(Month(Request.QueryString("PickDateTime")))&"/"&Year(Request.QueryString("PickDateTime"))
   Session("RCM273_RequestDropDate")=Day(Request.QueryString("DropoffDateTime"))&"/"&MonthName(Month(Request.QueryString("DropoffDateTime")))&"/"&Year(Request.QueryString("DropoffDateTime"))
   Session("RCM273_RequestPickTime")=Left(Request.QueryString("PickupTime"),5)
    Session("RCM273_RequestDropTime")=Left(Request.QueryString("DropoffTime"),5)

 'Session("RCM273_RequestPickTime")= FormatDateTime(Request.QueryString("PickDateTime"),4)
 '  Session("RCM273_RequestDropTime")= FormatDateTime(Request.QueryString("DropoffDateTime"),4)

   Session("RCM273_RequestPickDateTime") = Session("RCM273_RequestPickDate")&" "&Session("RCM273_RequestPickTime")
   Session("RCM273_RequestDropDateTime") = Session("RCM273_RequestDropDate")&" "&Session("RCM273_RequestDropTime")

   'Session("RCM273_TotalRentalDays")= Left(CDbl(Request.QueryString("TotalRentalDays")),8)'---cause issues with the left
 Session("RCM273_TotalRentalDays")= CDbl(Request.QueryString("TotalRentalDays"))


   Set l_s = webConn.Execute("select Location, TimeDifferenceGMT, GETUTCDATE() as GMTTime FROM Location WHERE ID='"&CInt(Session("RCM273_PickupLocationID"))&"' ")
   if Not l_s.EOF then
          Session("RCM273_PickupLocation")=l_s("Location")
           LocalTime =l_s("GMTTime")+(l_s("TimeDifferenceGMT")/24)
           TheDate  = Day(LocalTime)&"/"&MonthName(Month(LocalTime),3)&"/"&Year(LocalTime)
           TheTime  = FormatDateTime(LocalTime,4)
           Session("RCM273_LocalTime")=TheDate&" "&TheTime
           'Session("RCM273_LocalTime") = CDate(TheDate&" "&TheTime)'--need convet to CDate, otherwise cause problem on the booking sheet

   end if
   l_s.Close
   Set l_s=nothing
    Set l_s = webConn.Execute("select Location FROM Location WHERE ID='"&CInt(Session("RCM273_DropoffLocationID"))&"' ")
   if Not l_s.EOF then
          Session("RCM273_DropoffLocation")=l_s("Location")
   end if
   l_s.Close
   Set l_s=nothing



    '---pass the booking details to next page, do not use sessions
   Response.Write("<form method='post' action='webstep4.asp?dir=Rate&categoryStatus="&Session("RCM273_categoryStatus") &"'  name='Rate'  id='Rate'  >")
   Response.Write("<input type='hidden' name='categoryStatus' size='5' value='"&Session("RCM273_categoryStatus") &"' />")
   Response.Write("<input type='hidden' name='CarSizeID' size='5' value='"&CInt(Session("RCM273_CarSizeID"))&"' />")
   Response.Write("<input type='hidden' name='PickupLocationID' size='5' value='"&CInt(Session("RCM273_PickupLocationID"))&"' />")
   Response.Write("<input type='hidden' name='DropoffLocationID' size='5' value='"&CInt(Session("RCM273_DropoffLocationID"))&"' />")
   'Response.Write("<input type='hidden' name='PickDateTime' size='15' value='"&Session("RCM273_RequestPickDateTime")&"' />")
   'Response.Write("<input type='hidden' name='DropoffDateTime' size='15' value='"&Session("RCM273_RequestDropDateTime")&"' />")
   Response.Write("<input type='hidden' name='PickupDate' size='15' value='"&Session("RCM273_RequestPickDate")&"' />")
   Response.Write("<input type='hidden' name='DropoffDate' size='15' value='"&Session("RCM273_RequestDropDate")&"' />")
   Response.Write("<input type='hidden' name='PickupTime' size='15' value='"&Session("RCM273_RequestPickTime")&"' />")
   Response.Write("<input type='hidden' name='DropoffTime' size='15' value='"&Session("RCM273_RequestDropTime")&"' />")
   Response.Write("<input type='hidden' name='TotalRentalDays' size='15' value='"&Session("RCM273_TotalRentalDays")&"' />")








SUB RelactionFees
 '----Relocation fee-------------
         Session("RCM273_RelocationFee")=0

                '--1. check Relocation record (with caterory, date range)
               '--1. check Relocation record (with caterory, date range)
         SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
         SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID and (CarSizeID)='"&CInt(Session("RCM273_CarSizeID"))&"'  and Mandatory=0 "
         SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&Session("RCM273_DropoffLocationID")&"'  "
         SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
                  '--next line of code will return the max Minbookingday if there are 2 records entered for same conditin
         SQL=SQL&"  AND MinBookingDay<="&Session("RCM273_TotalRentalDays")&" order by MinBookingDay desc "
         'Response.Write(SQL)
          Set s_o=webConn.Execute(SQL)
          if  s_o.EOF THEN
                  s_o.close
                  Set s_o=Nothing
                  '--2. if no vehicle category Relocation fee found, check Relocation record (with  date range only)
                  SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                  SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and Mandatory=0  "
                  SQL=SQL&" AND CarSizeID=0 "
                  SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&Session("RCM273_DropoffLocationID")&"'  "
                  SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
                  SQL=SQL&"  AND MinBookingDay<="&Session("RCM273_TotalRentalDays")&" order by MinBookingDay desc "
                  'Response.Write(SQL)
                   Set s_o=webConn.Execute(SQL)
                   if  s_o.EOF THEN
                        s_o.close
                        Set s_o=Nothing
                        '--3. if no vehicle category Relocation fee found, check Relocation record (with  catgory, no date ragne)
                        SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                        SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and (CarSizeID)='"&CInt(Session("RCM273_CarSizeID"))&"' "
                        SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&Session("RCM273_DropoffLocationID")&"'  "
                        SQL=SQL&"  AND  Year(PickupDateFrom)=2100  and Mandatory=0 "
                        SQL=SQL&"  AND MinBookingDay<="&Session("RCM273_TotalRentalDays")&" order by MinBookingDay desc "
                        'Response.Write(SQL)
                        Set s_o=webConn.Execute(SQL)
                        if  s_o.EOF THEN
                              s_o.close
                              Set s_o=Nothing
                                    '--4. if no vehicle category Relocation fee found, check Relocation record (with no catgory, no date ragne)
                              SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                              SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID   and Mandatory=0 "
                              SQL=SQL&" AND CarSizeID=0 "
                              SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&Session("RCM273_DropoffLocationID")&"'  "
                              SQL=SQL&"  AND  Year(PickupDateFrom)=2100 "
                              SQL=SQL&"  AND MinBookingDay<="&Session("RCM273_TotalRentalDays")&" order by MinBookingDay desc "
                              'Response.Write(SQL)
                              Set s_o=webConn.Execute(SQL)
                        end if
                  end if

         end if

         if NOT s_o.EOF THEN
               DO WHILE NOT s_o.EOF

                   if  s_o("DaysNocharge")=0 or (s_o("DaysNocharge")>0 and Session("RCM273_TotalRentalDays")<s_o("DaysNocharge"))  then

                     Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_o("Fees")
                     Response.Write("<div class='row'><div class='col-xs-7'><p class='text-left'><strong>"&s_o("Name")&"</strong></p></div><div class='col-xs-5'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_o("Fees"),2)&"</p></div></div>")


                 end if
                s_o.MoveNext
                Loop
         end if
         s_o.CLOSE
         SET s_o=NOTHING

             '---Aug 2014, new WebExtralAdditionalFees, added for wicked AU, allow to add additional fee for date range
            SQL="SELECT DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebExtralAdditionalFees,ExtraFees "
            SQL=SQL&" WHERE  PickuplocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' "
            SQL=SQL&"  AND (DropofflocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"' or DropofflocationID=0) "
              ' SQL=SQL&" and (CarSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or CarSizeID=0)   "
             SQL=SQL&" and (CarSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or WebExtralAdditionalFees.CategoryTypeID="&Session("RCM273_CategoryTypeID")&" or  (CarSizeID=0 and WebExtralAdditionalFees.CategoryTypeID=0)   ) "
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
                     Response.Write("<tr><td class=""text"" align=""left""  nowrap=""nowrap"">"&s_o("Name")&" </td><td class=""text"" align=""right"" >"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_o("Fees"),2)&"</td></tr>")
            s_o.MoveNext
            Loop
            s_o.CLOSE
            SET s_o=NOTHING

'Response.Write("<Br>")
'Response.Write(SQL)

 '-----Pickup Location After hours and befores fee, check if the pickup time is between the office time-------------
        Session("RCM273_AfterHoursFee")=0
        Session("RCM273_PickupAfterHoursFee")=0
        AfterHoursFeeGST=0
        AfterHoursFeeStampDuty=0
        if Session("RCM273_PickupAfterHourFeeID")<>0  then 'check the time
               AfterHoursFee=0
               Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&Session("RCM273_PickupAfterHourFeeID")&"' ")
               if s_st("Fees")  <>0 then
                       '---Pickup Location After Hour fees   'do not convert to CDate
                     PickupOpeningTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeOpeningTime"))
                     PickupClosingTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeClosingTime"))
                     if (Session("RCM273_RequestPickDateTime") < PickupOpeningTime) or (Session("RCM273_RequestPickDateTime") > PickupClosingTime) THEN
                         Response.Write("<div class='row'><div class='col-xs-7'><p class='text-left'><strong>"&s_st("Name")&"</strong></p></div><div class='col-xs-5'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_st("Fees"),2)&"</p></div></div>")
                           Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_st("Fees")
                     end if
            end if
            s_st.close
            SET s_st=nothing

      end if

    '------------Dropoff Location After Hour Fees ------------------
    Session("RCM273_DropoffAfterHoursFee")=0
    if Session("RCM273_DropoffAfterHourFeeID")<>0  then
                  Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&Session("RCM273_DropoffAfterHourFeeID")&"' ")
                  'response.write s_st.source
                  'response.Write s_st("Fees")
                 if s_st("Fees")  <>0 then
                         DropoffOpeningTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeOpeningTime"))
                        DropoffClosingTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeClosingTime"))
                        if (Session("RCM273_RequestDropDateTime") < DropoffOpeningTime) or (Session("RCM273_RequestDropDateTime") > DropoffClosingTime) THEN
                               Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_st("Fees")
                                Response.Write("<div class='row'><div class='col-xs-7'><p class='text-left'><strong>"&s_st("Name")&"</strong></p></div><div class='col-xs-5'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_st("Fees"),2)&"</p></div></div>")

                        end if
            end if

            s_st.close
            SET s_st=nothing
         end if



END SUB


SUB SelectAreaofUsed
          Set l_s = webConn.Execute("select * FROM AreaofUsed where WebItem=1 and (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or LocationID=0) order by AreaofUsed")
               'Response.Write("<tr><td colspan='4' style='height: 1px' bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>")
               if Not l_s.EOF then
                     Response.Write("<tr><td  class='text' align='left' colspan='4' >&nbsp;Area of Use <select name='AreaofUsed' >")
            While Not l_s.EOF
            if l_s("Defaulted")=True  then
            Response.Write("<option value='"&l_s("AreaofUsed")&"' selected='selected'>"&l_s("AreaofUsed")&"</option>")

               else
               Response.Write("<option value='"&l_s("AreaofUsed")&"' >"&l_s("AreaofUsed")&"</option>")
                   end if
           l_s.MoveNext
            Wend

          Response.Write("</select></td></tr>")
           'Response.Write("<tr><td colspan='4' style='height: 1px' bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>")
         end if
         l_s.Close
      Set l_s=nothing
END SUB  %>
<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface==
'== File Name :  webbookingstep3.asp==
'== Method Name : HTML Page Starts here - Allows entry and selection of extra fees, kms, and insurance options==
'== Tables Used : CarRateHourly, QSeason, Location==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->
<%

 SUB  GetEachSeasonRateStructureID
          '--check system set up if Calculate Seasonal Rates using total rental days (long hire rate)
         TotalDays= Session("RCM273_TotalRentalDays")


        longHireRate="No"
        Set RG=webConn.Execute("select * FROM SystemTable WHERE Code='LONGR'  ")
         If not RG.EOF then
            longHireRate= RG("syValue")'--TOTAL booking days
         END IF
         RG.CLOSE
         SET RG=NOTHING
         DaysForRate=NoOfDaysEachSeason
        if longHireRate="Yes" then
               'DaysForRate= TotalDays '--TOTAL booking days
                  '--in step2 the DaysForRate for 12 hour rate is not inculde 1 more day
               '--should use the same number of days to get RateStructureID in step3

               DaysForRate= Session("RCM273_TotalRentalDays24Hour")

               if Session("RCM273_TotalRentalDays24Hour")=0 then
               DaysForRate= TotalDays '--TOTAL booking days
               end if
        end if
        if  DaysForRate > Fix(DaysForRate) then
         DaysForRate=Fix(DaysForRate)+1
         end if
         'Response.Write("<br>test"&TotalDays&"</br>")

          RateName="Rate"
          Rate=0
          RateStructureID=1
          Session("RCM273_RateStructureID"&i&"")=1
         '--get rateStructureID ----------------
         '---weekend  booking

         '---only bookings <=3 days and between 12:00 Friday to 12:00 Monday (AU Coutesy Cars)
   '---check systemtable if use weekend rate
       Set RG=webConn.Execute("select * FROM SystemTable WHERE Code='WKEND'  ")
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
            SQL=SQL&" AND FromDay<="&DaysForRate&" order by FromDay desc "
            'Response.Write(SQL)
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
        Response.Write("<input type='hidden' name='RateStructureID"&i&"' size='5' value='"&RateStructureID&"' />")
       Response.Write("<input type='hidden' name='RateName"&i&"' size='5' value='"&RateName&"' />")


END SUB %>

<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface ==
'== File Name :  webbookingstep3.asp==
'== Method Name : FindTheRate2==
== Method Description :  Selects Required Rates based on Rate Structure==
'== Tables Used : WebLocationCategory, CarSize, WebRelocationFees, QCarRateDetails, Discount==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->
<%

SUB subGETHourlyRate
            Less1DayHireHourRate=0
            TotalDays4HourlyRate=Round((RequestDropDateTime - RequestPickDateTime),2) '--do not include Grace period
            'Response.Write TotalDays4HourlyRate
            if TotalDays4HourlyRate<1 then
                  CategoryID=CInt(Session("RCM273_CarSizeID"))
                  NumberOfHours= Round((RequestDropDateTime - RequestPickDateTime)*1440/60,2)
                  if NumberOfHours>Fix(NumberOfHours) then
                     NumberOfHours=Fix(NumberOfHours) + 1
                  end if
                  SQL= "SELECT Top 1 (ToHour), ID, RateName,NumberDays FROM CarRateStructureHour   "
                  SQL=SQL&" where FromHour<"&NumberOfHours&"  and ToHour>="&NumberOfHours&" order by Fromhour desc "
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
                                       SQL="SELECT * FROM QCarRateDetailsHour  where CarRateStructureID="&s_rs("ID")&" and  SeasonID="&s_m("ID")&"   and CarSizeID="&CategoryID&"  "
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

                                             SQL="SELECT * FROM QCarRateDetailsHour  where CarRateStructureID="&s_rs("ID")&" and  SeasonID="&s_m("ID")&"   and CarSizeID="&CategoryID&"  "
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
               if Less1DayHireHourRate>0 then
                         Session("RCM273_DiscountID")=0
                        Session("RCM273_DiscountRate")=0
                        Rate=StandardRate1
                        SQL=" select CarSize.*  FROM CarSize "
                        SQL=SQL&" WHERE CarSize.ID="&CInt(Session("RCM273_CarSizeID"))&" "
                        'Response.Write(SQL)
                        Set s_cs=webConn.Execute(SQL)      '----------Vehicle type-------------------
                        theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_cs("ImageName")
                        Session("RCM273_CarImageName")=s_cs("ImageName")
                         Session("RCM273_CarType")=s_cs("Size")
                        if s_cs("WebDesc")<>"" then
                        Session("RCM273_CarType")=s_cs("WebDesc")
                        end if
                     '--check location Image
                      SQL="SELECT * FROM CarSizeLocation WHERE  (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' and CategoryID="&s_cs("ID")&"  "
                     'Response.Write SQL
                     Set s_li=webConn.Execute(SQL)
                     if NOT s_li.EOF then
                           if s_li("ImageName")<>"" then
                              Session("RCM273_CarImageName")=s_li("ImageName")
                             theimage=RCMURL&"/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")
                            end if
                           if s_li("WebDesc")<>"" then
                              Session("RCM273_CarType") =s_li("WebDesc")
                           end if
                        end if
                        s_li.close
                        'Set s_li=nothing

                        'Response.Write("<tr><td align='left' class='text' valign='top'>Vehicle Type:</td>")
                        'Response.Write("<td align='left' class='text' valign='top'>")

                        'Response.Write("<table align='left' cellspacing='0' cellpadding='2'  border='0'>")
                        'Response.Write("<tr><td align='left' class='text' valign='top'  nowrap='nowrap'>"&Session("RCM273_CarType")&"</td></tr>")
                        'Response.Write("<tr><td class='text'>"&s_cs("VehicleDesc")&"</td></tr>")
                        's_cs.close
                        Set s_cs=nothing
                           i=1
                          Response.Write("<tr><td class='text' nowrap='nowrap'>")
                          Response.Write("&nbsp;"&RateName&"  @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Less1DayHireHourRate,2)&" ")
                           costEachSeason=Rate*Session("RCM273_NoOfDaysEachSeason"&i&"")
                          Response.Write("</td></tr>")
                           Response.Write("<input  type='hidden'  name='StandardRate"&i&"'  value='"&Rate&"' />")
                          Response.Write("<input  type='hidden'  name='Rate"&i&"'  value='"&Rate&"' />")
                          Response.Write("<input type='hidden' name='NoOfDaysEachSeason'  maxlength=6 size=3 value='"&NumberDays&"'>")
                           Response.Write("<input type='hidden' name='lessDayNumberOfHours'  maxlength=6 size=3 value='"&NumberOfHours&"'>")

                           'Response.Write("<input type=hidden name='RateName1' size=3 value='"&RateName&"'>")
                           'Response.Write("<input type='hidden' name='Season1'  value='"&Season&"'>")
                           'Response.Write("<input type='hidden' name='SeasonID1'  value='"&SeasonID&"'>")

                        Response.Write("</table>")
                        Response.Write("</td>")
                        Response.Write("<td align='left' valign='top'><img src='"&theimage&"'  width='150' alt='' /></td>")
                        Response.Write("</tr>")

               end if



            end if
  END SUB





SUB  findTheRate2  '---Sample Web Quotation1       '---get the rate Discount-----------

         Session("RCM273_DiscountRate")=0
            Session("RCM273_DiscountID")=0
            Session("RCM273_DiscountType")="p"
        '--Mar 2012 new  code change, WebCategoryAvailability releated , select the category from carSize Table only--
         SQL=" SELECT  CarSize.*  FROM CarSize  "
       SQL=SQL&" WHERE CarSize.ID="&CInt(Session("RCM273_CarSizeID"))&" "

         Set s_cs=webConn.Execute(SQL)
         j=0
      if NOT s_cs.EOF   then
             if CInt(Session("RCM273_MinimunAge"))>s_cs("AgeYoungestDriver") THEN
            Session("RCM273_MinimunAge")=s_cs("AgeYoungestDriver")
            end if

               'Response.Write(s_cs("Size"))
               'Response.Write(s_cs("WebDesc"))

               'Response.Write("<br>")

             '---get discount Rate
                '--- check if the rental period in the Discount Date range
                '--- then check if there is a category discount rate
                '--- if not then use the location discount rate
               Session("RCM273_DiscountRate")=0
            Session("RCM273_DiscountID")=0
            Session("RCM273_DiscountType")="p"
            BookingDate=Year(Date)&"-"&Month(Date)&"-"&Day(Date)
            SQL=" SELECT TOP 1 LocationID,DropoffLocationID, CarSizeID,  DateRange from Discount left join CampaignCode on CampaignCodeID=CampaignCode.ID  "
            SQL=SQL&" WHERE WebItems=1   "
            if Session("RCM273_PromoCode")<>"" then
            SQL=SQL&" AND CampaignCode='"&Left(Session("RCM273_PromoCode"),30)&"'  "
            else
            SQL=SQL&" AND CampaignCodeID=0  "
            end if
            SQL=SQL&" AND (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or LocationID=0 )  "
            SQL=SQL&" AND (DropoffLocationID='"&Session("RCM273_DropoffLocationID")&"' or DropoffLocationID=0 )  "
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

                                 Session("RCM273_DiscountName")=s_disCat("DiscountName")
                                 Session("RCM273_DiscountType")=s_disCat("DiscountType")
                                 Session("RCM273_DiscountRate")=s_disCat("DiscountRate")
                                 Session("RCM273_DiscountID")=s_disCat("ID")
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
            SQL=SQL&" AND CampaignCode='"&Session("RCM273_PromoCode")&"' "
            SQL=SQL&" and (df_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or df_LocationID=0) "
            SQL=SQL&" and (df_CategoryID="&s_cs("ID")&" or df_CategoryID=0) "
            SQL=SQL&" and df_NoDays<="&Session("RCM273_TotalRentalDays")&" "
            SQL=SQL&" and df_DiscoutStart<='"&Session("RCM273_RequestPickDate")&"' and df_DiscoutEnd>='"&Session("RCM273_RequestPickDate")&"'  "
       else
            SQL=" select df_extraFeeID,ExtraFees.Name,ExtraFees.Fees  FROM DiscoutFixedRate,  ExtraFees "
            SQL=SQL&" WHERE df_CampaignCodeID=0 and df_extraFeeID=ExtraFees.ID "
             SQL=SQL&" and (df_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' or df_LocationID=0) "
            SQL=SQL&" and (df_CategoryID="&s_cs("ID")&" or df_CategoryID=0) "
            SQL=SQL&" and df_NoDays<="&Session("RCM273_TotalRentalDays")&" "
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


            Response.Write("<input type=hidden name='DiscountID' value='"&Session("RCM273_DiscountID")&"'>")
            Response.Write("<input type=hidden  name='DiscountRate' value='"&Session("RCM273_DiscountRate")&"'>")
            Response.Write("<input type=hidden  name='DiscountType' value='"&Session("RCM273_DiscountType")&"'>")
            '--for each season, get the rate
                  j=j+1

               '----------Vehicle type-------------------
                 theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_cs("ImageName")
                   Session("RCM273_CarImageName")=s_cs("ImageName")
                Session("RCM273_CarType")=s_cs("Size")
                        if s_cs("WebDesc")<>"" then
                        Session("RCM273_CarType")=s_cs("WebDesc")
                        end if

                         '--check location Image
                      SQL="SELECT * FROM CarSizeLocation WHERE  (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' and CategoryID="&s_cs("ID")&"  "
                     'Response.Write SQL
                     Set s_li=webConn.Execute(SQL)
                     if NOT s_li.EOF then
                           if s_li("ImageName")<>"" then
                              Session("RCM273_CarImageName")=s_li("ImageName")
                             theimage=RCMURL&"/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")
                            end if
                           if s_li("WebDesc")<>"" then
                              Session("RCM273_CarType") =s_li("WebDesc")
                           end if
                        end if
                        s_li.close
                        Set s_li=nothing


               '' Response.Write("<tr><td align='left' class='text' valign='top'>Vehicle Type: </td>")
         'Response.Write("<td class='text' valign='top' colspan=2 align='left'>"&s_cs("WebDesc")&"</td></tr>")
         'Response.Write("<tr><td align='left' class='text' valign='top'> </td>")
         'Response.Write("<td class='text' valign='top' align='left' colspan=2>"&s_cs("VehicleDesc")&"</td></tr>")




          '---the rates-----------
          costEachSeason=0
          totalStandardRate=0
          totalRate=0
          Rate=0
          for i=1 to SeasonCount
                         DiscountRate=0
                         DiscountID=0
                         DiscountType="p"
                        'Get the Season Start Date to use for Discount
                        SQL="select DiscountRate, DiscountType, DiscountName, ID from Discount	"
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
                          DiscountName=s_DisRate("DiscountName")
                          DiscountID=s_DisRate("ID")

                          Session("RCM273_DiscountName")=DiscountName
                          Session("RCM273_DiscountType")=DiscountType
                          Session("RCM273_DiscountRate")=DiscountRate
                          Session("RCM273_DiscountID")=DiscountID
                        end if

                     '--for each Season Rate
                     SQL="select * FROM QCarRateDetails "
                     SQL=SQL&"WHERE CarSizeID="&s_cs("ID")&" "
                     SQL=SQL&"AND CarRateStructureID="&Session("RCM273_RateStructureID"&i&"")&" "
                     SQL=SQL&"AND (SeasonID)='"&Session("RCM273_SeasonID"&i&"")&"' "
                     'Response.Write(SQL)
                     Set s_m=webConn.Execute(SQL)

                     if NOT s_m.EOF   then
                            '----only display the rate if the rate >0
                                 if Session("RCM273_DiscountType")="p" then
                                          Rate=s_m("Rate")*(1-Session("RCM273_DiscountRate")/100)
                                 else
                                          Rate=s_m("Rate")-Session("RCM273_DiscountRate")
                                 end if
                                 costEachSeason=Rate*Session("RCM273_NoOfDaysEachSeason"&i&"")
                                 totalRate=totalRate+costEachSeason
                                 totalStandardRate=totalStandardRate+s_m("Rate")*Session("RCM273_NoOfDaysEachSeason"&i&"")
                                 if  Session("RCM273_useEverageRate")<>"Yes" then
                                    if Session("RCM273_DiscountType")="p" then
                                       Discount=""&Session("RCM273_DiscountRate")&"% Discount"
                                    else
                                       Discount=""&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_DiscountRate"))&" Discount"
                                    end if
                                    'Response.Write("<div class='row'>")
                                    if Session("RCM273_DiscountRate")>0 then
                                          Response.Write("<div class='row'><div class='col-xs-7'><p class='text-left'><strong>"&Session("RCM273_NoOfDaysEachSeason"&i&"")&" Days at "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Rate,2)&" per day </strong><br /><span class='label label-danger'>"&Discount&"</span></p></div>" )
                                    else
                                          Response.Write("&nbsp;"&Session("RCM273_NoOfDaysEachSeason"&i&"")&" Days @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Rate,2)&" (per day)")
                                    end if
                                    Response.Write("<input  type='hidden'  name='StandardRate"&i&"'  value='"&s_m("Rate")&"' />")
                                    Response.Write("<input  type='hidden'  name='Rate"&i&"'  value='"&Rate&"' />")
                                    Response.Write("<div class='col-xs-5'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(costEachSeason,2)&"</p></div></div>")
                                 end if
                       else'--no rate set up, copy last season rate
                                 if  WeekendRate="Yes"   then
                                    SQL= "SELECT Top 1 (FromDay),ToDay, ID, RateName FROM CarRateStructure where ID<>1 "
                                    SQL=SQL&" AND FromDay<="&Session("RCM273_DaysForRate")&" order by FromDay desc "
                                    Set s_rs=webConn.Execute(SQL)
                                    if NOT s_rs.EOF   then  '---  not weekend (more then 2 days)
                                             RateName=s_rs("RateName")
                                             RateStructureID=s_rs("ID")
                                             'Session("RCM273_RateStructureID"&i&"")=s_rs("ID")
                                    end if
                                    s_rs.Close
                                    SET s_rs=nothing
                                    SQL="SELECT * FROM QCarRateDetails  where CarRateStructureID="&RateStructureID&" and  SeasonID='"&Session("RCM273_SeasonID1")&"'  and LocationID="&CINT(Session("RCM273_PickupLocationID"))&"  and CarSizeID="&CInt(Session("RCM273_CarSizeID"))&"  "
                                    Set s_r=webConn.Execute(SQL)
                                    IF NOT s_r.EOF THEN
                                          if ISNull(s_r("Rate"))<>True then
                                                Rate=s_r("Rate")
                                          end if
                                    END IF
                                    s_r.close
                                    set s_r=nothing
                                          costEachSeason=Rate*Session("RCM273_NoOfDaysEachSeason"&i&"")
                                          if Session("RCM273_useEverageRate")<>"Yes" then
                                                Response.Write("<tr><td class='text' align=left nowrap='nowrap'>")
                                                Response.Write("&nbsp;"&Session("RCM273_NoOfDaysEachSeason"&i&"")&" Days @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Rate,2)&" (per day)")
                                                Response.Write("<input  type='hidden'  name='StandardRate"&i&"'  value='"&Rate&"' />")
                                                Response.Write("<input  type='hidden'  name='Rate"&i&"'  value='"&Rate&"' />")
                                                Response.Write("</td><td class='text' align=right nowrap='nowrap'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(costEachSeason,2)&"</td></tr>")
                                          end if
                                 else'---end if  '--if WeekendRate=0 and WeekendRate="Yes"   then
                                          costEachSeason=Rate*Session("RCM273_NoOfDaysEachSeason"&i&"")
                                         if Session("RCM273_useEverageRate")<>"Yes" then
                                             Response.Write("<tr><td class='text' align=left nowrap='nowrap'>")
                                             Response.Write("&nbsp;"&Session("RCM273_NoOfDaysEachSeason"&i&"")&" Days @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Rate,2)&" (per day)")
                                             Response.Write("<input  type='hidden'  name='StandardRate"&i&"'  value='"&Rate&"' />")
                                             Response.Write("<input  type='hidden'  name='Rate"&i&"'  value='"&Rate&"' />")
                                             Response.Write("</td><td class='text' align=right nowrap='nowrap'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(costEachSeason,2)&"</td></tr>")
                                           end if
                                 end if
                                 totalRate=totalRate+costEachSeason
                       end if
                       s_m.close
                       set s_m=nothing
                      'Session("RCM273_FreeDayRate")=Rate
            next
            EverageRate=round(totalRate/Session("RCM273_totalDays4Rate"),2) '---not use  Session("RCM273_DaysForRate"), it is fixed to integer

            StandardRate= round(totalStandardRate/Session("RCM273_totalDays4Rate"),2)
            Session("RCM273_FreeDayRate")=EverageRate
            totalRate=EverageRate*Session("RCM273_totalDays4Rate")

            if Session("RCM273_useEverageRate")="Yes" then
                  '---if display everage rate unblock the code below, and block the code line Response.Write("<input type='hidden' name='SeasonCount' value='"&i&"' />")
                  if Session("RCM273_DiscountRate")>0 then
                         if Session("RCM273_DiscountType")="p" then
                                       Discount="("&Session("RCM273_DiscountRate")&"% discount)"
                        else
                                       Discount="("&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_DiscountRate"))&" discount)"
                        end if
                       Response.Write("<tr><td class='text' align='left'>")
                        Response.Write(""&Session("RCM273_totalDays4Rate")&" Days <font color=red>"&Discount&"</font> @ "&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(EverageRate,2)&" (per day)")
                  else
                       Response.Write("<tr><td class='text' align='left'>")
                        Response.Write(""&Session("RCM273_totalDays4Rate")&" Days  @ "&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(EverageRate,2)&" (per day)")
                  end if
                  i=1
                  Response.Write("<input type='hidden' name='NoOfDaysEachSeason"&i&"' size='5' value='"&Session("RCM273_totalDays4Rate")&"' />")
                  Response.Write("<input type='hidden' name='StandardRate"&i&"' size=3 value="&StandardRate&" />")
                  Response.Write("<input type='hidden' name='Rate"&i&"' size=3 value="&EverageRate&" />")
                  Response.Write("<input type='hidden' name='SeasonCount' value='"&i&"' />")
                  '---end display everage rate --
             end if

             Response.Write("<input  type='hidden'  name='TotalDays'  value='"&Session("RCM273_totalDays4Rate")&"' />")
           if FixedDiscountRate<>0 then
                  Response.Write("<input  type='hidden'  name='FixedDiscountID'  value='"&FixedDiscountID&"' />")
                  Response.Write("<input  type='hidden'  name='FixedDiscountRate'  value='"&FixedDiscountRate&"' />")
                  Response.Write("<div class='row'><div class='col-xs-7'><p class='text-left'><strong>"&FixedDiscountName&"</strong></p></div><div class='col-xs-5'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(FixedDiscountRate,2)&"</p></div></div>")
            else
                  Response.Write("<input  type='hidden'  name='FixedDiscountID'  value='0' />")

            end if
             RelactionFees
             'totalRate=totalRate +  FixedDiscountRate +Session("RCM273_RelocationFee")
            'Response.Write("<tr><td class='text'  nowrap='nowrap'>Total   "&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(totalRate,2)&"</td></tr>")
             FreeDaysOffer
             totalCost=EverageRate*Session("RCM273_totalDays4Rate") + FixedDiscountRate + Session("RCM273_RelocationFee")-Session("RCM273_FreeDayTotal")
             RentalCost=EverageRate*Session("RCM273_totalDays4Rate")

         '----------------check Holiday Charges for pickup  from table HolidayextraFees  ---------------------
         SQL="SELECT Fees,eh_HolidayName FROM ExtraFees4Holiday,ExtraFees WHERE eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestPickDate")&"' and  eh_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"')  "
        ' Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)

       if NOT s_ex.EOF THEN
               Response.Write("<tr><td class='text' align=left nowrap>"&s_ex("eh_HolidayName")&"</td><td class='text' colspan=2 align=left>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_ex("Fees"),2)&"</td></tr>")
               TotalCost=TotalCost+ s_ex("Fees")
      end if
      s_ex.close
      Set s_ex=Nothing

      '--check Holiday Charges for dropoff  from table HolidayextraFees
       SQL="SELECT Fees,eh_HolidayName  FROM ExtraFees4Holiday,ExtraFees WHERE   eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestdropDate")&"' and  eh_LocationID='"&Session("RCM273_DropoffLocationID")&"') "
       'Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)
      if NOT s_ex.EOF THEN
               Response.Write("<tr><td class='text' align=left nowrap>"&s_ex("eh_HolidayName")&"</td><td class='text' colspan=2 align=left >"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_ex("Fees"),2)&"</td></tr>")
              TotalCost=TotalCost+ s_ex("Fees")
      end if
      s_ex.close
      Set s_ex=Nothing
          '----------------END check Holiday Charges for pickup  from table HolidayextraFees  ---------------------

              '----Mandatory Feees-------

          SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
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

            MandatoryExtraFeesNoDays=1

             if s_ex("Type")="Daily" then

                  '---daily fees always use   Session("RCM273_TotalRentalDays") (fixed number)
                        MandatoryExtraFeesNoDays=Session("RCM273_TotalRentalDays")
                        EachMandatoryExtraFees=MandatoryExtraFeesNoDays*s_ex("Fees")
             elseif s_ex("Type")="Fixed" then
                        EachMandatoryExtraFees = s_ex("Fees")
             else
                       EachMandatoryExtraFees = (s_ex("Fees")/100)*RentalCost
             end if
                     '--27/Jun/2009 for Daily and % fees allow check max price in step4
                if s_ex("MaxPrice") >0  and  EachMandatoryExtraFees>s_ex("MaxPrice") then
                                 EachMandatoryExtraFees=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and EachMandatoryExtraFees<s_ex("MaxPrice") then
                                  EachMandatoryExtraFees=s_ex("MaxPrice")
                end if
                 '----Stamp Duty for Mandatory extraFees ---------
                  eachMandatoryExtraStamp=0
                  if s_ex("StampDuty")="True" Then
                        eachMandatoryExtraStamp=EachMandatoryExtraFees
                 end if
                 MandatoryExtraStamp=MandatoryExtraStamp+eachMandatoryExtraStamp

                 '----GST for Mandatory extraFees ---------
                  eachMandatoryExtraGST=0
                  if s_ex("GST")="True" Then
                        eachMandatoryExtraGST=EachMandatoryExtraFees
                  end if
                 MandatoryExtraGST=MandatoryExtraGST+eachMandatoryExtraGST

                     TotalMandatoryExtraFees=EachMandatoryExtraFees+ TotalMandatoryExtraFees

                   Response.Write("<tr><td class='text' align=left  nowrap='nowrap'>"&s_ex("Name")&"<td class='text' colspan=2 align=right>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachMandatoryExtraFees,2)&"</td></tr>")



      s_ex.MoveNext
     j=j+1
     Loop
     s_ex.close
     SET s_ex=nothing
     TotalCost=TotalCost+TotalMandatoryExtraFees


            SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
          SQL=SQL&" AND (VehicleSizeID='"&Session("RCM273_CarSizeID")&"' or VehicleSizeID=0) "
          SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
          SQL=SQL&" and WebItems=1 and Mandatory=1 AND inUse=1 and "
           SQL=SQL&" MerchantFee=1 " '---(PercentageTotalCost=1 or MerchantFee=1) "
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
                  if s_ex("GST")="True" then
                              MerchantFeeGST=MerchantFeeGST+EachMerchantFee
                  end if
                  if s_ex("StampDuty")="True" Then
                              MerchantFeeStamp=MerchantFeeStamp +EachMerchantFee
                  end if
                  TotalCost=TotalCost+EachMerchantFee
                  BankBaseCalculationFee=BankBaseCalculationFee+EachMerchantFee
               Response.Write("<tr><td class='text' align=left  nowrap='nowrap'>"&s_ex("Name")&" </td><td class='text' colspan=2 align=right>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachMerchantFee,2)&"</td></tr>")

          s_ex.MoveNext
         j=j+1
         Loop
         s_ex.close
         SET s_ex=nothing
            Response.Write("<hr /><div class='row'><div class='col-xs-12'><p class='text-right lead'><strong>Total:  <mark>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(totalCost,2)&"</mark></strong></p></div></div></div></div>")

        end if
          s_cs.Close
         set s_cs=nothing


END SUB
SUB FreeDaysOffer
   NoFreeDays=0
   ExtraFeeID=0
   Session("RCM273_FreeDayTotal")=0
   SQL=" select  fd_NoofFreeDays,fd_apply2EveryNoDays,fd_extraFeeID,max(fd_NoDays) as MinDay "
   SQL=SQL&" FROM FreeDaySpecial left join CampaignCode on fd_CampaignCodeID=CampaignCode.ID  "
   SQL=SQL&" WHERE fd_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
   if Session("RCM273_PromoCode")<>"" then
            SQL=SQL&" AND CampaignCode='"&Left(Session("RCM273_PromoCode"),30)&"'  "
    else
            SQL=SQL&" AND fd_CampaignCodeID=0  "
    end if
   SQL=SQL&" and (fd_CategoryID='"&CInt(Session("RCM273_CarSizeID"))&"' or fd_CategoryID=0) "
   SQL=SQL&" and fd_NoDays<="&Session("RCM273_TotalRentalDays")&" "
   SQL=SQL&" and fd_FreeDayStart<='"&Session("RCM273_RequestPickDate")&"' and fd_FreeDayEnd>='"&Session("RCM273_RequestPickDate")&"' "
   SQL=SQL&" AND fd_NoDays in (select max(fd_NoDays)   "
   SQL=SQL&" FROM  FreeDaySpecial left join CampaignCode on fd_CampaignCodeID=CampaignCode.ID  "
   SQL=SQL&" where fd_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  "
   if Session("RCM273_PromoCode")<>"" then
            SQL=SQL&" AND CampaignCode='"&Left(Session("RCM273_PromoCode"),30)&"'  "
    else
            SQL=SQL&" AND fd_CampaignCodeID=0  "
    end if
   SQL=SQL&" and (fd_CategoryID='"&CInt(Session("RCM273_CarSizeID"))&"' or fd_CategoryID=0) "
   SQL=SQL&" and fd_NoDays<="&Session("RCM273_TotalRentalDays")&"   "
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
             Session("RCM273_FreeDayTotal")=NoFreeDays*Session("RCM273_FreeDayRate")
         'Response.Write("<tr><td  style='height: 22px'  align='left' class='text'><b>You qualify for a "&NoFreeDays&" Free Day Special!</td></tr>")
           Response.Write("<tr><td align='left' class='text'>"&NoFreeDays&" Free Day Special! <td class='text' colspan=2 align=right> -"&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_FreeDayTotal"),2)&"</td></tr>")

    END IF
         Response.Write("<input  type='hidden'  name='NoFreeDays'  value='"&NoFreeDays&"' />")
         Response.Write("<input  type='hidden'  name='FreeDayRate'  value='"&Session("RCM273_FreeDayRate")&"' />")
         Response.Write("<input  type='hidden'  name='FreeDayExtraFeeID'  value='"&ExtraFeeID&"' />")

 END SUB %>
<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface==
'== File Name :  webbookingstep3.asp==
'== Method Name : ExtraForm==
'== Method Description :  Selects Required Extra fees based Location selection and displays to screen==
'== Tables Used : ExtraFees==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->

<%
 SUB ExtraForm  '--sub
    '--if the cost of extrafees > maxprice then display maxprice---
    SQL="select * from ExtraFees WHERE InsuranceExtra=0 AND  inUse=1 and (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
      SQL=SQL&" and WebItems=1  AND (VehicleSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or VehicleSizeID=0) "
       SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
      SQL=SQL&" and  Mandatory=0  order by  Type,QTYApply,Name"
    Set s_ex=webConn.Execute(SQL)
    if NOT s_ex.EOF  then        %>
               <h4 class="smallm_title centered bigger"><span>Choose extra equipment</span></h4>

             <div class="table-responsive">
               <table class="table table-hover">
                <thead>
                  <tr>
                      <th width="85%">Description</th>
                      <th width="10%">Price</th>
                      <th width="5%">QTY</th>
                  </tr>
                </thead>

                <tbody>

    <%         j=0
                 DO WHILE NOT s_ex.EOF
                  Response.Write("<tr><td><label><input type='checkbox'  name='ExtraFeesID"&j&"' value='"&s_ex("ID")&"' />")
                     if s_ex("Type")="Percentage" then
                              Response.Write(" "&s_ex("Name")&"</label></td><td>"&s_ex("Fees")&"% (of Rental)</td>")

                    else
                           '--for daily extra if the cost of extrafees > maxprice then display maxprice---
                              if s_ex("Type")="Daily" and s_ex("Maxprice")>0 and TotalRentalDays*s_ex("Fees")>s_ex("Maxprice") then
                                    Response.Write(" "&s_ex("Name")&"</label></td><td>"&Session("RCM273_CompanyCurrency")&" "&s_ex("Maxprice")&" (Fixed)</td>")
                              else
                                       Response.Write(" "&s_ex("Name")&"</label></td><td>"&Session("RCM273_CompanyCurrency")&" "&s_ex("Fees")&" ("&s_ex("Type")&")</td>")
                              end if
                    end if
                    if s_ex("QTYApply")=True  then
                    Response.Write("<td><select name='QTY"&j&"' class='form-control' style='width:60px;' onblur='checkNumeric(this);'><option value='1'>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option></select></td></tr>")
                    else
                    Response.Write("<td><select name='QTY"&j&"' class='form-control' style='width:60px; display: none;' onblur='checkNumeric(this);'><option value='1'>1</option><option value='2'>2</option><option value='3'>3</option><option value='4'>4</option><option value='5'>5</option></select></td></tr>")
                    end if
                    'Response.Write("<tr><td colspan='4'  bgcolor='"&Session("RCM273_CompanyColour")&"' style='height: 1px' >")
                    Response.Write("<input type='hidden'  name='MaxPrice"&j&"'  value='"&s_ex("MaxPrice")&"' />")
                    Response.Write("<input type='hidden'  name='ExtraFeesName"&j&"'  value='"&s_ex("Name")&"' />")
                    Response.Write("<input type='hidden'  name='ExtraFees"&j&"'  value='"&s_ex("Fees")&"' />")
                    Response.Write("<input type='hidden'  name='FeeType"&j&"' value='"&s_ex("Type")&"' />")
                    Response.Write("<input type='hidden'  name='GST"&j&"' value='"&s_ex("GST")&"' />")
                    Response.Write("<input type='hidden'  name='StampDuty"&j&"' value='"&s_ex("StampDuty")&"' />")
                    Response.Write("<input type='hidden'  name='PercentageTotalCost"&j&"' value='"&s_ex("PercentageTotalCost")&"' />")
                    Response.Write("<input type='hidden'  name='MerchantFee"&j&"' value='"&s_ex("MerchantFee")&"' />")

                    'Response.Write("</td></tr>")
                    if s_ex("ExtraDesc")<>"" then
                       Response.Write("<tr><td></td><td><span class='text-muted'>"&s_ex("ExtraDesc")&"</td><td></td></tr>")
                    end if
                 s_ex.MoveNext
            j=j+1
            Loop
               s_ex.close
            SET s_ex=nothing
                 Response.Write("<tr><td colspan='4'><input type='hidden' name='ExtraFeeCount' value='"&j&"' /></td></tr>")
              Response.Write("</tbody>")
     END IF
 END SUB
 %>
<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface==
'== File Name :  webbookingstep3.asp==
'== Method Name : InsuranceExtra==
'== Method Description :  Selects Required Insurance fees based Location selection and displays to screen==
'== Tables Used : ExtraFees==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->
<%
SUB InsuranceExtra '--new code, jul 2012 by rita
        VehicleSizeID=CInt(Session("RCM273_CarSizeID"))
        TotalDays=  Session("RCM273_TotalRentalDays")
         iAge=0
         iAge=Session("RCM273_driverage")

         SQL="select top 1 VehicleSizeID, CategoryTypeID "
         SQL=SQL&" FROM ExtraFees WHERE InsuranceExtra=1 and inUse=1   "
         SQL=SQL&" and (LocationID=0 or LocationID="&CINT(Session("RCM273_PickupLocationID"))&" )  "
         SQL=SQL&" and WebItems=1 and (VehicleSizeID=0 or VehicleSizeID="&VehicleSizeID&") "
         SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
         SQL=SQL&" and ( FromAge=0 or ( FromAge<="&iAge&" and (ToAge>="&iAge&" or (ToAge=0 and FromAge>0)) ) )  "
         SQL=SQL&" and ((FromDay=0 and ToDay=0) or ( FromDay<="&TotalDays&" and (ToDay>="&TotalDays&" or (ToDay=0 and FromDay>0)) )) "
        ' SQL=SQL&" Group By   VehicleSizeID, CategoryTypeID "
        ' SQL=SQL&" order by  VehicleSizeID Desc, CategoryTypeID Desc "
         'Response.Write(SQL)
      Set s_in1=webConn.Execute(SQL)

    if s_in1.EOF then '---if no insurance set up then do not show the dropdown list for selection
            Response.Write("<input type=hidden  name=InsuranceID Value='0' >")
    else

            Response.Write("<h4 class='smallm_title centered bigger'><span>Choose additional protection</span></h4>")
            'Response.Write("<tr><td  align=left >")
            'Response.Write("<TABLE  width=100% align=left  cellSpacing=0 cellPadding=2  border=0>")
            SQL="select * "
            SQL=SQL&" FROM ExtraFees WHERE InsuranceExtra=1 and inUse=1   "
             SQL=SQL&" and (LocationID=0 or LocationID="&CINT(Session("RCM273_PickupLocationID"))&" )  "
            SQL=SQL&" and   WebItems=1 "
           ' SQL=SQL&" and CategoryTypeID="&s_in1("CategoryTypeID")&" "
            'SQL=SQL&" and VehicleSizeID="&s_in1("VehicleSizeID")&" "
            SQL=SQL&" and (CategoryTypeID=0 or CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" ) "
            SQL=SQL&" and (VehicleSizeID=0 or VehicleSizeID="&VehicleSizeID&") "
            SQL=SQL&" and ( FromAge=0 or ( FromAge<="&iAge&" and (ToAge>="&iAge&" or (ToAge=0 and FromAge>0)) ) )  "
            SQL=SQL&" and ((FromDay=0 and ToDay=0) or ( FromDay<="&TotalDays&" and (ToDay>="&TotalDays&" or (ToDay=0 and FromDay>0)) )) "
            SQL=SQL&" order by Name "
             Set s_in=webConn.Execute(SQL)
            'Response.Write(SQL)
            DO WHILE NOT s_in.EOF
                     if  s_in("Mandatory")=True then
                     Response.Write("<div class='radio'><label style='width: 100%;'><input type=radio  name=InsuranceID  Value="&s_in("ID")&" CHECKED> "&s_in("Name")&" <span style='float: right;'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_in("Fees"),2)&"  (Daily)</span> ")
                     else
                     Response.Write("<div class='radio'><label style='width: 100%;'><input type=radio  name=InsuranceID  Value="&s_in("ID")&" > "&s_in("Name")&" <span style='float: right;'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_in("Fees"),2)&"  (Daily)</span>")
                  end if
                  if s_in("ExtraDesc")<>"" then
                      Response.Write("<span class='text-muted'>"&s_in("ExtraDesc")&"</span></label></div>")
                  else
                  Response.Write("</label></div>")
                  end if


            s_in.MoveNext
            Loop
             s_in.close
            set s_in=nothing
            'Response.Write("</table></td></tr>")
    end if
    s_in1.close
    set s_in1=nothing


END SUB



%>
<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface==
'== File Name :  webbookingstep3.asp==
'== Method Name : KmFeesSelection==
'== Method Description :  Selects valid KM fees based on Location and Web Available flags and displays to screen==
'== Tables Used : ExtraFees==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->


<%


SUB KmsFeesSelection
      Session("RCM273_KmsDesc")=""

      '---Jul 2012 new code, Rita, BookingDate, PickupDate added
      VehicleSizeID=CInt(Session("RCM273_CarSizeID"))
      PickupLocationID =Session("RCM273_PickupLocationID")
        '--if 4 + days is unlimited kms, 1 -3 days 100kms free $0.24 per extra kms,
       '--if the rental is 3.6 days then should be same as 3 days rental

        'TotalDays=  Session("RCM273_TotalRentalDays")
       '--Session("RCM273_KmsTotaldays")=Round(Totalmin/1440,2)
       'KmsTotaldays=Session("RCM273_KmsTotaldays")
     ' KmsTotaldays=Fix(Session("RCM273_TotalRentalDays"))
      KmsTotaldays=Session("RCM273_DaysForRate")
    BookingDate=Year(Date)&"-"&Month(Date)&"-"&Day(Date)

   SQL="select top 1  LocationID,VehicleSizeID, DateRange "
   SQL=SQL&" from KmsFree "
   SQL=SQL&" WHERE (WebAvaliable=1 ) "
   SQL=SQL&" and (LocationID=0 or LocationID="&PickupLocationID&" )  "
   SQL=SQL&" and (VehicleSizeID=0 or VehicleSizeID="&VehicleSizeID&") "
   SQL=SQL&" and ((FromDay=0 and ToDay=0)  or ( FromDay<="&KmsTotaldays&" and (ToDay>="&KmsTotaldays&" or (ToDay=0 and FromDay>0)) )) "
   SQL=SQL&" AND (DateRange=0 or (DateRange=1 "
   SQL=SQL&" AND PickupDateFrom<='"&RequestPickDate&"' and DropoffDateTo>='"&RequestDropDate&"' "
   SQL=SQL&" AND ( (BookingDateFrom<='"&BookingDate&"' and BookingDateTo>='"&BookingDate&"') "
   SQL=SQL&" or (BookingDateFrom is Null or YEAR(BookingdateFrom)=2100) ) ))"
   SQL=SQL&" Group By LocationID, VehicleSizeID,  DateRange "
   SQL=SQL&" order by LocationID desc, VehicleSizeID Desc, DateRange Desc "
'Response.Write SQL
    Set s_tk = webConn.Execute(SQL)
   if NOT s_tk.EOF then
           Response.Write("<tr><td  bgcolor='"&Session("RCM273_CompanyColour")&"' style='height: 22px' colspan='4' align='center' class='header'>Select "&Session("RCM273_Mileage")&" option</td></tr>")

         SQL="select * from KmsFree WHERE  (WebAvaliable=1 ) "
         SQL=SQL&" and LocationID="&s_tk("LocationID")&"  "
         SQL=SQL&" and VehicleSizeID="&s_tk("VehicleSizeID")&" "
         SQL=SQL&" and ((FromDay=0 and ToDay=0)  or ( FromDay<="&KmsTotaldays&" and (ToDay>="&KmsTotaldays&" or (ToDay=0 and FromDay>0)) )) "
         if s_tk("DateRange")=0 then
               SQL=SQL&" and DateRange=0 "
         else 'if s_tk("DateRange")=1 then
               SQL=SQL&" AND DateRange=1 "
               SQL=SQL&" AND PickupDateFrom<='"&RequestPickDate&"' and DropoffDateTo>='"&RequestDropDate&"' "
               SQL=SQL&" AND ( (BookingDateFrom<='"&BookingDate&"' and BookingDateTo>='"&BookingDate&"') "
               SQL=SQL&" or (BookingDateFrom is Null or YEAR(BookingdateFrom)=2100) )"
         end if
         SQL=SQL&" ORDER BY DefaultKM,KmsFree,AddKmsFee "


         'Response.Write SQL
          Set s_km = webConn.Execute(SQL)

               DO WHILE NOT s_km.EOF
             maxprice=""
             if s_km("maxprice")>0 then
                    maxprice=", max charge "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("maxprice"))&" per day"
             end if
             KmsDesc=s_km("KmsFree")&" "&Session("RCM273_Mileage")&" free per day, additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("AddKmsFee"))&" "&maxprice
             if s_km("AddKmsFee")=0 and s_km("KmsFree")=0 and s_km("DailyRate")>0 then
                     KmsDesc="Unlimited Kms, daily rate @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("DailyRate"))
             end if
             if s_km("AddKmsFee")>0  and s_km("DailyRate")>0 then
                     KmsDesc=s_km("KmsFree")&" "&Session("RCM273_Mileage")&" per day,  "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("DailyRate"))&"/day,  additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("AddKmsFee"))
                     'KmsDesc="Daily rate @"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("DailyRate"))&", "&s_km("KmsFree")&" "&Session("RCM273_Mileage")&" free per day, additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("AddKmsFee"))
             end if
             if s_km("AddKmsFee")=0 and s_km("KmsFree")=0 and s_km("DailyRate")=0 then
                     KmsDesc="Unlimited "&Session("Mileage")
             end if
             if  s_km("DefaultKM")=True then
                  Response.Write("<tr><td  class='text'><input type='radio'  name='KmsFreeID'   value='"&s_km("ID")&"' checked='checked' /><td class='text' align='left' colspan='3'>"&KmsDesc&"</option>")
             else
                  Response.Write("<tr><td  class='text'><input type='radio'  name='KmsFreeID'   value='"&s_km("ID")&"' /><td class='text' align='left' colspan='3'>"&KmsDesc&"</option>")
             end if
       s_km.MoveNext
       Loop
       s_km.close
       Set s_km=nothing

   end if
   s_tk.close
   Set s_tk=nothing

END SUB



%>




<!-- +++++++++++++++++++++++++++++++++++++++++++++++==
'== Module Name : Web Site Interface==
'== File Name :  webbookingstep3.asp==
'== Method Name : HTML Page Starts here - Allows entry and selection of extra fees, kms, and insurance options==
'== Tables Used : CarRateHourly, QSeason, Location==
'== +++++++++++++++++++++++++++++++++++++++++++++++   -->

<div class="container">
        <h4 class="smallm_title centered bigger"><span>Your quote details</span></h4>



  <%

     



       Graceperiod=Session("RCM273_Graceperiod")
 RequestPickDate = Session("RCM273_RequestPickDate")

        RequestPickTime=Session("RCM273_RequestPickTime")
      RequestDropDate = Session("RCM273_RequestDropDate")
        RequestDropTime=Session("RCM273_RequestDropTime")
       
        RequestPickDateTime = CDate(RequestPickDate&" "& RequestPickTime)
        RequestDropDateTime =CDate(RequestDropDate&" "& RequestDropTime)
'Set s_li=webConn.Execute(SQL)
'Session("RCM273_CarImageName")=s_li("ImageName")

SQL=" SELECT  CarSize.*  FROM CarSize  "
       SQL=SQL&" WHERE CarSize.ID="&CInt(Session("RCM273_CarSizeID"))&" "

         Set s_cs=webConn.Execute(SQL)

         theimage=RCMURL&"/DB/"&CompanyCode&"/"&s_cs("ImageName")

        Response.Write("<div class='row'><div class='col-xs-5'>")
   
        Response.Write(" <p class='text-center'><img src='"&theimage&"'  class='vehicle-image' alt='' /></p></div>")

       'Set s_li=nothing
        'set s_cs=nothing
        'Response.Write("<tr><td align='left' class='text' valign='top'>Vehicle Type:</td>")
        'Response.Write("<td align='left' class='text' valign='top'>")

        'Response.Write("<table align='left' cellspacing='0' cellpadding='2'  border='0'>")
        Response.Write("<div class='col-xs-7'><h2>"&Session("RCM273_CarType")&"</h2><hr />")
        'Response.Write("<div class='col-xs-7'><h2>"&s_cs("VehicleDesc")&"</h2><hr />")
        's_cs.close

         Response.Write("<div class='row'><div class='col-xs-2'>")
          Response.Write("<p class='text-left'><strong>Pickup:</strong></p></div>")
         Response.Write("<div class='col-xs-3'><p class='text-right'>")
          Response.Write(Session("RCM273_PickupLocation"))
          Response.Write("</div>")
          Response.Write("<div class='col-xs-1'>&nbsp;</div>")

           Response.Write("<div class='col-xs-2'>")
           Response.Write("<p class='text-left'><strong>Pickup Date:</strong></p>")
           Response.Write("</div><div class='col-xs-4'><p class='text-right'>"&WeekdayName(WeekDay(Session("RCM273_RequestPickDate")))&", ")
           Response.Write(Day(Session("RCM273_RequestPickDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestPickDate"))),3)&"/"&Year(Session("RCM273_RequestPickDate")))
           Response.Write("&nbsp;")
           Response.Write(Session("RCM273_RequestPickTime"))
          Response.Write("</p>")
          Response.Write("</div></div>")



           Response.Write("<div class='row'><div class='col-xs-2'>")
           Response.Write("<p class='text-left'><strong>Return:</strong></p>")
           Response.Write("</div><div class='col-xs-3'><p class='text-right'>")
           Response.Write(Session("RCM273_DropoffLocation"))
           Response.Write("</p></div>")
          Response.Write("<div class='col-xs-1'>&nbsp;</div>")

           Response.Write("<div class='col-xs-2'>")
            Response.Write("<p class='text-left'><strong>Return Date:</strong></p>")
           Response.Write("</div><div class='col-xs-4'><p class='text-right'>"&WeekdayName(WeekDay(Session("RCM273_RequestDropDate")))&", ")
           Response.Write(Day(Session("RCM273_RequestDropDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestDropDate"))),3)&"/"&Year(Session("RCM273_RequestDropDate")))
           Response.Write("&nbsp;")
           Response.Write(Session("RCM273_RequestDropTime"))
          Response.Write("</p></div></div>")
          Response.Write("<hr />")


               '-----Start Hourly Rate ------------
        '----1st need find the extra hours in which season, (use the extra hours as a day to find the season)
        '--- need a acture No of Days (actureTotalDays) and actureNoOfDaysEachSeason to find out the extra hours in which season
        '---if there are hourly rate set up then use it for calculattion
       '----if the total cost for the no of hours > = one day rate, then do not calculate hourly rate

        theExtraHour=0
        HourRate=0
        if Session("RCM273_RequestPickTime")<> Session("RCM273_RequestDropTime") then
            theExtraHour=Round(((1440*(CDate(Session("RCM273_RequestDropDate")&" "& Session("RCM273_RequestDropTime")) - CDate(Session("RCM273_RequestDropDate")&" "& Session("RCM273_RequestPickTime")))-Graceperiod) /60),2)
                if  theExtraHour > Fix(theExtraHour)  then
                  theExtraHour=Fix(theExtraHour)+1
            end if
                'theExtraHour=Round(theExtraHour,2)
                'Response.Write(theExtraHour)
            'Response.Write("<br>")
            SQL="select * from CarRateHourly WHERE (LocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"'  and HourFrom<"&theExtraHour&" and HourTo>="&theExtraHour&" "
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
        Response.Write("<input type='hidden' name='HourRate' value='"&HourRate&"' />")                  'SeasonCount





    '---1. get the no of days for each season----
    '---2. For each car type, get each season rate, then calculate the total cost for each car type---



          SQL="select * from Season where LocationID="&CINT(Session("RCM273_PickupLocationID"))&" and notActive=0 and (Season='Default' or  (EndDate>='"&RequestPickDate&"' and StartDate<='"&RequestDropDate&"')  ) order by StartDate DESC "
          'Response.Write(SQL)
          SET s_m=webConn.Execute(SQL)



    TotalDays= Session("RCM273_TotalRentalDays") '-----------------------------need here for 12 hour rate------------

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
                      IF  (Round((SeasonStartingDateTime-RequestPickDateTime)*1440,2)/60)=>12  and Round((RequestDropDateTime-SeasonStartingDateTime)*1440,2)/60>0 then
                              ' _____________ calculate the no of days in each season
                          if Round((SeasonEndingDateTime - RequestDropDateTime )*1440,2)/60>=0  then
                               '--the season strat time should be same as Pickup time
                                 SStartingDateTime=CDate(s_m("StartDate")&" "& Session("RCM273_RequestPickTime"))
                                 themin= Round((RequestDropDateTime - SStartingDateTime)*1440-Graceperiod,2)
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
                                  themin= Round((RequestDropDateTime - RequestPickDateTime)*1440-Graceperiod,2)
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
                        end      if
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

             SeasonCost=0
             if NoOfDaysEachSeason>0 then
                  i=i+1
                  if s_m("Season")="Default" then
                        Season="Rates"
                        Session("RCM273_Season"&i&"")="Standard Rates"

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
             if  i=1 and theExtraHour<>0 and  HourRate<>0 then
                             '-- add the extra hour to the  1st season number of days
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
                        Response.Write("<input type='hidden' name='actureNoOfDaysEachSeason"&i&"' size='5' value='"&actureNoOfDaysEachSeason&"' />")
                        if  Session("RCM273_useEverageRate")<>"Yes" then
                        Response.Write("<input type='hidden' name='NoOfDaysEachSeason"&i&"' size='5' value='"&NoOfDaysEachSeason&"' />")
                        end if
                        Response.Write("<input type='hidden' name='Season"&i&"' size='5' value='"&Session("RCM273_Season"&i&"")&"' />")
                       Response.Write("<input type='hidden' name='SeasonID"&i&"' size='5' value='"&s_m("ID")&"' />")



               end if
                  s_m.MoveNext
            Loop
            s_m.Close
            SeasonCount=i
            if  Session("RCM273_useEverageRate")<>"Yes" then
                  Response.Write("<input type='hidden' name='SeasonCount' value='"&i&"' />")                  'Count
            end if




                '---get car type, for each car type, get each season rate, then calculate total cost for each type
          Less1DayHireHourRate=0
          subGETHourlyRate

           if Less1DayHireHourRate=0 then
          findTheRate2
          end if
                 'Response.Write("</table></td>")



         'Response.Write("</table></td></tr>")



                InsuranceExtra
                ExtraForm
                KmsFeesSelection
                SelectAreaofUsed

                 'Response.Write("<tr>")
                'Response.Write("<td colspan='4' align='center' class='text'>")
                Response.Write("<tfoot><tr><td>")
                Response.Write("<input type='button' name='Back'  value='Back'  class='btn btn-default' onclick='javascript:history.back(-1)' />")
                  '---do NOT USE  image submit buttons, it will have problem to pass the value to next page (only the regular submit button that Microsoft browser manages to deal with properly )
                  '-- if you want to change the looks of the buttons,  use css to style them ---
               'Response.Write("&nbsp;&nbsp;<input  name=submit type=image src=quote.jpg value='Email Me Quote'  />")  '---not work!!!!!!!



               if Request.QueryString("categoryStatus")="2" then '---LIMITED AVAILABILTY,
                     Response.Write("&nbsp;<input NAME=submit type='submit' class='btn btn-default' value='Email Me Quote'  /></td>") '---if does not allow email quote, block this line of code
                     Response.Write("<td colspan='2'><input NAME=submit type='submit' class='btn btn-success btn-block btn-lg' value='Request Booking'  />")
                else
                     Response.Write("&nbsp;<input NAME=submit type='submit' class='btn btn-default' value='Email Me Quote'  /></td>")
                     Response.Write("<td colspan='2'><input NAME=submit type='submit' class='btn btn-info btn-block btn-lg' value='BOOK NOW'  />")
                end if
                Response.Write("</td></tr></tfoot></table>")


      webConn.CLOSE
         SET webConn=nothing

%>
</div>
  
  </form>
</div>
<!-- END RCM HTML CODE-->

<!-- #include file="include_footer.asp" -->

</body>
</html>

