<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->

<!-- #include file="include_meta.asp" -->

</head>



<%
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
%>


<body  class="webstep6" >

<!-- #include file="include_header.asp" -->

<!-- RCM HTML CODE-->

<%       if Session("RCM273_OnlineBookingNo")="" then
             Response.Redirect "webstep2.asp"
         end if

         CompanyCode="PtWestCoastCampers273"
         RCMURL="https://secure2.rentalcarmanager.com"

         imageURL = "http://secure2.rentalcarmanager.com"


         Session("RCM273_CarImageName")=Replace(Session("RCM273_CarImageName"), "." , "&#46;")
         theimage="http://secure2.rentalcarmanager.com/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")
         ServerID=2
         RCMCompanyID=273

         logo="logo.gif"
         FilePath=Server.MapPath("/DB")&"\"&CompanyCode

         logo="logo.gif"
         FilePath=Server.MapPath("/DB")&"\"&CompanyCode
          Set FileObject=Server.CreateObject("Scripting.FileSystemObject")
          '----------connect RCM Database---------------------
          Dim webConn
         Set webConn = Server.CreateObject("ADODB.Connection")
         DatabaseFile="PtWestCoastCampers273"
         webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"

        BookingNo = Left(Session("RCM273_BookingBufferNo"),8)


         GSTInclusive = "Yes"
         Set s_st=WebConn.Execute("SELECT * FROM SystemTable WHERE Code='GSINC'  ")
         If not s_st.EOF then
               GSTInclusive = s_st("syValue")
         END IF
         s_st.CLOSE
         SET s_st=NOTHING


            SQL="update Reslog set  exRateNotes='Credit Card entered, and email send online' where resBufferNo = "&CLng(BookingNo)&" "
             webConn.Execute(SQL)
            ' Response.Write(SQL)
            '  Response.Write("<br>")
            '---confirmed booking with cc details entered, we update to booking
            '--for old booking steps this may not work !!!!!!!!!!!!!!!!!!
             if Session("RCM273_bookingType")="Booking" then

                   SQL="UPDATE ReservationBuffer SET "
                  SQL=SQL&" Quotation =0  "
                  SQL=SQL&"WHERE (ReservationNo) ='"&CLng(BookingNo)&"'"
                  webConn.Execute(SQL)
                  ' Response.Write(SQL)

                   if Session("RCM273_bookingType") ="Quote" then
                        ReservationTypeID=3
                  else
                        '--insert Unallocaited RA# to WebReservaiton table,
                         if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY
                              ReservationTypeID=0 '--need release the on request menu in RCM
                        else
                              ReservationTypeID=1  '--confirmed booking
                        end if
                  end if

                   SQL="UPDATE WebReservation SET "
                   SQL=SQL&" ReservationTypeID ="&ReservationTypeID&"  "
                   SQL=SQL&"WHERE (ReservationNo) ='"&Clng(Session("RCM273_OnlineBookingNo"))&"'"
                   webConn.Execute(SQL)
                    'Response.Write "<Br>"
                   'Response.Write(SQL)


              end if


        SQL="SELECT QReservationforBuffer.*, TotalCost = dbo.BookingValueBuffer(ReservationNo, '"&GSTInclusive&"'), "
        SQL=SQL&" CASE WHEN dbo.CategoryPerNightRate(CarSizeID)=0 THEN 'day' ELSE 'night' END as sRateTypeDesc, "
         SQL=SQL&" dbo.BookingBalanceBuffer(ReservationNo,'"&GSTInclusive&"')  AS BalanceDue "
         SQL=SQL&" from QReservationforBuffer where ReservationNo="&CLng(BookingNo)&"  "
      Set s_d=webConn.Execute(SQL)
      If not s_d.EOF then





          Set s_sy=webConn.Execute("SELECT * FROM SystemTable WHERE Code='CURCY'  ")
         CurrencySample="$"
         If not s_sy.EOF then
               CurrencySample= s_sy("syValue")
         END IF
        s_sy.CLOSE
         SET s_sy=NOTHING

         Function DisplayCurrency(n)
           DisplayCurrency = CurrencySample&""&FormatNumber(n,2)
         End Function


      '-----use company info

        if s_d("BrandID")<>0 then
            Set s_com=webConn.Execute("SELECT * from Company WHERE ID="&s_d("BrandID")&"   ")
        else
            Set s_com=webConn.Execute("SELECT * from Company WHERE Defaulted=1  ")
        end if
        if s_com("BrandLogo")<>"" then
            logo=s_com("BrandLogo")
        end if
        Company=s_com("Company")
        CompanyEmail=s_com("Email")
        BranchAddress=s_com("Address")
        BranchCity = s_com("Suburb")
        BranchPostCode = s_com("State")&"&nbsp;"&s_com("PostCode")&"&nbsp;"&s_com("Country")
        BranchPhone=s_com("Phone")
        BranchFax=s_com("Fax")
        if s_com("WebSite")<>"" then
            WebSite="<a href=http://"&s_com("WebSite")&">"&s_com("WebSite")&"</a>"
            homeURL=s_com("WebSite")
        end if
            FreePhone=""
        if s_com("FreePhone")<>"" then
            FreePhone="&nbsp;&nbsp;Free Phone: "&s_com("FreePhone")&""
        end if
         s_com.close
         set  s_com=nothing
         
           '--check use location or company address---
          Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='DOCS'  ")
          DOCSAddress="Yes"   '---Location Address
          If not RG.EOF then
               DOCSAddress= RG("syValue")
          END IF
          RG.CLOSE
          SET RG=NOTHING
                         
        if DOCSAddress="Yes" then '--use Location address
            Set s_add=webConn.Execute("SELECT  QLocationState.* FROM QLocationState WHERE ID='"&s_d("PickupLocationID")&"'  ")
            if NOT s_add.EOF  then
                        if s_add("Email")<>"" then
                              CompanyEmail=s_add("Email")
                        end if
                        BranchAddress = s_add("Address")
                        BranchCity = s_add("Suburb")&"&nbsp;"&s_add("City")
                        if s_add("StateCode")<>"N/A" then
                           BranchPostCode = s_add("StateCode")&"&nbsp;"&s_add("PostCode")&"&nbsp;"&s_add("Country")
                        else
                           BranchPostCode = s_add("PostCode")&"&nbsp;"&s_add("Country")
                        end if
                        BranchPhone = s_add("Phone")
                        BranchFax = s_add("Fax")
                        '--location may have their own free phone
                         if s_add("FreeCallLocal")<>"" then
                              FreePhone="&nbsp;&nbsp;Free Phone: "&s_add("FreeCallLocal")&""
                         end if
               end if
               s_add.close
               Set s_add=nothing
        end If
        
            '-------------customer details
            Set r_cust=webConn.Execute("Select * from customers where acid='"&s_d("CustomerID")&"'")
            if NOT r_cust.EOF  then
                        Session("RCM273_CustomerEmail") = r_cust("acEmail")
                  Session("RCM273_FirstName")=r_cust("acFirstName")
                  Session("RCM273_LastName")=r_cust("acLastName")
            end if
            r_cust.close
            Set r_cust=nothing

             if s_d("Quotation")=True then

                  ReservationTypeID=3  '--for quotation
                  onlineType="Online Quotation"
                  Subject1="Online Quotation with "&Company&" -  Ref #Q-"&s_d("DocPrFix")&""&s_d("ReservationNo")&" ("&s_d("PickupCity")&")"
            else

                  ReservationTypeID=1  '--for quotation
                  onlineType="Online Booking Request"
                  FreeSale  = 0
                  Set s_w=WebConn.Execute("select ReservationTypeID from WebReservation where UnallocatedRA="&s_d("ReservationNo")&" ")
                  if Not s_w.EOF   then
                        if s_w("ReservationTypeID")=1 then
                        FreeSale =1
                        onlineType="Online Booking"
                        end if
                        'Response.Write s_w("ReservationTypeID")
                  end if
                  s_w.close
                  set s_w=Nothing

                  if FreeSale  = 0  then '---LIMITED AVAILABILTY
                     Subject1="Online Booking Request with "&Company&" - Ref #U-"&s_d("DocPrFix")&""&s_d("ReservationNo")&" ("&s_d("PickupCity")&")"
                  else
                     Subject1="Online Booking with "&Company&" - Ref #U-"&s_d("DocPrFix")&""&s_d("ReservationNo")&" ("&s_d("PickupCity")&")"
                  end if
            end if






        '-----Send a confirmation to customer

         Dim  HTML
         Dim MyCDONTSMail

         Set MyCDONTSMail = CreateObject("CDONTS.NewMail")

         MyCDONTSMail.From = Company&"<"&CompanyEmail&">"
         MyCDONTSMail.To = Session("RCM273_CustomerEmail")
         MyCDONTSMail.Cc = Company&"<"&CompanyEmail&">"
         MyCDONTSMail.Subject=Subject1
         MyCDONTSMail.BodyFormat=0
         MyCDONTSMail.MailFormat=0
         MyCDONTSMail.Body=HTML
         if s_d("Quotation")=False then
            MyCDONTSMail.Importance = 2
         end if

         HTML = "<html>"
         HTML = HTML & "<head>"
         HTML = HTML & "<META http-equiv=Content-Type content='text/html; charset=windows-1252'>"
         HTML = HTML & "<STYLE>"
         HTML = HTML & ".text{  font-size: 11px; line-height: 15px; font-size: 11px;  font-family: Verdana,Arial;}   "
         HTML = HTML & ".white{  font-size: 11px;  font-weight: bold;color: #FFFFFF;  font-family: Verdana,Geneva;}    "
         HTML = HTML & ".greytext{  font-size: 11px;  color:#666666;  font-family: Verdana,Geneva;}    "
         HTML = HTML & ".Yellow{  font-weight: normal;  font-size: 11px;  color: #FDFEBC;  font-family: Verdana,Geneva;}    "
         HTML = HTML & "A.linkBlue{  font-size: 11px;  color: blue;  font-weight: 900; font-family: Verdana,Geneva;  text-decoration: none;}  "
         HTML = HTML & ".HighlightRow {background-color:#EEF1F4;} "
         HTML = HTML & ".TitleRow  {background-color:#376293;}"
         HTML = HTML & ".OpeningTD {     "
         HTML = HTML & "border:solid windowtext 1.0pt;   "
         HTML = HTML & "border-color:"&Session("RCM273_CompanyColour")&"; }  "
         HTML = HTML & "BODY   {      "
         HTML = HTML & "background-color:#FFFFFF;  "
         HTML = HTML & "scrollbar-3dlight-color:#000000; "
         HTML = HTML & "scrollbar-arrow-color:#4D6185;   "
         HTML = HTML & "scrollbar-base-color:#BDD1FB;"
         HTML = HTML & "scrollbar-darkshadow-color:#000000; "
         HTML = HTML & "scrollbar-face-color:#B4BDC5;   "
         HTML = HTML & "scrollbar-highlight-color:#ffffff;"
         HTML = HTML & "scrollbar-track-color:#BDD1FB; "
         HTML = HTML & "scrollbar-shadow-color:#BDD1FB;}"

         HTML = HTML & "</STYLE>"
         HTML = HTML & "<META content='MSHTML 6.00.2900.2722' name=GENERATOR></HEAD>"
         HTML = HTML & "<title>"&Subject1&" </title>"
         HTML = HTML & "</head>"
         HTML = HTML & "<body bgcolor='#FFFFFF'>"
      
         HTML = HTML & "<center><img  src='"&imageURL&"/db/"&CompanyCode&"/"&logo&"' /></center>"
     
         HTML = HTML & "<table style='border-collapse:collapse;  border:solid windowtext 1.0pt; border-color:"&Session("RCM273_CompanyColour")&"; ' cellspacing='0' cellpadding='2' width='600' align='center' 'bgcolor=#ffffff' >"
         HTML = HTML & "<tr style='border:solid windowtext 1.0px;  border-color:"&Session("RCM273_CompanyColour")&"'>"
         HTML = HTML & "<td style='border:solid windowtext 1.0px;height:25px;  border-color:"&Session("RCM273_CompanyColour")&"' bgcolor='"&Session("RCM273_CompanyColour")&"' colspan='2'  class='Yellow'>"

         HTML = HTML & Subject1
         HTML = HTML & "</td></tr>"

          DateEntered = Day(Session("RCM273_LocalTime"))&"/"&MonthName(Month(Session("RCM273_LocalTime")))&"/"&Year(Session("RCM273_LocalTime"))
          if Session("RCM273_LocalTime")="" then
                  DateEntered = Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
          end if

         HTML = HTML & "<tr><td colspan='2' class='text'>"
         if s_d("Quotation")=True  then
               HTML = HTML & "<div align=justify> "

                quotLetter="QuotationLett"&CINT(Session("RCM273_PickupLocationID"))
                If (FileObject.FileExists(FilePath & "\"&quotLetter&".txt" )=true) Then
                     theLetter=FilePath & "\"&quotLetter&".txt"
               else
                     theLetter=FilePath & "\QuotationLett.txt"
               end if
               Set MFile=FileObject.OpenTextFile(theLetter, 1)
               Set objFile = FileObject.GetFile(theLetter)
               If objFile.Size > 0 Then
                     HTML = HTML & MFile.ReadAll
               end if
                  MFile.Close
               Set MFile=nothing
               HTML = HTML & "</pre>"


                   HTML = HTML & "<tr><td colspan='2' class='text'>If you would like to proceed with this quotation :</td></tr>"

                     '---1. need pass CompanyID and ServerID ( not compayCode/DbName)
                     '---2. need pass Quotation No
                     '---3.  need pass s_d("DateEntered")

                     rcmReferenceID=s_d("rcmReferenceID")
                     DateEntered=Day(s_d("DateEntered"))&"/"&Left(MonthName(Month(s_d("DateEntered"))),3)&"/"&Year(s_d("DateEntered"))
                     QuotationNo = s_d("ReservationNo")
                     ValidDateTo=""
                     if s_d("DaysQuoteValid")  <>0 and ISNull(s_d("DaysQuoteValid"))<>True then
                              ValidDateTo=Day(Now+s_d("DaysQuoteValid"))&"/"&Left(MonthName(Month(Now+s_d("DaysQuoteValid"))),3)&"/"&Year(Now+s_d("DaysQuoteValid"))
                     end if
                     HTML = HTML & "<tr ><td class=OpeningTD colspan='2'> "
                     HTML = HTML & "<table><tr><td class='text'><b>TURN YOUR QUOTE INTO A BOOKING REQUEST "
                     HTML = HTML & " <b><a href='"&rcmURL&"/s_QPay.asp?id="&RCMCompanyID&"&MC2="&ServerID&"&ferq="&s_d("ReservationNo")&"&refID="&rcmReferenceID&"&E604="&DateEntered&"' >"
                     HTML = HTML & " <font color='#FF0000'>CLICK HERE</font></a> </b></td><td><img src='"&imageURL&"/images/SafePayment.jpg'  border='0' /></td></tr>"


                     if s_d("DaysQuoteValid")  <>0 then
                              HTML = HTML & "<tr><td class='text'><b>Please note this quote is valid for "&s_d("DaysQuoteValid")&" days. </b></td></tr>"
                     end if
                     HTML = HTML & "<tr><td class='text'  colspan='2'>Note : If you are viewing this message in text mode and having difficulties opening the above link, please try copying and pasting the following entire link into the address bar of your Internet homepage.</td></tr>"
                     HTML = HTML & "<tr><td class='text' colspan='2'>"&rcmURL&"/s_QPay.asp?id="&RCMCompanyID&"&MC2="&ServerID&"&ferq="&s_d("ReservationNo")&"&refID="&rcmReferenceID&"&E604="&DateEntered&"&VD2="&ValidDateTo&"</td></tr>"
                     HTML = HTML & "</table></td></tr>"

                  HTML = HTML & "<tr><td colspan='2' class='text'>Your Quotation details are as follows:</td></tr>"
                  HTML = HTML & "<tr class='HighlightRow' bgcolor='#EEF1F4'><td class=OpeningTD colspan='2'><FONT class='text'>Quotation Date: "



                  HTML = HTML & DateEntered
                  HTML = HTML & "</font></td></tr>"
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor="&Session("RCM273_CompanyColour")&"></td></tr>"
                  HTML = HTML & "<tr ><td class='GREYTEXT' style='width:120px' ><b>Ref:</b></td> <td class='text' align='left' style='width:520px'>#Q-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")</td></tr>"
       else
                    HTML = HTML & "Thank you for your "&onlineType&" with "&Company&"."
                  HTML = HTML & "<tr><td colspan='2' class='text'>Your "&onlineType&" has now been forwarded to the Location - "
                  HTML = HTML & Session("RCM273_PickupLocation")
                  HTML = HTML & ".<tr><td colspan='2' class='text'>You will receive a email confirmation from the renting location shortly."
                  HTML = HTML & "<tr><td colspan='2' class='text'>Your booking details are as follows:</td></tr>"
                  HTML = HTML & "<tr class=HighlightRow bgcolor=#EEF1F4><td class=OpeningTD colspan='2'><FONT class='text'>Booking Date: "
                  HTML = HTML & DateEntered
                  HTML = HTML & "</font></td></tr>"
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                  HTML = HTML & "<tr ><td class='GREYTEXT'><b> Ref:</b></td><td class='text' align='left'><b>#U-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")</b></td></tr>"
                 

        end if

              '-------get booking information
            HTML = HTML & "<tr ><td class='GREYTEXT' >Name: <td class='text' align='left' >"&Session("RCM273_FirstName")&"&nbsp;"&Session("RCM273_LastName")&"</td>"



               HTML = HTML & "<tr ><td class='GREYTEXT' >Email: <td class='text' align='left'>"&Session("RCM273_CustomerEmail")&"</td>"
               HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor="&Session("RCM273_CompanyColour")&"></td></tr>"
               HTML = HTML & "<tr ><td class='GREYTEXT' >Vehicle Type<td class='text' align='left'>"&Session("RCM273_CarType")&"</td>"

                HTML = HTML & "<tr ><td><td class='text' align='left'><img  src='"&theimage&"' width='192' /></td></tr>"
                HTML = HTML & "<tr ><td class='GREYTEXT' >Pickup Location: <td class='text' align='left'>"&Session("RCM273_PickupLocation")&"</td>"
                HTML = HTML & "<tr ><td class='GREYTEXT' >Pickup Date Time: <td class='text' align='left'>"
               HTML = HTML & Day(Session("RCM273_RequestPickDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestPickDate"))),3)&"/"&Year(Session("RCM273_RequestPickDate"))
               HTML = HTML & "&nbsp;"
               HTML = HTML & Session("RCM273_RequestPickTime")
               HTML = HTML & "</td></tr>"

        if Session("RCM273_Flight")<>"" then
               HTML = HTML & "<tr ><td class='GREYTEXT' >Arrival: <td class='text' align='left'>"&Session("RCM273_Flight")&"</td>"
          end if

            HTML = HTML & "<tr ><td class='GREYTEXT' >Dropoff Location: <td class='text' align='left'>"&Session("RCM273_DropoffLocation")&"</td>"
            HTML = HTML & "<tr ><td class='GREYTEXT' >Dropoff Date Time: <td class='text' align='left'>"
            HTML = HTML & Day(Session("RCM273_RequestDropDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestDropDate"))),3)&"/"&Year(Session("RCM273_RequestDropDate"))
            HTML = HTML & "&nbsp;"
            HTML = HTML & Session("RCM273_RequestDropTime")
            HTML = HTML & "</td></tr>"

          if Session("RCM273_Flightout")<>"" then
            HTML = HTML & "<tr ><td class='GREYTEXT' >Departure: <td class='text' align='left'>"&Session("RCM273_Flightout")&"</td>"
          end if

          if Session("RCM273_NoTravelling")<>"" then
            HTML = HTML & "<tr ><td class='GREYTEXT' >Number Travelling: <td class='text' align='left'>"&Session("RCM273_NoTravelling")&"</td>"
         end if

            HTML = HTML & "<tr ><td class='GREYTEXT' >No Days: <td class='text' align='left'>"&Session("RCM273_TotalRentalDays")&"</td>"
            'HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
            HTML = HTML & "<tr class=HighlightRow><td class=OpeningTD colspan='2'><Span class='text'>Rental Rate and Fees</td></tr>"
               
                      
         '----the seasonal rates
         HTML = HTML & "<tr><td colspan='2' >"

         HTML = HTML & "<table>"


              Set s_rate=webConn.Execute("msp_res_DailyRatesBuffer "&s_d("ReservationNo")&"  ")
             CarTotalCost=0
            cost=0


            DO WHILE NOT s_rate.EOF

                         NoDaysDescription=s_rate("Days")&" "&s_d("sRateTypeDesc")&"s @ "&DisplayCurrency(s_rate("Rate"))
                        if s_rate("Days")<=1 and InStr(s_rate("RateName"), "hour")>0 then
                                 NoDaysDescription=s_rate("RateName")
                        end if
                        if   s_rate("NoHours")>0 then '--for hour rate
                              NoDaysDescription=s_rate("RateName")
                        end if
                        if s_rate("DiscountType")="p" then
                                 Discount="(include "&s_rate("DiscountPerc")&"% discount)"
                        else
                                 Discount="(include "&DisplayCurrency(s_rate("DiscountPerc"))&" discount)"
                        end if


                        if s_rate("DiscountPerc")>0 then
                              HTML = HTML & "<tr><td class=greytext>"&s_rate("Season")&":"
                              HTML = HTML & "&nbsp;"&NoDaysDescription&"&nbsp;  "
                              if s_rate("NoHours")=0 or isNull(s_rate("NoHours"))=True then
                               HTML = HTML & "&nbsp;"&s_rate("DiscountName")&"  "&Discount&" Rate @ "&DisplayCurrency(s_rate("Rate"))&""
                               end if
                               HTML = HTML & "</td><td align=right class=text>&nbsp;"
                              HTML = HTML & DisplayCurrency(s_rate("Cost"))
                              HTML = HTML & "</td></tr>"
                        else
                              HTML = HTML & "<tr><td class=greytext>"&s_rate("Season")&":"
                              HTML = HTML & "&nbsp;"&NoDaysDescription&"&nbsp; </td>"
                              HTML = HTML & "<td align=right class=text>&nbsp;"
                              HTML = HTML & DisplayCurrency(s_rate("Cost"))
                              HTML = HTML & "</td></tr>"
                        end if

         s_rate.MoveNext
         Loop
         s_rate.close
         set s_rate=nothing


         if Session("RCM273_AreaofUsed")<>"" then
                      HTML = HTML & "<tr><td class='GREYTEXT'>Area of Use:<td class='text'>"&Session("RCM273_AreaofUsed")&" </td></tr>"
         end if

         '--- extra fees
         ' Set s_extra=WebConn.Execute("msp_res_ExtraFeesBuffer "&s_d("ReservationNo")&"  ")
         Set s_extra=webConn.Execute("SELECT PaymentExtraFeesBuffer.*, Name,ExtraDesc, Type FROM PaymentExtraFeesBuffer, ExtraFees WHERE  (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' and ExtraFeesID = ExtraFees.ID ORDER BY TYPE, NAME ")
         '----- try this
         EachExtraFees=0
         TotalExtraFees=0
         DO WHILE NOT s_extra.EOF
               ExtraValue= DisplayCurrency(s_extra("ExtraValue"))
               ExtraFeesQTY=s_extra("QTY")
               theQTY=""
               if ExtraFeesQTY>1 then
                  theQTY="(QTY "&ExtraFeesQTY&")"
               end if
               ExtraRate=DisplayCurrency(s_extra("Fees"))
               TotalExtraFees = TotalExtraFees + s_extra("ExtraValue")
                if s_d("ReservationTypeID")=4 then
                      if s_extra("CancellationFee")="True" or s_extra("MerchantFee")="True"  then
                                 CancellationFee = CancellationFee + s_extra("ExtraValue")
                                 HTML = HTML & "<tr><td class=greytext>"&s_extra("Name")&" :</td><td align=right class=text>"&ExtraValue&"</td></tr>"
                      end if
                else
                     if s_extra("Type")="Percentage" then
                           HTML = HTML & "<tr><td class=greytext>"&s_extra("Name")&" :</td>"
                     else
                        HTML = HTML & "<tr><td class=greytext>"&s_extra("Name")&",  "&s_extra("Type")&" @ "&ExtraRate&" "&theQTY&":</td>"
                     end if

                     HTML = HTML & "<td align=right class=text>"&ExtraValue&"</td></tr>"
                     if s_extra("ExtraDesc")<>"" then
                        HTML = HTML & "<tr><td class=greytext colspan=2>"&s_extra("ExtraDesc")&" </td></tr>"
                     end if
                end if

         s_extra.MoveNext
         Loop
         s_extra.close
         set s_extra=nothing

             if  s_d("StampDuty") >0 then
                  HTML = HTML & "<tr><td class='GREYTEXT'>"
                  HTML = HTML & Session("RCM273_TaxName2")
                  HTML = HTML & "</td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""
                  HTML = HTML & FormatNumber(s_d("StampDuty"),2)
                  HTML = HTML & "</td></tr>"
             end if

              TotalCost=s_d("TotalCost")

               if GSTInclusive = "Yes"  then
                        HTML = HTML & "<tr><td class='GREYTEXT'><b>Estimate of Charges:</b></td><td  align=right class='text'>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""
                        HTML = HTML & FormatNumber(TotalCost,2)
                        HTML = HTML & "</td></tr>"
                        HTML = HTML & "<tr><td class='GREYTEXT' colspan='2' align=right>(All Prices GST Inclusive)"
                        HTML = HTML & "</td></tr>"
                 else
                        if s_d("GST")>0 then
                        HTML = HTML & "<tr><td class='GREYTEXT'>"&Session("RCM273_TaxName1")&"</b></td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""
                        HTML = HTML & FormatNumber(s_d("GST"),2)
                        HTML = HTML & "</td></tr>"
                        end if
                        HTML = HTML & "<tr><td class='GREYTEXT'><b>Estimate of Charges:</b></td><td  align=right class='text'>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""
                        HTML = HTML & FormatNumber(TotalCost,2)


                        HTML = HTML & "</td></tr>"
                end if

             IF s_d("Quotation")=False then
               '----- payment
                  PaidAmount=0
                 Set s_p=WebConn.Execute("msp_res_PaymentBuffer "&s_d("ReservationNo")&"  ")
               While Not s_p.EOF
                           HTML = HTML & "<tr><td class=greytext>"
                           if s_p("PaymentType")<>"Agent Collected" then
                                 HTML = HTML & "Paid by "
                           end if
                           HTML = HTML & ""&s_p("PaymentType")&""&s_p("PaymentRefNo")&" ("&s_p("sPaymentDate")&")</td>"
                           HTML = HTML & "<td class=text align=right>"&DisplayCurrency(s_p("Paid"))&"</td></tr>"
                          if s_p("PaymentType")="Agent Collected" then
                           AgentCollected=s_p("Paid")
                          end if
             s_p.MoveNext
               Wend
               s_p.close
               set s_p=nothing

               Balance=s_d("BalanceDue")
               HTML = HTML & "<tr><td class=greytext><b>Balance Due:</td><td class=text align=right><b>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(Balance,2)&"</td></tr>"

            end if


         HTML = HTML & "</table>"
         HTML = HTML & "</td></tr>"
                 
      if s_d("Notes") <>"" then
            HTML = HTML & "<tr><td colspan='2' >&nbsp;</td></tr>"
            HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>Comments/Requests:</td></tr>"
            HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"&s_d("Notes")&"</td></tr>"
       end if   


      IF s_d("Quotation")=True  then

               '---if pickuplocation quotation text exist, use it,
                  quotFile="Quotation"&CINT(Session("RCM273_PickupLocationID"))

                  'Response.Write(FilePath)
               If (FileObject.FileExists(FilePath & "\"&quotFile&".txt" )=true) Then
                   FileName=FilePath & "\"&quotFile&".txt"
               else
                   FileName=FilePath+"\Quotation.txt"
               end if
                  Set MFile=FileObject.OpenTextFile(FileName, 1)
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                  HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"
                  HTML = HTML & MFile.ReadAll
                  HTML = HTML & "</td></tr>"
                  MFile.Close
                  Set MFile=nothing

      ELSE
                '---if pickuplocation confirmation text exist, use it,
               EmailText="EmailConfirmation"&CINT(Session("RCM273_PickupLocationID"))
              
                'Response.Write(FilePath)
               If (FileObject.FileExists(FilePath & "\"&EmailText&".txt" )=true) Then
                   FileName=FilePath & "\"&EmailText&".txt"
               else
                   FileName=FilePath+"\EmailConfirmation.txt"
               end if
                 ' Set MFile=FileObject.OpenTextFile(FileName, 1)
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                  HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"
                 ' HTML = HTML & MFile.ReadAll
                  HTML = HTML & "</td></tr>"
                  'MFile.Close
                  'Set MFile=nothing

      END IF

         HTML = HTML & "<tr><td class='text' colspan='2'><br><b>"&Company&"</b></br></td></tr>"
         HTML = HTML & "<tr><td class='text' colspan='2'>"&Session("RCM273_Pickuplocation")&"</td></tr>"
    
         HTML = HTML & "<tr><td style='border:solid windowtext 1.0px;  border-color:"&Session("RCM273_CompanyColour")&"' bgcolor='"&Session("RCM273_CompanyColour")&"' align=center colspan='2'  class=Yellow>"
         HTML = HTML & BranchAddress
         HTML = HTML & "<br>"
         HTML = HTML & BranchCity
         HTML = HTML & ",&nbsp;"
         HTML = HTML & BranchPostCode
         HTML = HTML & "</br>"
         HTML = HTML & "<br>Phone:"
           HTML = HTML & BranchPhone
           HTML = HTML & "&nbsp;&nbsp;&nbsp;Fax: "
           HTML = HTML & BranchFax
           HTML = HTML & FreePhone
           HTML = HTML & "</br>"
          HTML = HTML & "</td></tr>"
          HTML = HTML & "</table><center>"
          HTML = HTML & WebSite
            HTML = HTML & "</center>"


       MyCDONTSMail.Body=HTML
       MyCDONTSMail.Send
       set MyCDONTSMail=Nothing
       
      'Response.write "<br/><br/>" & HTML




        '---print out on the screen

        Response.Write("<table   align='center' width='720'  bgcolor='"&Session("RCM273_CompanyColour")&"' cellspacing='1' cellpadding='1'  border='0'>")
        Response.Write("<tr><td>")
      Response.Write("<table width='100%' align='center' bgcolor='#FFFFFF'  cellspacing='0' cellpadding='2'  border='0'>")
      Response.Write("<tr><td  style='height:20px' align='center' class='header' ><h4 class='smallm_title centered bigger'><span>Thank you for your "&onlineType&"</span></h4></td></tr>")
      
        
      if FreeSale  = 1 then
            Response.Write("<tr><td>&nbsp;</td></tr>")
            if s_d("Quotation")=True  then
                  Response.Write("<tr><td align='center' class='text'>This quote  has been forwarded to your email</td></tr>")
                  Response.Write("<tr><td>&nbsp;</td></tr>")
                  Response.Write("<tr><td class='text' align='center' >Any additional info, please contact us at <a href='mailto:"&CompanyEmail&"'  class='re'd><font color='#3cb4e8'>"&CompanyEmail&"</font></a></h1></td></tr>")
                  Response.Write("<tr><td align='center' class='text'>Thanks for choosing "&Company&"..</td></tr> ")
              Response.Write("<tr><td>&nbsp;</td></tr>")
       
             else
                  Response.Write("<tr><td align='center' class='text'>Your Online Booking  has been forwarded to "&Company&"</td></tr>")
                  Response.Write("<tr><td align='center' class='text'>You will be contacted shortly with pick-up details and confirmation.</td></tr>")
                  Response.Write("<tr><td align='center' class='text'>An Online Booking confirmation has also been emailed to you.</td></tr>")

                  Response.Write("<tr><td align='center' class='text'>If you have any queries regarding the status of your Online Booking please don't hesitate to email us at <a href='mailto:"&CompanyEmail&"'  class='red'><font color='#3cb4e8'>"&CompanyEmail&"</font></a></td></tr>  ")
                  Response.Write("<tr><td align='center' class='text'>Thanks for choosing  "&Company&"..</td></tr> ")
              Response.Write("<tr><td>&nbsp;</td></tr>")
           end if
      else
            if s_d("Quotation")=True  then
                     Response.Write("<tr><td align='center' class='text'>This quote  has been forwarded to your email</td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")
                     Response.Write("<tr><td class='text' align=center >Any additional info, please contact us at <a href='mailto:"&CompanyEmail&"'  class=red><font color='#3cb4e8'>"&CompanyEmail&"</font></a></h1></td></tr>")
                    Response.Write("<tr><td  align='center' class='text'>Thanks for choosing "&Company&"..</td></tr> ")
               Response.Write("<tr><td>&nbsp;</td></tr>")
       
             else
                     Response.Write("<tr><td class='text' align=center>Your booking request has been forwarded to "&Company&" </td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")
                     Response.Write("<tr><td class='text' align=center >As this vehicle is heavily booked during this period you will be contacted shortly to confirm availability.</h1></td></tr>")
                     'Response.Write("<tr><td class='text' align=center >Please note: No charge will be made against your credit card until your reservation is confirmed.</h1></td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")

                    Response.Write("<tr><td align='center' class='text'>If you have any queries regarding the status of your Online Booking Request please don't hesitate to email us at <a href='mailto:"&CompanyEmail&"'  class=red><font color='#3cb4e8'>"&CompanyEmail&"</font></a></td></tr>  ")
                     Response.Write("<tr><td  class='text'></td></tr>")
                     Response.Write("<tr><td align='center' class='text'>Thanks for choosing  "&Company&".</td></tr> ")
                Response.Write("<tr><td>&nbsp;</td></tr>")
            end if
       end if
    Response.Write("<tr><td  class='text' align='center'><A HREF=http://"&homeURL&">RETURN TO HOME PAGE</a> ")
    '---close window dose not work in some browser 
 ' Response.Write("<a href=""javascript:window.close()"" >Close Window</a> ")
  '  Response.Write("<form method='post' action='webstep5.asp'  >")

  ' Response.Write("<input type='rcmbutton' name='Back' class=rcmbutton value='Close Window' onclick='javascript:window.close()'>")
  'Response.Write("</form>")
    Response.Write("</td></tr> ")
    Response.Write("<tr><td style='height:40px'>&nbsp;</td></tr>")
    Response.Write("</table>")
    Response.Write("</td></tr>")
    Response.Write("</table> ")
    
    Session("RCM273_OnlineBookingNo")=""
    s_d.close
  set s_d=nothing
    Session.Abandon

    webConn.CLOSE
    SET webConn=nothing

else
       s_d.close
      set s_d=nothing
      webConn.CLOSE
     SET webConn=nothing
      Session.Abandon
     Response.Redirect "webstep2.asp"

end if


%>



</div>
<!-- END RCM HTML CODE-->

<!-- #include file="include_footer.asp" -->

</body>
</html>
 
