<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->


<%Response.AddHeader "p3p", "CP=\""IDC DSP COR ADM DEVi TAIi PSA PSD IVAi IVDi CONi HIS OUR IND CNT\"%>


<!-- #include file="include_meta.asp" -->

</head>
<body class="webstep1">

<%
         Session("RCM273_CompanyColour")="#0080C0" '---the booking form color is blue, web designer can change it here
         CompanyCode="PtWestCoastCampers273"
         RCMURL="https://secure2.rentalcarmanager.com/ssl/PtWestCoastCampers273/webstep1.asp"

         Dim webConn
         Set webConn = Server.CreateObject("ADODB.Connection")
         DatabaseFile="PtWestCoastCampers273"
         webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"

'---- +++++++++++++++++++++++++++++++++++++++++++++++
'---- Module Name : Web Site Interface
'---- File Name :  webbookingstep1.asp
'---- Method Name :  SelectPickupLocation
'---- Method Description :  Selects Pickup Location
'---- Tables Used : Location
'---- +++++++++++++++++++++++++++++++++++++++++++++++

SUB SelectPickupLocation
      Response.Write("<select class='form-control input-sm' name='PickupLocationID' >")

      Set l_s = webConn.Execute("select * FROM Location WHERE PickupAvailable=1 order by location")
      While Not l_s.EOF
               if Session("RCM273_PickupLocationID")="" and l_s("WebDefault") =True then
                   Response.Write("<option value='"&l_s("ID")&"' selected='selected'>"&l_s("Location")&" &nbsp;</option>")
               elseif CStr(l_s("ID"))=Session("RCM273_PickupLocationID")  then
                  Response.Write("<option value='"&l_s("ID")&"' selected='selected'>"&l_s("Location")&" &nbsp;</option>")
               else
                  Response.Write("<option value='"&l_s("ID")&"' >"&l_s("Location")&" &nbsp;</option>")
               end if
        l_s.MoveNext
        Wend
        l_s.Close
        Set l_s=nothing
        Response.Write("</select>")
END SUB

'---- +++++++++++++++++++++++++++++++++++++++++++++++
'---- Module Name : Web Site Interface
'---- File Name :  webbookingstep1.asp
'---- Method Name :  SelectDropoffLocation
'---- Method Description :  Selects Dropoff Location
'---- Tables Used : Location
'---- +++++++++++++++++++++++++++++++++++++++++++++++
SUB SelectDropoffLocation
      Set l_s = webConn.Execute("select * FROM Location WHERE DropoffAvailable=1  order by location")
      Response.Write("<select class='form-control input-sm' name='DropoffLocationID'  >")

      Response.Write("<option value='Same' >Same location</option>")
      While Not l_s.EOF
            if CStr(l_s("ID"))=Session("RCM273_DropoffLocationID")  then
                  Response.Write("<option value='"&l_s("ID")&"' selected='selected'>"&l_s("Location")&" &nbsp;</option>")
            else
                  Response.Write("<option value='"&l_s("ID")&"' >"&l_s("Location")&" &nbsp;</option>")
            end if
      l_s.MoveNext
      Wend
      l_s.Close
      Set l_s=nothing
      Response.Write("</select>")
END SUB

'---- +++++++++++++++++++++++++++++++++++++++++++++++
'---- Module Name : Web Site Interface
'---- File Name :  webbookingstep1.asp
'---- Method Name :  TheQutoationForm
'---- Method Description :  Form to accept Pickup and Drop Off Location, Time and Age of Driver
'---- Tables Used : Location
'---- +++++++++++++++++++++++++++++++++++++++++++++++
SUB TheQutoationForm

        Session("RCM273_ErrorMesage")=""

         if Session("RCM273_RequestPickTime")<>"" then
                 defaultPickupTime= Session("RCM273_RequestPickTime")
         else
                 defaultPickupTime="11:00"
         end if
         if Session("RCM273_RequestDropTime")<>"" then
                defaultDrpoffTime= Session("RCM273_RequestDropTime")
         else
                defaultDrpoffTime="10:00"
         end if

         if Session("RCM273_RequestPickDate")<>"" then
                defaultPickupDate = Session("RCM273_RequestPickDate")
         else
                defaultPickupDate=Date+6
         end if

         if Session("RCM273_RequestDropDate")<>"" then
                defaultDropoffDate = Session("RCM273_RequestDropDate")
         else
                defaultDropoffDate=Date+16
         end if

          Dim PickupDateArray
          'PickupDateArray = Split(defaultPickupDate, "/")
          PickupDateMonth = month(defaultPickupDate)
          PickupDateDay   = day(defaultPickupDate)
          PickupDateYear  = year(defaultPickupDate)

          if len(PickupDateDay)<2 then 
            PickupDateDay="0"&PickupDateDay
          end if

          if len(PickupDateMonth)<2 then 
            PickupDateMonth="0"&PickupDateMonth
          end if

          dtPickUp  = PickupDateMonth&"/"&PickupDateDay&"/"&PickupDateYear
          'Response.write "dtPickUp= "&dtPickUp&"<br>"

%>

<!--if  webstep1.asp is a iframe inside home page, use  below form open a popup window to go to webstep2.asp-->
<!--<form method=Post action='webstep2.asp?refid=<%=Request.QueryString("refid")%>&amp;URL=<%=Request.QueryString("URL")%>' onsubmit="openTarget(this,'resizable=1,scrollbars=1,height=950,width=828'); return true;" id="theform" >-->
<!--if  webstep1.asp is not a iframe inside home page, use  below form -->
<!-- <form method="post" name="theform" action="webstep2.asp?refid=<%=Request.QueryString("refid")%>&amp;URL=<%=Request.QueryString("URL")%>" id="theform" target="_parent"> -->

<!--iframe inside home page, Validate and open a popup window to go to webstep2.asp,-->
 <form method="post" name="theform"  action="webstep2.asp?refid=<%=Request.QueryString("refid")%>&amp;URL=<%=Request.QueryString("URL")%>#book" id="theform"  onsubmit="javascript:return Validate();" rel="nofollow" target="_parent" class="form-inline"/>

<div class="row">
  <div class="col-sm-5" style="padding-top:20px;">
    <h4 class="text-center">Check our prices and book your campervan!</h4>
  </div>


<div class="col-sm-7">
<div class="row">
<div class="form-group">
  
    <label class="control-label">Pickup Location:</label>
    <%    SelectPickupLocation  %>
    <label class="sr-only" for="PickupDay">Pickup Day</label>
    <select name="PickupDay" id="PickupDay" class="form-control input-sm">
      <%     for pd=1 to 9
                     zero = "0"
                     if  pd=Day(DefaultPickupDate)   then
                              Response.Write("<option value='")
                              Response.Write(""& zero & pd &"' selected='selected'>"&pd&"</option>")
                     else
                              Response.Write("<option value='"& zero & pd &"' >"&pd&"</option>")
                     end if
                     zero = ""
                     next
                     for pd=10 to 31
                           if  pd=Day(DefaultPickupDate) then
                                 Response.Write("<option value='")
                                 Response.Write(""& pd &"' selected='selected'>"&pd&"</option>")
                           else
                                 Response.Write("<option value='"& pd &"' >"&pd&"</option>")
                           end if
                           zero = ""
                     next
           %>  </select>

            <label class="sr-only" for="PickupMonth">Pickup Month</label>
            <select name="PickupMonth" id="PickupMonth" class="form-control input-sm">
                  <% for pm = 1 to 9
               theMonth=Left(MonthName(pm),3)
               monthValue="0"&pm
                           if pm=Month(defaultPickupDate)    then
                                 Response.Write("<option value='"&monthValue&"' selected='selected'>"&theMonth&"</option>")
                           else
                                 Response.Write("<option value='"&monthValue&"' >"&theMonth&"</option>")
                           end if
                  next
                  for pm = 10 to 12
                  theMonth=Left(MonthName(pm),3)
                           if pm=Month(defaultPickupDate)    then
                                 Response.Write("<option value='"&pm&"' selected='selected'>"&theMonth&"</option>")
                           else
                                 Response.Write("<option value='"&pm&"' >"&theMonth&"</option>")
                           end if

                  next %>
                  </select>

                 <label class="sr-only" for="PickupYear">Pickup Year</label>
                   <select name="PickupYear" id="PickupYear" class="form-control input-sm">
           <%   for theYear=(Year(Date)-1) to (Year(Date)+2)
                     if theYear=Year(defaultPickupDate)  then
                           Response.Write("<option value='"&theYear&"' selected='selected'>"&theYear&"</option>")
                     else
                           Response.Write("<option value='"&theYear&"'>"&theYear&"</option>")
                     end if
                next
                %> </select>
                   
                   <%'=day(defaultPickupDate) %>
      <input type="hidden" value="<% if len(month(defaultPickupDate))<2 then Response.Write("0"&month(defaultPickupDate)) else Response.Write(month(defaultPickupDate)) end if%><%="/" & day(defaultPickupDate) & "/" & year(defaultPickupDate)%>" id="txtStartDate" size="12" />

            
            <label class="sr-only" for="PickupYear">Pickup Time</label>
            <select name="pickupTime" id="pickupTime"  class="form-control input-sm hidden">
            <option value='00:00' >midnight</option>
            <option value='00:30' >00:30 AM</option>
                     <%
                      for i=1 to 23
                          for j=1 to 2
                              if j=1  then
                                 min="00"
                               else
                                 min="30"
                               end if
                               if  i<12 then
                                 test=i&":"&min&" AM"
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
                                 Response.Write("<option value='"&theTime&"' selected='selected'>"&test&"</option>")
                               else
                                 Response.Write("<option value='"&theTime&"' >"&test&"</option>")

                               end if
                           next
                      next
                %>
      </select>
</div><!-- form-group -->
</div>

<div class="row">                   
<div class="form-group"> 
      <label class="control-label">Return Location:</label>
      <%   SelectDropoffLocation    '--Subs         %>

      <label class="sr-only" for="DropoffDay">Return Day</label>
        <select name="DropoffDay" id="DropoffDay"  class="form-control input-sm">
           <%  for pd=1 to 9
                     zero = "0"
                     if  pd=Day(DefaultDropoffDate)   then
                           Response.Write("<option value='")
                              Response.Write(""& zero & pd &"' selected='selected'>"&pd&"</option>")
                           else
                                    Response.Write("<option value='"& zero & pd &"' >"&pd&"</option>")
                           end if
                     zero = ""
                  next
                     for pd=10 to 31
                           if  pd=Day(DefaultDropoffDate) then
                        Response.Write("<option value='")
                        Response.Write(""& pd &"' selected='selected'>"&pd&"</option>")
                     else
                              Response.Write("<option value='"& pd &"' >"&pd&"</option>")
                           end if
               zero = ""
            next
           %>     </select>
       
       <label class="sr-only" for="DropoffMonth">Return Month</label>
       <select name="DropoffMonth" id="DropoffMonth"  class="form-control input-sm">
         <%

          for pm = 1 to 9
               theMonth=Left(MonthName(pm),3)
               monthValue="0"&pm
                          if  pm=Month(defaultDropoffDate)   then
                        Response.Write("<option value='"&monthValue&"' selected='selected'>"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&monthValue&"' >"&theMonth&"</option>")
                                end if
                  next
                  for pm = 10 to 12
               theMonth=Left(MonthName(pm),3)
                         if  pm=Month(defaultDropoffDate)   then
                        Response.Write("<option value='"&pm&"' selected='selected'>"&theMonth&"</option>")
                           else
                              Response.Write("<option value='"&pm&"' >"&theMonth&"</option>")
                                end if

             next
           %>  </select>
  
  <label class="sr-only" for="DropoffYear">Return Year</label>
   <select name="DropoffYear" id="DropoffYear"  class="form-control input-sm">
          <%    for theYear=(Year(Date)-1) to (Year(Date)+2)
                     if   theYear=Year(defaultDropoffDate)  then
                               Response.Write("<option value='"&theYear&"' selected='selected'>"&theYear&"</option>")
                     else
                           Response.Write("<option value='"&theYear&"'>"&theYear&"</option>")
                     end if
                  next    %>
            </select>
           

<%'=day(defaultDropoffDate) %>
            <input type="hidden" value="<% if len(month(defaultDropoffDate))<2 then Response.Write("0"&month(defaultDropoffDate)) else Response.Write(month(defaultDropoffDate)) end if%>/<%if len(day(defaultDropoffDate)) < 2 then Response.Write("0"&day(defaultDropoffDate)) else Response.Write(day(defaultDropoffDate)) end if%>/<%=year(defaultDropoffDate)%>" id="txtEndDate" size="12" />
     
        <label class="sr-only" for="DropoffTime">Return Time</label>
       <select name="DropoffTime"   class="form-control input-sm hidden">
            <option value='00:00' >midnight</option>
                      <option value='00:30' >00:30 AM</option>
                     <%
                     for i=1 to 23
                           for j=1 to 2
                              if j=1  then
                                 min="00"
                               else
                                 min="30"
                               end if
                               if  i<12 then
                                 test=i&":"&min&" AM"
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
                                 Response.Write("<option value='"&theTime&"' selected='selected'>"&test&"</option>")
                               else
                                 Response.Write("<option value='"&theTime&"' >"&test&"</option>")

                               end if
                           next
                      next
                %>

            </select>
    


           <%   Set s_c = webConn.Execute("SELECT count(*) as SizeCount FROM CategoryType WHERE RentingType=1  ")

               if s_c("SizeCount") >1 then
                        Response.Write("<tr><td class=""text"" align=""left""><font color=red>*</font>Vehicle Type:</td>")
                        Response.Write("<td align=left valign=""bottom"" >")
                        Set l_s = webConn.Execute("SELECT * FROM CategoryType WHERE RentingType=1  order by Ordering")
                        Response.Write("<SELECT name='CategoryTypeID'>")

                        'Response.Write("<option value='0' >Choose Vehcile Type</option>")
                        While Not l_s.EOF
                              if CStr(l_s("ID"))=Session("RCM273_CategoryTypeID")  then
                                    Response.Write("<option value='"&l_s("ID")&"' selected>"&l_s("CategoryType")&" &nbsp;</option>")
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
               Response.Write("<input name=CategoryTypeID type=hidden  Value="&s_c("ID")&"  />")
               end if
               s_c.close
               set s_c=nothing
              %>

&nbsp;<input name="submit" type="submit" class="rcmbutton btn btn-success" value="GET A QUOTE"  />
</div><!-- form-group -->
</div>
<div class="row">                   
<div class="form-group"> 

         <label class="">Youngest Driver:</label>
       
        <%
            Response.Write("<SELECT name='driverage' class='form-control input-sm' style='width: 138px;'><option value=0 selected>Please select</option>")
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
                           Response.Write("<option value='"&s_m("DriverAge")&"' selected>"&DriverAge&"</option>")
                  elseif Session("RCM273_driverage")="" and s_m("DefaultAge")="True" then
                           Response.Write("<option value='"&s_m("DriverAge")&"' selected>"&DriverAge&"</option>")
                  else
                           Response.Write("<option value='"&s_m("DriverAge")&"' >"&DriverAge&"</option>")
                  end if
        s_m.MoveNext
        Wend
        s_m.Close
        Set s_m=nothing  %>
                   </SELECT>
        
        <label class="sr-only">Promo code?: </label>
       <input type=hidden name=PromoCode maxlength=30 size=20 value="<%=Session("RCM273_PromoCode")%>">
      <!--    <tr><td class="text" align="left"><font color=red>&nbsp; *</font> <strong>First Name:</td>
       <td class="text" align="left"><input type="text" name="firstname" maxlength="30" size="30" value="<%=Session("RCM273_FirstName")%>" /></td></tr>
       <tr><td class="text" align="left"><font color=red>&nbsp; *</font> <strong>Last Name:</td>
        <td class="text" align="left"><input type="text" name="lastname" maxlength="30" size="30" value="<%=Session("RCM273_LastName")%>" /></td></tr>
       <tr><td class="text" align="left"><font color=red>&nbsp; *</font> <strong>Phone (incl area code):</td>
      <td class="text" align="left"><input type="text" name="Phone" maxlength="20" size="30" value="<%=Session("RCM273_phone")%>" /></td></tr>
      <tr><td class="text" align="left"><font color=red>&nbsp; *</font> <strong>Email Address:</td>
      <td class="text" align="left"><input type="text" name="CustomerEmail" maxlength="50" size="30" value="<%=Session("RCM273_CustomerEmail")%>" /></td></tr>
       -->
       

</div><!-- form-group -->



</div><!-- row -->
</div><!-- col-xs-5 -->

</div><!-- row -->
 </form>
 <%  END SUB       %>

<!-- RCM HTML CODE-->




  <%  if Session("RCM273_ErrorMesage")<>"" then
      Response.Write("<div class=""alert alert-danger"" role=""alert""> ")
            Response.Write(""&Session("RCM273_ErrorMesage")&"")
            Session("RCM273_ErrorMesage")=""

            Response.Write(" </div>")
   end if   %>

<%      TheQutoationForm
         webConn.CLOSE
   SET webConn=nothing               %>



 </div> <!-- END - Comntainer-fluid -->
<!-- END RCM HTML CODE-->



</body>
</html>




