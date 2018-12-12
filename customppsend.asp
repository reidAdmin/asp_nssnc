<%
' This is an example of how you would go about setting up a custom payment 
' provider for the Ecommerce Plus template range. More information can be found
' at http://www.ecommercetemplates.com
' Here we have used the 2Checkout.com system as an example of how a common payment
' processor works. You can edit this file to match the details of your particular payment system
' Other useful parameters are countryCode, shipCountryCode, countryCurrency
' Firstly you will need to set the URL to pass payment variables below in the FORM action %>
<form method="post" action="https://www.2checkout.com/cgi-bin/sbuyers/cartpurchase.2c">
<% ' A unique id is assigned to each order so that we can track the order. This is available as the orderid. Edit the name cart_order_id to that which is used by your payment system. %>
	<input type="hidden" name="cart_order_id" value="<%=orderid%>" />
<% ' In the Ecommerce Templates admin section for the Custom Payment System, up to 2 pieces of data can be entered %>
<% ' to configure a payment system. These are Data 1 and Data 2 and are available in the variables data1 and data2 %>
	<input type="hidden" name="sid" value="<%=data1%>" />
<% ' Our example of 2Checkout.com does not require a return URL, but I´ve included one below as an example if needed %>
	<input type="hidden" name="returnurl" value="<%=storeurl%>thanks.asp" />
<% ' The variable ppmethod is available if needed to choose between authorize only and authorize capture payments. If this does not apply to your payment system just delete the line below %>
	<input type="hidden" name="paymenttype" value="<% if ppmethod=1 then response.write "1" else response.write "0" %>" />
<% ' The following should be quite self explanatory %>
	<input type="hidden" name="total" value="<%=grandtotal%>" />
	<input type="hidden" name="card_holder_name" value="<%=ordName%>" />
	<input type="hidden" name="street_address" value="<%=ordAddress%>" />
	<input type="hidden" name="city" value="<%=ordCity%>" />
	<input type="hidden" name="state" value="<%=ordState%>" />
	<input type="hidden" name="zip" value="<%=ordZip%>" />
	<input type="hidden" name="country" value="<%=ordCountry%>" />
	<input type="hidden" name="email" value="<%=ordEmail%>" />
	<input type="hidden" name="phone" value="<%=ordPhone%>" />
<%	if ordShipName <> "" OR ordShipAddress <> "" then %>
	<input type="hidden" name="ship_name" value="<%=ordShipName%>" />
	<input type="hidden" name="ship_street_address" value="<%=ordShipAddress%>" />
	<input type="hidden" name="ship_city" value="<%=ordShipCity%>" />
	<input type="hidden" name="ship_state" value="<%=ordShipState%>" />
	<input type="hidden" name="ship_zip" value="<%=ordShipZip%>" />
	<input type="hidden" name="ship_country" value="<%=ordShipCountry%>" />
<%	end if %>
<%	' A variable "demomode" is made available to the admin section that signals the payment method is in demo mode
	if demomode then Response.write "<input type=""hidden"" name=""demo"" value=""Y"" />"
	' IMPORTANT NOTE ! You may notice there is not closing </form> tag. This is intentional. %>