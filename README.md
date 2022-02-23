# ShippingTimeQuote
Automated shipping time quotation from 4 major carriers and summarize the returned data into a table.

**I will not be responsible for any violation of the Terms of Use of the corresponding websites. Use at your own risk.**

This is a simple VBA script to gather international shipping time information from 4 major carriers (UPS/FedEx/DHL/TNT). It has the following main features:

1. A combination generator written by PowerQuery M to generate a list of origin-destination pairs, shipping dates and shipment dimensions.
2. A crawler written in VBA (with the exception of DHL, which is in PowerQuery M because of library compatibility issue) which will call the corresponding carriers' web service API and store the return in the Excel file.
3. Data transformation written in PowerQuery M to transform the corresponding JSON/XML/HTML data into similar format for comparison.

Please note the following:
1. Usually the web services will not support "date-back" of the ship date. i.e. ship date is earlier than current day. Please test with the corresponding website for the acceptable date range.
2. Different carriers may have different city name and postal code format. Please test with the corresponsding website to see if your address(es) can be accepted.

Data source:

UPS - https://wwwapps.ups.com/time (HTML)

FedEx - https://www.fedex.com/en-us/online/rating.html (JSON)

DHL - https://dct.dhl.com/ (XML)

TNT - https://www.tnt.com/express/en_sg/site/get-quote.html (JSON)
