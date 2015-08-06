Attribute VB_Name = "Module2"
Sub DeepSearch()

'change this to your zillow id
' Zillow Web Service ID
ZWSID = "X1-ZWz1b36ikuko3v_1is62"

' Number of header columns
Headers = 1

' Columns containing addresses
' A is used for the reference no
Address = "B"
City = "C"
State = "D"
Zip = "E"

' Columns to return data
County = "F"
LotSQFT = "G"
Landuse = "H"
YearBuilt = "I"
Bedrooms = "J"
Bathrooms = "K"
SQFT = "L"
LastSoldDate = "M"
lastSoldPrice = "N"
Listed = "O"
ListedPrice = "P"
Comparables = "Q"
Zestimate = "R"
ErrorMessage = "S"
zpropID = "T"



Dim xmldoc As MSXML2.DOMDocument60
Dim xmlNodeList As MSXML2.IXMLDOMNodeList
Dim myNode As MSXML2.IXMLDOMNode
Dim WS As Worksheet: Set WS = ActiveSheet


' Tell user the code is running
Application.StatusBar = "Starting search"

' Count Rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' For loop to iterate for each row
    ' starts at row after header
    For I = Headers + 1 To LastRow
    
        'Clear previous data from cells
        WS.Range(County & I) = ""
        WS.Range(LotSQFT & I) = ""
        WS.Range(Landuse & I) = ""
        WS.Range(YearBuilt & I) = ""
        WS.Range(Bedrooms & I) = ""
        WS.Range(Bathrooms & I) = ""
        WS.Range(SQFT & I) = ""
        WS.Range(LastSoldDate & I) = ""
        WS.Range(lastSoldPrice & I) = ""
        WS.Range(Listed & I) = ""
        WS.Range(ListedPrice & I) = ""
        WS.Range(Comparables & I) = ""
        WS.Range(Zestimate & I) = ""
        WS.Range(ErrorMessage & I) = ""
        WS.Range(zpropID & I) = ""
        
        ' Create Zillow API URL
        rowAddress = WS.Range(Replace(Address, " ", "+") & I)
        rowCity = WS.Range(City & I)
        rowState = WS.Range(State & I)
        rowZip = WS.Range(Zip & I)
        
        ' Zillow URL
        URL = "http://www.zillow.com/webservice/GetDeepSearchResults.htm?zws-id=" & ZWSID & "&address=" & rowAddress & "&citystatezip=" & rowCity & "%2C+" & rowState & "%2C+" & rowZip & "&rentzestimate=false"
        
                   
         ' notify status bar of current task
         Application.StatusBar = "Retrieving: " & I & " of " & LastRow - Headers & ": " & rowAddress & ", " & rowCity & ", " & rowState
        
        'Open XML page
        Set xmldoc = New MSXML2.DOMDocument60
        xmldoc.async = False
        
        ' Check XML document is loaded
        If xmldoc.Load(URL) Then
        
            
            Set xmlMessage = xmldoc.SelectSingleNode("//message/text")
            Set xmlMessageCode = xmldoc.SelectSingleNode("//message/code")
            
            ' Check for an error message
            If xmlMessageCode.Text <> 0 Then
            
                ' Return error message
                WS.Range(ErrorMessage & I) = xmlMessage.Text
                
            Else
                'get the county
                Set xmlCounty = xmldoc.SelectSingleNode("//response/results/result/FIPScounty")
                If xmlCounty Is Nothing Then
                    WS.Range(County & I) = "No County Information Available"
                Else
                    WS.Range(County & I) = xmlCounty.Text
                End If
                    
                'get LotSQFT
                Set xmlLotSQFT = xmldoc.SelectSingleNode("//response/results/result/lotSizeSqFt")
                If xmlLotSQFT Is Nothing Then
                    WS.Range(LotSQFT & I) = "No Lot Size Information Available"
                Else
                    WS.Range(LotSQFT & I) = xmlLotSQFT.Text
                End If
                
                'get landuse
                Set xmlLanduse = xmldoc.SelectSingleNode("//response/results/result/useCode")
                If xmlLanduse Is Nothing Then
                    WS.Range(Landuse & I) = "No Land Use Information Available"
                Else
                    WS.Range(Landuse & I) = xmlLanduse.Text
                End If
                
                'get year built
                Set xmlYearBuilt = xmldoc.SelectSingleNode("//response/results/result/yearBuilt")
                If xmlYearBuilt Is Nothing Then
                    WS.Range(YearBuilt & I) = "No Year Built Information Available"
                Else
                    WS.Range(YearBuilt & I) = xmlYearBuilt.Text
                End If
                
                'get bedroom count
                Set xmlBedrooms = xmldoc.SelectSingleNode("//response/results/result/bedrooms")
                If xmlBedrooms Is Nothing Then
                    WS.Range(Bedrooms & I) = "No Bedroom Count Available"
                Else
                    WS.Range(Bedrooms & I) = xmlBedrooms.Text
                End If
                
                'get bathroom count
                Set xmlBathrooms = xmldoc.SelectSingleNode("//response/results/result/bathrooms")
                If xmlBathrooms Is Nothing Then
                    WS.Range(Bathrooms & I) = "No Bathroom Count Available"
                Else
                    WS.Range(Bathrooms & I) = xmlBathrooms.Text
                End If
                
                'get SQFT count
                Set xmlSQFT = xmldoc.SelectSingleNode("//response/results/result/finishedSqFt")
                If xmlSQFT Is Nothing Then
                    WS.Range(SQFT & I) = "No SQFT Available"
                Else
                    WS.Range(SQFT & I) = xmlSQFT.Text
                End If
                
                'get last sold date
                Set xmlLastSoldDate = xmldoc.SelectSingleNode("//response/results/result/lastSoldDate")
                If xmlLastSoldDate Is Nothing Then
                    WS.Range(LastSoldDate & I) = "No Last Sold Date Available"
                Else
                    WS.Range(LastSoldDate & I) = xmlLastSoldDate.Text
                End If
                
                'get last sold price
                Set xmllastSoldPrice = xmldoc.SelectSingleNode("//response/results/result/lastSoldPrice")
                If xmllastSoldPrice Is Nothing Then
                    WS.Range(lastSoldPrice & I) = "No Last Sold Price Available"
                Else
                    WS.Range(lastSoldPrice & I) = xmllastSoldPrice.Text
                End If
                
                'get zillow property id
                Set xmlzpropID = xmldoc.SelectSingleNode("//response/results/result/zpid")
                If xmlzpropID Is Nothing Then
                    WS.Range(ErrorMessage & I) = "No Zillow ID..."
                Else
                    WS.Range(zpropID & I) = xmlzpropID.Text
                End If
                
                'get zestimate
                Set xmlZAmount = xmldoc.SelectSingleNode("//response/results/result/zestimate/amount")
                If xmlZAmount Is Nothing Then
                    WS.Range(Zestimate & I) = "No Zestimate Available"
                Else
                    WS.Range(Zestimate & I) = xmlZAmount.Text
                    WS.Range(Zestimate & I).NumberFormat = "$#,##0_);($#,##0)"
                End If
                
                'get comparables
                Set xmlComparables = xmldoc.SelectSingleNode("//response/results/result/links/comparables")

                If xmlComparables Is Nothing Then
                    WS.Range(Comparables & I) = "No comparables available"
                Else
                    WS.Range(Comparables & I).Formula = "=HYPERLINK(""" & xmlComparables.Text & """,""Zillow Comparables"")"
                End If
                
                                                
                zipID = WS.Range(zpropID & I)
                URL = "http://www.zillow.com/webservice/GetUpdatedPropertyDetails.htm?zws-id=" & ZWSID & "&zpid=" & zipID
                
                'get listed status
                Set xmlListed = xmldoc.SelectSingleNode("//response/status")
                If xmlListed Is Nothing Then
                    WS.Range(Listed & I) = "Not listed "
                Else
                    WS.Range(Listed & I) = xmlListed.Text
                End If
                

            
                
                
            'if we are ending here is from check for an error message
            End If
            
       ' Document failed to load statement
       Else
       WS.Range(ErrorMessage & I) = "The document failed to load. Check your internet connection."
       
       End If
    
    ' Loop to top for next row
    Next I
    
' notify status bar of completion
Application.StatusBar = "Search complete!"

End Sub
