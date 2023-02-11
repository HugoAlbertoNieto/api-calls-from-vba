Attribute VB_Name = "mdlMain"
Sub AddDraftOrder()
    ' This sub adds a new draft order
    ' API key and password must be obtained first on the .myshopify.com page.
    ' Click on Apps on the left menu, manage private apps,
    ' And click on Create a new private app
    ' Make sure the app has the proper privileges for read / write as needed.
    
    Dim strShopifyShop As String
    Dim strTargetURL As String
    Dim strJson As String
    Dim strAPIKey As String
    Dim strPassword As String
    Dim strProductNumber As String
    Dim sData As String
    Dim strResponse As String
    
    strShopifyShop = "yourShopifyShop"
    strAPIKey = "yourAPIKey"
    strPassword = "yourPassword"

    ' Test order
    strVariantNumber = "29015763124329"
    strJson = "" _
        & "{ " _
        & " ""draft_order"": { " _
        & "     ""line_items"": [ " _
        & "         { " _
        & "             ""variant_id"":" _
        & strVariantNumber & ", " _
        & "             ""quantity"": 1 " _
        & "         } " _
        & "     ] " _
        & " } " _
        & "} "
        strJson = Replace(strJson, " ", "")
    
        '*** Result of above strJson ***
        '*** strJson = {"draft_order":{"line_items":[{"variant_id":29220310253670,"quantity":1}]}} ***
        
    Dim objHttp As Object
    Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    sTargetURL = "https://" & strAPIKey & ":" & strPassword & "@" & strShopifyShop & ".myshopify.com/admin/api/2019-07/draft_orders.json"

    objHttp.Open "POST", sTargetURL, False
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.setRequestHeader "X-Shopify-Access-Token", strPassword
    objHttp.Option(6) = False
    objHttp.Option(12) = False
    objHttp.SetCredentials strAPIKey, strPassword, 0
    objHttp.Send (strJson)
    strResponse = objHttp.responseText
End Sub


Sub AddProduct()
    ' This function adds a new shopify product
    ' Other fields are available by adding them on the objHttp.Open line
    ' API key and password must be obtained first on the .myshopify.com page.
    ' Click on Apps on the left menu, manage private apps,
    ' And click on Create a new private app
    ' Make sure the app has the proper privileges for read / write as needed.
    
    Dim strShopifyShop As String
    Dim strTargetURL As String
    Dim strAPIKey As String
    Dim strPassword As String
    Dim strProductNumber As String
    Dim strResponse As String
    Dim sData As String
    
    strShopifyShop = "yourShopifyShop"
    strAPIKey = "yourAPIKey"
    strPassword = "yourPassword"
    strProductNumber = "75007614"
        
    Dim objHttp As Object
    Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    sTargetURL = "https://" & strAPIKey & ":" & strPassword & "@" & strShopifyShop & ".myshopify.com/admin/api/2019-07/products.json"

    sData = "{" _
        & Chr(34) & "product" & Chr(34) & ": {" _
        & Chr(34) & "title" & Chr(34) & ": " & Chr(34) & "Burton Custom Freestyle 151" & Chr(34) _
        & "}" _
        & "}"
    
    'sData = "{" _
        & " ""product"": {" _
        & " ""title"": " _
        & " ""Burton Custom Freestyle 151"", " _
        & "}" _
        & "}"

    Cells(2, 1) = sData
    objHttp.Open "POST", sTargetURL, False
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.setRequestHeader "X-Shopify-Access-Token", strPassword
    objHttp.Option(6) = False
    objHttp.Option(12) = False
    objHttp.SetCredentials strAPIKey, strPassword, 0
    objHttp.Send (sData)
    strResponse = objHttp.responseText
End Sub

Function getShopifyProducts()
    ' This function gets the shopify product titles.
    ' Other fields are available by adding them on the objHttp.Open line
    ' API key and password must be obtained first on the .myshopify.com page.
    ' Click on Apps on the left menu, manage private apps,
    ' And click on Create a new private app
    ' Make sure the app has the proper privileges for read / write as needed.
    ' The function will return a string in JSON format with all the requested fields.
    
    Dim strShopifyShop As String
    Dim strAPIKey As String
    Dim strPassword As String
    Dim strResponse As String
    
    strShopifyShop = "yourShopifyShop"
    strAPIKey = "yourAPIKey"
    strPassword = "yourPassword"
       
    Dim objHttp As Object
    Set objHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    objHttp.Open "GET", "https://" & strShopifyShop & ".myshopify.com/admin/api/2019-07/products.json?fields=title"
    strJson = ""
    '*********************************
                         
    objHttp.setRequestHeader "Content-Type", "application/json"
    objHttp.SetCredentials strAPIKey, strPassword, 0
    objHttp.Send (strJson)
    getShopifyProducts = objHttp.responseText
    ' See the result here:
    strResponse = CStr(getShopifyProducts)
End Function

