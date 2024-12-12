JSON Parser module implementation for Visual Basic 6.0 projects.

# Pre-requisites
- Find and check the "Microsoft Scripting Runtime" reference under `Project > References...`.

# Usage example
Given a sample JSON like this:
```json
{
  "id": 123,
  "address": {
    "name": "fake st",
    "number": 200,
    "receiver": {
      "firstname": "John",
      "lastname": "Doe"
    }
  },
  "products": [
    {
      "sku": "ABC-123",
      "quantity": 5
    }
  ]
}
```

You can work it like this:
```vb
Dim jsonSring As String
Dim parsed As Object

' The input JSON string
jsonString = "{""id"": 123,""address"": {""name"": ""fake st"",""number"": 200,""receiver"": {""firstname"": ""John"",""lastname"": ""Doe""}},""products"": [{""sku"": ""ABC-123"",""quantity"": 5}]}"

' The parsed JSON Object
Set parsed = ParseJSON(jsonString)

' Reading properties
Debug.print (parsed("id"))

' Reading an object property
Debug.print (parsed("address")("name"))

' Iterating through arrays
Dim productsArray As Collection
Set productsArray = jsonParser("products")

For Each product In productsArray
  Debug.print (product("sku"))
  Debug.print (product("quantity"))
Next product
```