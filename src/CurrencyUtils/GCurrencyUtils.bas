Attribute VB_Name = "GCurrencyUtils"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GCurrencyUtils"

'@================================================================================
' Member variables
'@================================================================================

Private mCurrencyDescsColl                              As SortedDictionary
Private mCurrencyDescriptors()                          As CurrencyDescriptor

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function CreateCurrencyConverter( _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pContractStore As IContractStore) As ICurrencyConverter
Const ProcName As String = "CreateCurrencyConverter"
On Error GoTo Err

Dim lCurrencyConverter As New CurrencyConverter
lCurrencyConverter.Initialise pMarketDataManager, pContractStore
Set CreateCurrencyConverter = lCurrencyConverter

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptor( _
                ByVal CurrencyCode As String) As CurrencyDescriptor
Const ProcName As String = "GetCurrencyDescriptor"
On Error GoTo Err

GetCurrencyDescriptor = getCurrencyDescr(CurrencyCode)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "GetCurrencyDescriptors"
On Error GoTo Err

GetCurrencyDescriptors = getCurrencyDescrs

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidCurrencyCode(ByVal CurrencyCode As String) As Boolean
Const ProcName As String = "IsValidCurrencyCode"
On Error GoTo Err

IsValidCurrencyCode = isValidCurrCode(CurrencyCode)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addCurrencyDesc( _
                ByVal Code As String, _
                ByVal Description As String)
Const ProcName As String = "addCurrencyDesc"
On Error GoTo Err

Dim lDescriptor As CurrencyDescriptor
Code = UCase$(Code)
lDescriptor.Code = Code
lDescriptor.Description = Description
mCurrencyDescsColl.Add lDescriptor, Code

Exit Sub

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCurrencyDescr( _
                ByVal Code As String) As CurrencyDescriptor
Const ProcName As String = "getCurrencyDescr"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then setupCurrencyDescs

Code = UCase$(Code)
AssertArgument mCurrencyDescsColl.Contains(Code), "Invalid currency Code"

getCurrencyDescr = mCurrencyDescsColl.Item(Code)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCurrencyDescrs() As CurrencyDescriptor()
Const ProcName As String = "getCurrencyDescrs"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then
    setupCurrencyDescs
    ReDim mCurrencyDescriptors(mCurrencyDescsColl.Count - 1) As CurrencyDescriptor
    Dim lDesc As Variant
    Dim i As Long
    For Each lDesc In mCurrencyDescsColl
        mCurrencyDescriptors(i) = lDesc
        i = i + 1
    Next
End If
getCurrencyDescrs = mCurrencyDescriptors

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidCurrCode(ByVal Code As String) As Boolean
Const ProcName As String = "isValidCurrCode"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then setupCurrencyDescs

isValidCurrCode = mCurrencyDescsColl.Contains(UCase$(Code))

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupCurrencyDescs()
Const ProcName As String = "setupCurrencyDescs"
On Error GoTo Err

Set mCurrencyDescsColl = CreateSortedDictionary(KeyTypeString)

addCurrencyDesc "AED", "United Arab Emirates, Dirhams"
addCurrencyDesc "AFN", "Afghanistan, Afghanis"
addCurrencyDesc "ALL", "Albania, Leke"
addCurrencyDesc "AMD", "Armenia, Drams"
addCurrencyDesc "ANG", "Netherlands Antilles, Guilders"
addCurrencyDesc "AOA", "Angola, Kwanza"
addCurrencyDesc "ARS", "Argentina, Pesos"
addCurrencyDesc "AUD", "Australia, Dollars"
addCurrencyDesc "AWG", "Aruba, Guilders"
addCurrencyDesc "AZN", "Azerbaijan, New Manats"
addCurrencyDesc "BAM", "Bosnia and Herzegovina, Convertible Marka"
addCurrencyDesc "BBD", "Barbados, Dollars"
addCurrencyDesc "BDT", "Bangladesh, Taka"
addCurrencyDesc "BGN", "Bulgaria, Leva"
addCurrencyDesc "BHD", "Bahrain, Dinars"
addCurrencyDesc "BIF", "Burundi, Francs"
addCurrencyDesc "BMD", "Bermuda, Dollars"
addCurrencyDesc "BND", "Brunei Darussalam, Dollars"
addCurrencyDesc "BOB", "Bolivia, Bolivianos"
addCurrencyDesc "BRL", "Brazil, Brazil Real"
addCurrencyDesc "BSD", "Bahamas, Dollars"
addCurrencyDesc "BTN", "Bhutan, Ngultrum"
addCurrencyDesc "BWP", "Botswana, Pulas"
addCurrencyDesc "BYR", "Belarus, Rubles"
addCurrencyDesc "BZD", "Belize, Dollars"
addCurrencyDesc "CAD", "Canada, Dollars"
addCurrencyDesc "CDF", "Congo/Kinshasa, Congolese Francs"
addCurrencyDesc "CHF", "Switzerland, Francs"
addCurrencyDesc "CLP", "Chile, Pesos"
addCurrencyDesc "CNY", "China, Yuan Renminbi"
addCurrencyDesc "COP", "Colombia, Pesos"
addCurrencyDesc "CRC", "Costa Rica, Colones"
addCurrencyDesc "CUP", "Cuba, Pesos"
addCurrencyDesc "CVE", "Cape Verde, Escudos"
addCurrencyDesc "CYP", "Cyprus, Pounds"
addCurrencyDesc "CZK", "Czech Republic, Koruny"
addCurrencyDesc "DJF", "Djibouti, Francs"
addCurrencyDesc "DKK", "Denmark, Kroner"
addCurrencyDesc "DOP", "Dominican Republic, Pesos"
addCurrencyDesc "DZD", "Algeria, Algeria Dinars"
addCurrencyDesc "EEK", "Estonia, Krooni"
addCurrencyDesc "EGP", "Egypt, Pounds"
addCurrencyDesc "ERN", "Eritrea, Nakfa"
addCurrencyDesc "ETB", "Ethiopia, Birr"
addCurrencyDesc "EUR", "Euro Member Countries, Euro"
addCurrencyDesc "FJD", "Fiji, Dollars"
addCurrencyDesc "FKP", "Falkland Islands (Malvinas), Pounds"
addCurrencyDesc "GBP", "United Kingdom, Pounds"
addCurrencyDesc "GEL", "Georgia, Lari"
addCurrencyDesc "GGP", "Guernsey, Pounds"
addCurrencyDesc "GHC", "Ghana, Cedis"
addCurrencyDesc "GHS", "Ghana, Cedis"
addCurrencyDesc "GIP", "Gibraltar, Pounds"
addCurrencyDesc "GMD", "Gambia, Dalasi"
addCurrencyDesc "GNF", "Guinea, Francs"
addCurrencyDesc "GTQ", "Guatemala, Quetzales"
addCurrencyDesc "GYD", "Guyana, Dollars"
addCurrencyDesc "HKD", "Hong Kong, Dollars"
addCurrencyDesc "HNL", "Honduras, Lempiras"
addCurrencyDesc "HRK", "Croatia, Kuna"
addCurrencyDesc "HTG", "Haiti, Gourdes"
addCurrencyDesc "HUF", "Hungary, Forint"
addCurrencyDesc "IDR", "Indonesia, Rupiahs"
addCurrencyDesc "ILS", "Israel, New Shekels"
addCurrencyDesc "IMP", "Isle of Man, Pounds"
addCurrencyDesc "INR", "India, Rupees"
addCurrencyDesc "IQD", "Iraq, Dinars"
addCurrencyDesc "IRR", "Iran, Rials"
addCurrencyDesc "ISK", "Iceland, Kronur"
addCurrencyDesc "JEP", "Jersey, Pounds"
addCurrencyDesc "JMD", "Jamaica, Dollars"
addCurrencyDesc "JOD", "Jordan, Dinars"
addCurrencyDesc "JPY", "Japan, Yen"
addCurrencyDesc "KES", "Kenya, Shillings"
addCurrencyDesc "KGS", "Kyrgyzstan, Soms"
addCurrencyDesc "KHR", "Cambodia, Riels"
addCurrencyDesc "KMF", "Comoros, Francs"
addCurrencyDesc "KPW", "Korea (North), Won"
addCurrencyDesc "KRW", "Korea (South), Won"
addCurrencyDesc "KWD", "Kuwait, Dinars"
addCurrencyDesc "KYD", "Cayman Islands, Dollars"
addCurrencyDesc "KZT", "Kazakhstan, Tenge"
addCurrencyDesc "LAK", "Laos, Kips"
addCurrencyDesc "LBP", "Lebanon, Pounds"
addCurrencyDesc "LKR", "Sri Lanka, Rupees"
addCurrencyDesc "LRD", "Liberia, Dollars"
addCurrencyDesc "LSL", "Lesotho, Maloti"
addCurrencyDesc "LTL", "Lithuania, Litai"
addCurrencyDesc "LVL", "Latvia, Lati"
addCurrencyDesc "LYD", "Libya, Dinars"
addCurrencyDesc "MAD", "Morocco, Dirhams"
addCurrencyDesc "MDL", "Moldova, Lei"
addCurrencyDesc "MGA", "Madagascar, Ariary"
addCurrencyDesc "MKD", "Macedonia, Denars"
addCurrencyDesc "MMK", "Myanmar (Burma), Kyats"
addCurrencyDesc "MNT", "Mongolia, Tugriks"
addCurrencyDesc "MOP", "Macau, Patacas"
addCurrencyDesc "MRO", "Mauritania, Ouguiyas"
addCurrencyDesc "MTL", "Malta, Liri"
addCurrencyDesc "MUR", "Mauritius, Rupees"
addCurrencyDesc "MVR", "Maldives (Maldive Islands), Rufiyaa"
addCurrencyDesc "MWK", "Malawi, Kwachas"
addCurrencyDesc "MXN", "Mexico, Pesos"
addCurrencyDesc "MYR", "Malaysia, Ringgits"
addCurrencyDesc "MZN", "Mozambique, Meticais"
addCurrencyDesc "NAD", "Namibia, Dollars"
addCurrencyDesc "NGN", "Nigeria, Nairas"
addCurrencyDesc "NIO", "Nicaragua, Cordobas"
addCurrencyDesc "NOK", "Norway, Krone"
addCurrencyDesc "NPR", "Nepal, Nepal Rupees"
addCurrencyDesc "NZD", "New Zealand, Dollars"
addCurrencyDesc "OMR", "Oman, Rials"
addCurrencyDesc "PAB", "Panama, Balboa"
addCurrencyDesc "PEN", "Peru, Nuevos Soles"
addCurrencyDesc "PGK", "Papua New Guinea, Kina"
addCurrencyDesc "PHP", "Philippines, Pesos"
addCurrencyDesc "PKR", "Pakistan, Rupees"
addCurrencyDesc "PLN", "Poland, Zlotych"
addCurrencyDesc "PYG", "Paraguay, Guarani"
addCurrencyDesc "QAR", "Qatar, Rials"
addCurrencyDesc "RON", "Romania, New Lei"
addCurrencyDesc "RSD", "Serbia, Dinars"
addCurrencyDesc "RUB", "Russia, Rubles"
addCurrencyDesc "RWF", "Rwanda, Rwanda Francs"
addCurrencyDesc "SAR", "Saudi Arabia, Riyals"
addCurrencyDesc "SBD", "Solomon Islands, Dollars"
addCurrencyDesc "SCR", "Seychelles, Rupees"
addCurrencyDesc "SDG", "Sudan, Pounds"
addCurrencyDesc "SEK", "Sweden, Kronor"
addCurrencyDesc "SGD", "Singapore, Dollars"
addCurrencyDesc "SHP", "Saint Helena, Pounds"
addCurrencyDesc "SKK", "Slovakia, Koruny"
addCurrencyDesc "SLL", "Sierra Leone, Leones"
addCurrencyDesc "SOS", "Somalia, Shillings"
addCurrencyDesc "SPL", "Seborga, Luigini"
addCurrencyDesc "SRD", "Suriname, Dollars"
addCurrencyDesc "STD", "São Tome and Principe, Dobras"
addCurrencyDesc "SVC", "El Salvador, Colones"
addCurrencyDesc "SYP", "Syria, Pounds"
addCurrencyDesc "SZL", "Swaziland, Emalangeni"
addCurrencyDesc "THB", "Thailand, Baht"
addCurrencyDesc "TJS", "Tajikistan, Somoni"
addCurrencyDesc "TMM", "Turkmenistan, Manats"
addCurrencyDesc "TND", "Tunisia, Dinars"
addCurrencyDesc "TOP", "Tonga, Pa'anga"
addCurrencyDesc "TRY", "Turkey, New Lira"
addCurrencyDesc "TTD", "Trinidad and Tobago, Dollars"
addCurrencyDesc "TVD", "Tuvalu, Tuvalu Dollars"
addCurrencyDesc "TWD", "Taiwan, New Dollars"
addCurrencyDesc "TZS", "Tanzania, Shillings"
addCurrencyDesc "UAH", "Ukraine, Hryvnia"
addCurrencyDesc "UGX", "Uganda, Shillings"
addCurrencyDesc "USD", "United States of America, Dollars"
addCurrencyDesc "UYU", "Uruguay, Pesos"
addCurrencyDesc "UZS", "Uzbekistan, Sums"
addCurrencyDesc "VEB", "Venezuela, Bolivares"
addCurrencyDesc "VND", "Viet Nam, Dong"
addCurrencyDesc "VUV", "Vanuatu, Vatu"
addCurrencyDesc "WST", "Samoa, Tala"
addCurrencyDesc "XAF", "Communauté Financière Africaine BEAC, Francs"
addCurrencyDesc "XAG", "Silver, Ounces"
addCurrencyDesc "XAU", "Gold, Ounces"
addCurrencyDesc "XCD", "East Caribbean Dollars"
addCurrencyDesc "XOF", "Communauté Financière Africaine BCEAO, Francs"
addCurrencyDesc "XPD", "Palladium Ounces"
addCurrencyDesc "XPF", "Comptoirs Français du Pacifique Francs"
addCurrencyDesc "XPT", "Platinum, Ounces"
addCurrencyDesc "YER", "Yemen, Rials"
addCurrencyDesc "ZAR", "South Africa, Rand"
addCurrencyDesc "ZMK", "Zambia, Kwacha"
addCurrencyDesc "ZWD", "Zimbabwe, Zimbabwe Dollars"

Exit Sub

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Sub







