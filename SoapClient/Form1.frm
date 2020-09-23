VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sClient As New MSSOAPLib.SoapClient
Private Const c_WSDL_URL As String = _
    "http://cst48/espresentation/webservices/systemstatus.asmx?WSDL"

'You need soap Type library 3.0 and microsoft xml v2.6to run this example
Private Sub Command1_Click()

'sClient.mssoapinit c_WSL_URL
Debug.Print Time

    TranslateBabel

Debug.Print Time
    
MsgBox "finish    "
End Sub

Public Sub TranslateBabel()

    ' Purpose: Translates text from one language to another.
    ' WSDL: http://services.xmltoday.com/vx_engine/wsdl_publish.vep/translate.wsdl
    ' More info: http://www.xmethods.net/detail.html?id=94
    
    Dim objClient As MSSOAPLib.SoapClient
    ' To package SOAP request.
    Dim objSerial As MSSOAPLib.SoapSerializer
    ' To read SOAP response.
    Dim objRead As MSSOAPLib.SoapReader
    ' To connect to Web service using SOAP.
    Dim objConn As MSSOAPLib.SoapConnector
    ' To parse the SOAP response.
    Dim objResults As MSXML2.IXMLDOMNodeList
    Dim objNode As MSXML2.IXMLDOMNode
    
    ' Set up the SOAP connector.
    Set objConn = New MSSOAPLib.HttpConnector
    ' Define the endpoint URL. This is the actual running code,
    ' not the WSDL file path! You can find it in the WSDL's
    ' <soap:address> tag's location attribute.
    objConn.Property("EndPointURL") = "http://cst48/espresentation/webservices/systemstatus.asmx"
    ' Define the SOAP action. You can find it in the WSDL's
    ' <soap:operation> tag's soapAction attribute for the matching
    ' <operation> tag.
    
    'GetNaturalInfo is the name of the service
    objConn.Property("SoapAction") = "http://tempuri.org/GetNaturalInfo"
    'objConn.Property("SoapAction") = "GetNaturalInfo"
    
    ' Begin the SOAP message.
    objConn.BeginMessage
    
    Set objSerial = New MSSOAPLib.SoapSerializer
    ' Initialize the serializer to the connector's input stream.
    objSerial.Init objConn.InputStream
    
    ' Build the SOAP message.
    With objSerial
        .startEnvelope              ' <SOAP-ENV:Envelope>
        .startBody                  ' <SOAP-ENV:Body>
        ' Use the Web method's name and schema target namespace URI.
        .startElement "GetNaturalInfo"
        .endElement
        .endBody                    ' </SOAP-ENV:Body>
        .endEnvelope                ' </SOAP-ENV:Envelope>
    End With
    
    ' Send the SOAP message.
    objConn.EndMessage
    
    Set objRead = New MSSOAPLib.SoapReader
    
    ' Initialize the SOAP reader to the connector's output stream.
    objRead.Load objConn.OutputStream
      
    Set objResults = objRead.RPCResult.childNodes
        
    ' Iterate through the returned nodes.
    For Each objNode In objResults
        'Debug.Print objNode.nodeValue
        MsgBox objNode.nodeTypedValue
     Next objNode
    
        
End Sub



