Option Explicit
Const ExportFolde As String = "CDbl(txtExportFolder.Text)"

Private Sub UserForm_Activate()
    cobPlane.AddItem "Front"
    cobPlane.AddItem "Top"
    cobPlane.AddItem "Right"
End Sub

Private Sub cmdCreate_Click()
Dim swApp As SldWorks.SldWorks
Set swApp = Application.SldWorks
Dim Part As SldWorks.ModelDoc2
Set Part = swApp.ActiveDoc
Dim boolstatus As Boolean
Dim FR As Double, SR As Double, CH As Double
Dim CP As String

'Select Plane'
CP = cobPlane.Text
boolstatus = Part.Extension.SelectByID2(CP, "PLANE", 0, 0, 0, False, 0, Nothing, 0)

'Create tow circle in plane '
FR = CDbl(txtRadius1.Text) / 1000
SR = CDbl(txtRadius2.Text) / 1000
Dim skSegment As Object
Set skSegment = Part.SketchManager.CreateCircle(0#, 0#, 0#, FR, 0#, 0#)
Set skSegment = Part.SketchManager.CreateCircle(0#, 0#, 0#, SR, 0#, 0#)

'Create Cylender'
Dim myFeature As Object

CH = CDbl(txtHight.Text) / 1000
Set myFeature = Part.FeatureManager.FeatureExtrusion2(False, False, False, 0, 0, CH, CH, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)

End Sub

Private Sub cmdPart_Click()

'Connect to SOLIDWORKS
Dim swApp As SldWorks.SldWorks
Set swApp = Application.SldWorks

'Connect to Model
Dim swModel As SldWorks.ModelDoc2
Dim partTemplate As String

'Create new part from solidworks default template
partTemplate = swApp.GetUserPreferenceStringValue(swDefaultTemplatePart)
Set swModel = swApp.NewDocument(partTemplate, 0, 0#, 0#)

'Add Rocket custom information when Rocket Custom Information is True
If chkCustProp.Value = True Then
    Dim retval As Long
    Dim custPropMan As CustomPropertyManager

    Set custPropMan = swModel.Extension.CustomPropertyManager("")
    retval = custPropMan.Add3("Artikelnummer", swCustomInfoText, "", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Revision", swCustomInfoText, "000", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Benennung", swCustomInfoText, "$PRP:""SW-File Name""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Material", swCustomInfoText, """SW-Material@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Gewicht", swCustomInfoText, """SW-Mass@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Artikeltyp", swCustomInfoText, "Einzelteil", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Warengruppe", swCustomInfoText, "Lift", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Zeichnung", swCustomInfoYesOrNo, "Yes", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Hersteller-Artikelnr.", swCustomInfoText, "", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Bemerkung", swCustomInfoText, "", swCustomPropertyDeleteAndAdd)
End If

'Add Rocket System Option when Rocket System Option is True
If chkSystOpti.Value = True Then

    'Change backgraund color = Light
    Dim boolstatus As Boolean
    boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsBackground, swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light)
Else
    'Change backgraund color = Dark
    boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsBackground, swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark)
End If

'Add Rocket Document Properties when Rocket Document Properties is True
If chkDocuProp.Value = True Then
    
    'Change image quality to 3mm
    swModel.SetUserPreferenceDoubleValue swImageQualityShadedDeviation, 0.0005
End If
'frmSolidAssistant.Hide
End Sub

Private Sub cmdPartExis_Click()
'Connect to model
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActivateDoc("")

If swModel Is Nothing Then 'Check to see if a document is loaded
    swApp.SendMsgToUser2 "Please open a Part Document.", swMbStop, swMbOk
End If

'Add custom property to Existing Part

If chkCustProp.Value = True Then
    Dim retval As Long
    Dim custPropMan As CustomPropertyManager

    Set custPropMan = swModel.Extension.CustomPropertyManager("")
    retval = custPropMan.Add3("Artikelnummer", swCustomInfoText, "L1.20.202111281151", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Revision", swCustomInfoText, "000", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Benennung", swCustomInfoText, "$PRP:""SW-File Name""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Material", swCustomInfoText, """SW-Material@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Gewicht", swCustomInfoText, """SW-Mass@Part.SLDPRT""", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Artikeltyp", swCustomInfoText, "Einzelteil", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Warengruppe", swCustomInfoText, "Lift", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Zeichnung", swCustomInfoYesOrNo, "Yes", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Hersteller-Artikelnr.", swCustomInfoText, "", swCustomPropertyDeleteAndAdd)
    retval = custPropMan.Add3("Bemerkung", swCustomInfoText, "", swCustomPropertyDeleteAndAdd)
End If

'Add Rocket System Option when Rocket System Option is True
If chkSystOpti.Value = True Then

    'Change backgraund color = Light
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsBackground, swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light)
Else
    boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swSystemColorsBackground, swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Dark)
End If

'Add Rocket Document Properties when Rocket Document Properties is True
If chkDocuProp.Value = True Then
    
    'Change image quality to 3mm
    swModel.SetUserPreferenceDoubleValue swImageQualityShadedDeviation, 0.003
End If
End Sub

Private Sub cmdPDF_Click()
    Dim swApp As SldWorks.SldWorks
    Dim fileName As String
    Dim filePath As String
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    Dim longstatus As Long

    If Not swModel Is Nothing Or Len(txtExportFolder.Text) < 0 Then
        PrintGeneralProperties swModel
        fileName = PrintGeneralProperties(swModel)
        
        'Save As active part and show mssege box
        filePath = txtExportFolder.Text
        longstatus = swModel.SaveAs3(filePath + "\" + fileName + ".pdf", 0, 2)
        MsgBox "The drawing was saved in PDF format. File path: " + filePath + "\" + fileName + ".pdf"
    Else
        MsgBox "Please make sure that a drawing in SolidWorks is open and an export path is selected."
    End If
frmSolidAssistant.Hide
End Sub

Private Sub cmdSTEP_Click()
    Dim swApp As SldWorks.SldWorks
    Dim fileName As String
    Dim filePath As String
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    Dim longstatus As Long

    If Not swModel Is Nothing Or Len(txtExportFolder.Text) < 0 Then
        PrintGeneralProperties swModel
        fileName = PrintGeneralProperties(swModel)
        
        'Save As active part and show mssege box
        filePath = txtExportFolder.Text
        longstatus = swModel.SaveAs3(filePath + "\" + fileName + ".step", 0, 2)
        MsgBox "The 3D file was saved in STEP format. File path: " + filePath + "\" + fileName + ".step"
    Else
        MsgBox "Please make sure that a 3D file in SolidWorks is open and an export path is selected."
    End If
frmSolidAssistant.Hide
End Sub

Function PrintGeneralProperties(model As SldWorks.ModelDoc2) As String
Dim cusPro As String
   
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    Set swCustPrpMgr = model.Extension.CustomPropertyManager("")
    Debug.Print "General Properties"
    cusPro = PrintProperties(swCustPrpMgr, False, "")
    PrintGeneralProperties = cusPro
    
End Function

Function PrintProperties(custPrpMgr As SldWorks.CustomPropertyManager, cached As Boolean, indent As String) As String
    
    Dim vPrpNames As Variant
    vPrpNames = custPrpMgr.GetNames()
    Dim aNumber As String
    Dim name As String
    Dim fileName As String
    Dim i As Integer

    If Not IsEmpty(vPrpNames) Then
    
        For i = 0 To UBound(vPrpNames)
            Dim prpName As String
            prpName = vPrpNames(i)
            Dim prpVal As String
            Dim prpResVal As String
            Dim wasResolved As Boolean
            Dim isLinked As Boolean
            
            Dim res As Long
            res = custPrpMgr.Get6(prpName, cached, prpVal, prpResVal, wasResolved, isLinked)
            
            Dim status As String
            Select Case res
                Case swCustomInfoGetResult_e.swCustomInfoGetResult_CachedValue
                    status = "Cached Value"
                Case swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue
                    status = "Resolved Value"
                Case swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent
                    status = "Not Present"
            End Select
            
            Debug.Print indent & "Property: " & prpName
            Debug.Print indent & "Value/Text Expression: " & prpVal
            Debug.Print indent & "Evaluated Value: " & prpResVal
            Debug.Print indent & "Was Resolved: " & wasResolved
            Debug.Print indent & "Is Linked: " & isLinked
            Debug.Print indent & "Status: " & status
            Debug.Print ""
            
            If prpName = "Artikelnummer" Then
            aNumber = aNumber + prpResVal + "-"
            End If
            
            If prpName = "Benennung" Then
            name = name + prpResVal
            End If
        Next
        
    Else
        Debug.Print indent & "-No Properties-"
    End If

    fileName = aNumber + name
    PrintProperties = fileName
End Function