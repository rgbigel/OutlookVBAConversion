VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOlExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' declare an object variable for the Explorers object IN CLASS MODULE cOlExEvent
' You must use the WithEvents keyword to specify that the object variable
' will be used to respond to events triggered by the object.
Public WithEvents OlExplorers As Outlook.Explorers
Attribute OlExplorers.VB_VarHelpID = -1
Public WithEvents OlExplorer As Explorer           ' for Explorer(s) events
Attribute OlExplorer.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set OlExplorer = olApp.ActiveExplorer
    Set OlExplorers = olApp.Explorers
    Debug.Print "cOlExplorer class Inited with Events for ActiveExplorer and Explorers"
End Sub

'---------------------------------------------------------------------------------------
' Method : Sub RegisterClose
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub RegisterClose(Cancel As Boolean)
Dim zErr As cErr
Const zKey As String = "EventRoutines.OlExplorer_Close"
   
    Debug.Print "cOlExplorer RegisterClose called"
    If Explorers.Count <= 1 Then
        DoVerify False
    End If
    
ProcReturn:

End Sub

' Event Procedure
' After OlExplorers object has been declared WithEvents, it appears in the Object list in the class module Code window,
' and you can select the object's event procedures from the Procedures/Events list.
' For example, when you select the Close event for an Explorers object declared as OlExplorers,
' the following empty procedure appears in the Code window.

'---------------------------------------------------------------------------------------
' Method : Sub OlExplorers_Close
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub OlExplorers_Close(Item As Object, Cancel As Boolean) ' Event Procedure
    Debug.Print "cOlExplorer OlExplorers_Close called"
End Sub

' Initialize the Declared Object
' connect the declared object OlExplorers with the Explorers object.
' If you declared the object in a class module named EventClassModule, then you can use the code in any module.

' Register_Event_Handler
' When this procedure is run, the OlExplorers object in the form or class module points
' to the Outlook Explorers object, and the event procedure will run when the event occurs.
Sub Register_Event_Handler()
    Debug.Print "cOlExplorer Register_Event_Handler called"
End Sub

'---------------------------------------------------------------------------------------
' Method : Sub OlExplorers_NewExplorer
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub OlExplorers_NewExplorer(ByVal Explorer As Explorer)
    Debug.Print "cOlExplorer NewExplorer called for new Explorer Window: " & Explorer.Caption
End Sub

