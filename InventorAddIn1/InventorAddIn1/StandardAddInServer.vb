Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Win32

Namespace InventorAddIn1
    <ProgIdAttribute("InventorAddIn1.StandardAddInServer"),
    GuidAttribute("0f5faded-228e-4dd6-b520-32786da4dd3f")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents m_uiEvents As UserInterfaceEvents

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' Release objects.
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        Private Sub AddToUserInterface()
            ' This is where you'll add code to add buttons to the ribbon.
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub
#End Region

    End Class

    Public Module Globals
        Public g_inventorApplication As Inventor.Application

        Public Function AddInClientID() As String
            Dim guid As String = ""
            Try
                Dim t As Type = GetType(InventorAddIn1.StandardAddInServer)
                Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
                Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
                guid = "{" + guidAttribute.Value.ToString() + "}"
            Catch
            End Try

            Return guid
        End Function
    End Module
End Namespace
