Option Explicit

'Origin:https://social.msdn.microsoft.com/Forums/office/en-US/62938e12-5cad-4de7-92e9-00314813d31a/publish-date-property-in-word?forum=worddev
Sub SetDeveloperTabActive()

    Dim RibbonTab As IAccessible
    Set RibbonTab = GetAccessible(CommandBars("Ribbon"), ROLE_SYSTEM_PAGETAB,"Developer")

    If Not RibbonTab Is Nothing Then
        If ((RibbonTab.accState(CHILDID_SELF) And (STATE_SYSTEM_UNAVAILABLE Or STATE_SYSTEM_INVISIBLE)) = 0) Then
            RibbonTab.accDoDefaultAction CHILDID_SELF
        Else
            MsgBox "Designated Tab is unavailable"
        End If
    End If

End Sub