Module modGlobal
    'contains all variables, constatns, proecdures and function
    'that need to be accessed in more than one form
#Region "Action Constants"
    Public Const ACTION_HOME As Integer = 0
    Public Const ACTION_MEMBER As Integer = 1
    Public Const ACTION_NONE As Integer = 2
    Public Const ACTION_ROLE As Integer = 3
    Public Const ACTION_EVENT As Integer = 4
    Public Const ACTION_RSVP As Integer = 5
    Public Const ACTION_COURSE As Integer = 6
    Public Const ACTION_SEMESTER As Integer = 7
    Public Const ACTION_HELP As Integer = 8
    Public Const ACTION_LOGOUT As Integer = 9




#End Region
    Public intNextAction As Integer
    Public myDB As New CDB
End Module
