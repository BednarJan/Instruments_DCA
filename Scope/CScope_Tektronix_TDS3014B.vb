Imports Ivi.Visa

Public Class CScope_Tektronix_TDS3014B
    Inherits BCScopeTEK
    Implements IScope

    Sub New(Session As IMessageBasedSession, ErrorLogger As CErrorLogger)

        MyBase.New(Session, ErrorLogger, 4)

    End Sub

End Class
