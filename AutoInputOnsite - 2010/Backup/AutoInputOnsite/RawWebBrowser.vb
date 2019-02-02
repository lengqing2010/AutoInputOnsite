Public Class RawWebBrowser
    Inherits System.Windows.Forms.AxHost
    Sub New()
        MyBase.New("8856f961-340a-11d0-a96b-00c04fd705a2")
    End Sub

    Event Initialized(ByVal sender As Object, ByVal e As EventArgs)

    Protected Overrides Sub AttachInterfaces()
        RaiseEvent Initialized(Me, New EventArgs())
    End Sub

    Public ReadOnly Property ActiveXInstance() As Object
        Get
            Return MyBase.GetOcx()
        End Get
    End Property
End Class
