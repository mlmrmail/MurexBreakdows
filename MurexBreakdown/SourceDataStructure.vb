Imports MurexBreakdown

Public Class SourceDataStructure

    Private tDate As Date
    Private instrument As String
    Private portfolio As String
    Private currency As String
    Private flows As Double

    Public Property Flow As Double
        Get
            Return flows
        End Get
        Set(value As Double)
            flows = value
        End Set
    End Property

    Public Property TxDate As Date
        Get
            Return tDate
        End Get
        Set(value As Date)
            tDate = value
        End Set
    End Property

    Public Property InstrumentDesc As String
        Get
            Return instrument
        End Get
        Set(value As String)
            instrument = value
        End Set
    End Property

    Public Property Portf As String
        Get
            Return portfolio
        End Get
        Set(value As String)
            portfolio = value
        End Set
    End Property

    Public Property Curr As String
        Get
            Return currency
        End Get
        Set(value As String)
            currency = value
        End Set
    End Property


End Class
