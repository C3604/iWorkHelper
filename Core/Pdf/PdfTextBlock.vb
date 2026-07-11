Imports System.Collections.Generic

''' <summary>
''' 由若干相邻逻辑行组成的文本区块（如发票的「购销信息区」「商品明细区」「合计区」）。
''' 用于常规发票的分区解析，避免全文一维正则跨区误取。
''' </summary>
Public Class PdfTextBlock
    Public Sub New()
        Lines = New List(Of PdfTextLine)()
    End Sub

    ''' <summary>区块角色（header / parties / lineItems / totals / other）。</summary>
    Public Property Role As String
    ''' <summary>区块内的逻辑行（按从上到下顺序）。</summary>
    Public Property Lines As List(Of PdfTextLine)
    ''' <summary>起始行在整份文档逻辑行列表中的下标。</summary>
    Public Property StartIndex As Integer
    ''' <summary>结束行下标（不含）。</summary>
    Public Property EndIndex As Integer

    Public ReadOnly Property IsEmpty As Boolean
        Get
            Return Lines Is Nothing OrElse Lines.Count = 0
        End Get
    End Property
End Class
