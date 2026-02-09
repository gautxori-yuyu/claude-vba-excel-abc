Class cMsgIEReporter
	' si quisiera volcar salida a consola / Echo, crearia OTRO REPORTER, que use la consola como salida...
    Private MsgIE
    
    Public Function Init(MsgIE)
        Set Me.MsgIE = MsgIE
        Set Init = Me
    End Function
    
    Sub StartSection(id, title, style, bNoWrap)
        If MsgIE Is Nothing Then Exit Sub
        Call MsgIE.Spoiler(True, style, title, id, bNoWrap)
    End Sub
    
    Sub ReportData(sectionId, dataMsg)
        If MsgIE Is Nothing Then Exit Sub
        Call MsgIE.setContainer(sectionId)
        MsgIE (dataMsg)
        MsgIE.popContainer
    End Sub
    
    Sub EndSection(id)
        If MsgIE Is Nothing Then Exit Sub
        Call MsgIE.setContainer(id)
        MsgIE.popContainer
    End Sub
    
    Sub LogError(msg)
        If MsgIE Is Nothing Then Exit Sub
        MsgIE.MsgLog msg
    End Sub
End Class