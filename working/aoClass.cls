VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocTypeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Csv As Object
Dim Main As Object
'
'========================================================================
'   v3.3 Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
'
'
Public Sub Init(MainObj As Object)
    Set Main = MainObj
End Sub
'
'
'
Public Function GetContent(OptionString As String) As String
    Dim Copy As String
    Dim Help As String
    Dim FieldValue As String
    Dim FieldName As String
    '
    ' DocType declaration
    '
    FieldName = Main.GetAddonOption("FieldName", OptionString)
    FieldValue = Main.GetAddonOption("FieldValue", OptionString)
    
    Copy = Main.GetFormInputTextExpandable(FieldName, FieldValue, 2, , "DTD")
    Help = "" _
        & "<div>This is the default DocType Declaration used on all pages of the site. Manually enter your DocType, or choose one from the list below.</div>" _
        & "<div style=""height:20px;"">&nbsp;</div>" _
        & "<div style=""clear:both;margin-left:40px;margin-top:5px;white-space:nowrap;""><input type=""radio"" name=""dmy"" value=""0"" onClick=""document.getElementById('DTD').value='&lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD HTML 4.01//EN&quot; &quot;http://www.w3.org/TR/html4/strict.dtd&quot;&gt;'""/>HTML 4.01 Strict</div>" _
        & "<div style=""clear:both;margin-left:40px;margin-top:5px;white-space:nowrap;""><input type=""radio"" name=""dmy"" value=""1"" onClick=""document.getElementById('DTD').value='&lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD HTML 4.01 Transitional//EN&quot; &quot;http://www.w3.org/TR/html4/loose.dtd&quot;&gt;'""/>HTML 4.01 Transitional</div>" _
        & "<div style=""clear:both;margin-left:40px;margin-top:5px;white-space:nowrap;""><input type=""radio"" name=""dmy"" value=""2"" onClick=""document.getElementById('DTD').value='&lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD XHTML 1.0 Strict//EN&quot; &quot;http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd&quot;&gt;'""/>XHTML 1.0 Strict</div>" _
        & "<div style=""clear:both;margin-left:40px;margin-top:5px;white-space:nowrap;""><input type=""radio"" name=""dmy"" value=""3"" onClick=""document.getElementById('DTD').value='&lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD XHTML 1.0 Transitional//EN&quot; &quot;http://www.w3.org/TR/xhtml1/DTD/xhtml1-Transitional.dtd&quot;&gt;'""/>XHTML 1.0 Transitional</div>" _
        & ""
    GetContent = Copy & "<div style=""padding:10px;"">" & Help & "</div>"
End Function

