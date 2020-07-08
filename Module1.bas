Attribute VB_Name = "Module1"
Global con As ADODB.Connection
Global rs As ADODB.Recordset


Public Function connectdb()
Set con = New ADODB.Connection
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\bus.mdb;Persist Security Info=False")

End Function
