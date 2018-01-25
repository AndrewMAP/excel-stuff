Attribute VB_Name = "Config"
Public Const stDB As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Backoffice\Funds Control\Base de Dados\BancoSTK01.accdb"
Public Const stDB_movimentacoes As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Sistema de Passivo\Movimentacoes_STK.accdb"
Global cnt As New ADODB.Connection
Global rs As New ADODB.Recordset
