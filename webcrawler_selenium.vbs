Private Declare PtrSafe Function FindWindowExA Lib "user32.dll" ( _
  ByVal hwndParent As LongPtr, _
  ByVal hwndChildAfter As LongPtr, _
  ByVal lpszClass As String, _
  ByVal lpszWindow As String) As Long

Private Declare PtrSafe Function PostMessageA Lib "user32.dll" ( _
  ByVal hwnd As LongPtr, _
  ByVal wMsg As LongPtr, _
  ByVal wParam As LongPtr, _
  ByVal lParama As LongPtr) As Long

Private Declare PtrSafe Function GetWindowLongA Lib "user32.dll" ( _
  ByVal hwnd As LongPtr, ByVal nIndex As Integer) As Long

Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc.dll" ( _
  ByVal hwnd As LongPtr, _
  ByVal dwId As Long, _
  ByRef riid As Any, _
  ByRef ppvObject As IAccessible) As Long
                                       
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc.dll" ( _
  ByVal paccContainer As IAccessible, _
  ByVal iChildStart As Long, _
  ByVal cChildren As Long, _
  ByRef rgvarChildren As Variant, _
  ByRef pcObtained As Long) As Long
  
Public Sub Download_Pos_Carteira()
    Dim driver As New IEDriver
    Dim iDay, iMonth, iYear As Integer
    iDay = Day(Application.WorksheetFunction.WorkDay(Date, -1))
    iMonth = Month(Application.WorksheetFunction.WorkDay(Date, -1))
    iYear = Year(Application.WorksheetFunction.WorkDay(Date, -1))
    Dim DirFile As String
    
    If iDay < 10 Then
        iDay = "0" & iDay
    End If
    
    If iMonth < 10 Then
        iMonth = "0" & iMonth
    End If
    
    Dim query, strSenha As String
    
    cnt.Open stDB
    
    query = " SELECT Senha_SMA_BNYMellon.SenhaBNY FROM Senha_SMA_BNYMellon;"
    
    rs.Open query, cnt, adUseClient
    
    strSenha = rs.Fields(0)
    
    cnt.Close
    
    
    driver.Get "https://gestores.bnymellon.com.br/Extranet/CarteiraWeb/Paginas/ComposicaoCarteira.aspx"
    
    ' Para selecionar um certificado
    driver.SwitchToAlert.Accept
    
    'Login
    driver.FindElementById("ctl00_ContentPlaceHolder1_Login1_txtLogin").SendKeys "stkcapital"
    driver.FindElementById("ctl00_ContentPlaceHolder1_Login1_txtSenha").SendKeys strSenha
    driver.FindElementById("ctl00_ContentPlaceHolder1_Login1_imgBtContinuar").Click
    
    driver.Wait 1000
    
    driver.FindElementById("ctl00_ContentPlaceHolder1_PesquisaCarteira1_ddlOpcao").AsSelect.SelectByText ("Todas")
    
    driver.Wait 5000
    'Seleciona todos para baixo
    driver.FindElementById("ctl00_ContentPlaceHolder1_PesquisaCarteira1_imgbInclui").Click
    driver.Wait 1000
    'Seleciona a data
    driver.FindElementById("ctl00_ContentPlaceHolder1_dataPosicao_btnData").Click
    driver.FindElementById("ctl00_ContentPlaceHolder1_dataPosicao_txtData").SendKeys "value", "" & iMonth & "/" & iDay & "/" & iYear & ""
    
    driver.Wait 5000
    'Clica em Relatório
    driver.FindElementById("ctl00_ContentPlaceHolder1_btnRelatorio").Click
    driver.ExecuteScript "ValidarPostar"
    
    driver.Wait 50000
    
    'Clica em Excel
    driver.FindElementById("ctl00_ContentPlaceHolder1_ExportaDadosGridViewParaExcel1_imgExcel_2").Click
    driver.ExecuteScript "IniciaEspera"
    
    driver.Wait 30000
    
    'Clica em download para todos
    driver.FindElementById("ctl00_ContentPlaceHolder1_uscGerenciaDownload1_linkDownload").ExecuteScript "this.click()"
    filePath = DownloadFileSyncIE("H:\Backoffice\BNYDownloads\Carteira_Excel")
    
    driver.Wait 15000
    
    'driver.SwitchToAlert.Accept
    
    driver.Quit
    'quit method for closing browser instance.
    
    Call Unzip1(iDay, iMonth, iYear, "H:\Backoffice\BNYDownloads\Carteira_Excel", "CA_EXCEL_", "H:\Backoffice\Funds Control\Rotinas\Gerencial\Carteiras Mellon\Carteiras nao liberadas\carteiras_excel")
    
End Sub
