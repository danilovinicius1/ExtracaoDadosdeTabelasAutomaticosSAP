set sh = CreateObject("WScript.Shell")
sh.run "name_server", 0
WScript.Sleep(5000)

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

dim dia 
dim tamanhodia
dia = day(now)
dim data

dim mes
dim tamanhomes
mes = month(now)

dim ano
ano = year(now)

tamanhodia = Len(dia)
tamanhomes = Len(mes)

if tamanhodia = 1 then
dia = "0" & day(now) 
else
dia = day(now)
end if

if tamanhomes = 1 then
mes = "0" & month(now)
else
mes = month(now)
end if

data = dia & mes & ano

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "logindosap"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "senhadosap"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[0]/okcd").text = "zcgat_tab"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "tabela"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[2]/menu[5]/menu[0]").select
session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").text = "selecionar a variante predefinida"
session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").setFocus
session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").caretPosition = 8
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,21]").text = data
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,21]").setFocus
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,21]").caretPosition = 8
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "selecionar o diretório"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = data & ".txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findByid("wnd[1]/tbar[0]/btn[11]").press
'session.findById("wnd[1]").sendVKey 0
Session.findbyid("wnd[0]").Close
Session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
