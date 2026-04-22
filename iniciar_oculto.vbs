Dim oShell, sDir
Set oShell = CreateObject("WScript.Shell")

' Definir diretorio de trabalho como a pasta onde este arquivo esta
sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
oShell.CurrentDirectory = sDir

' Iniciar o Streamlit sem janela (segundo parametro 0 = oculto)
oShell.Run "python -m streamlit run app.py --browser.gatherUsageStats false", 0, False

' Aguardar o servidor subir antes de abrir o navegador
WScript.Sleep 3000

' Abrir o navegador padrao na porta do app
oShell.Run "cmd /c start http://localhost:8501", 0, False
