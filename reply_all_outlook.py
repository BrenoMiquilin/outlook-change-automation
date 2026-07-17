import pyautogui
import pyperclip
import time
import re
import win32com.client

# ============================================================
#  CONFIGURAÇÕES — ajuste antes de usar
# ============================================================

TEMPO_INICIAL   = 5
POS_TITULO      = (4104, 536)   # campo título no card do Planner
POS_DATA        = (4248, 896)   # campo data no card do Planner

ENVIO_DIRETO    = False  # True = envia automático | False = abre para revisar

# ============================================================
#  FUNÇÕES AUXILIARES
# ============================================================

def aguardar(segundos, motivo=""):
    if motivo:
        print(f"  ⏳ Aguardando {segundos}s — {motivo}")
    time.sleep(segundos)

def copiar_campo(posicao, nome_campo):
    print(f"  🖱️  Clicando no campo: {nome_campo}")
    pyautogui.click(posicao)
    pyautogui.click(posicao)
    aguardar(0.5)
    pyautogui.hotkey('ctrl', 'a')
    aguardar(0.3)
    pyautogui.hotkey('ctrl', 'c')
    aguardar(0.3)
    valor = pyperclip.paste()
    print(f"  ✅ {nome_campo} capturado: {valor[:80]}")
    return valor.strip()

def extrair_chg(titulo):
    """Pega a primeira palavra do título — que é sempre o número da change."""
    return titulo.split()[0] if titulo else titulo

import html

def montar_html(titulo, data):
    chg = extrair_chg(titulo)

    # Escapa caracteres especiais HTML
    titulo_html = html.escape(titulo)
    data_html = html.escape(data)

    style_14pt = "font-family:'Aptos',Calibri,sans-serif; font-size:14pt; color:black; line-height:150%; mso-line-height-rule:exactly;"
    style_18pt = "font-family:'Aptos',Calibri,sans-serif; font-size:18pt; font-weight:bold; color:black; margin-bottom:14pt;"

    return f"""
<div style="font-family:'Aptos',Calibri,sans-serif;">
    <!-- Título 18pt -->
    <p style="{style_18pt}">
        <span style="font-size:18pt;">Scheduling da Change: {chg}</span>
    </p>

    <!-- Corpo 14pt -->
    <p style="{style_14pt}">
        <span style="font-size:14pt;">Prezados, informo que a Change 
        <b style="color:#2e7d32;">foi aprovada e liberada para execução</b>.</span>
    </p>

    <p style="{style_14pt}">
        <span style="font-size:14pt;">Por gentileza, enviar neste e-mail ou anexar na própria Change todas as evidências da execução (com data e horário).</span>
    </p>

    <p style="{style_14pt}">
        <span style="font-size:14pt;"><b>Título da Change:</b></span><br>
        <span style="font-size:14pt; font-weight:bold;">{titulo_html}</span>
    </p>

    <p style="{style_14pt}">
        <span style="font-size:14pt;"><b>Planned start date:</b></span><br>
        <span style="font-size:14pt;">{data_html}</span>
    </p>
</div>
"""

def buscar_email_outlook(titulo):
    print("  🔍 Conectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    for store in namespace.Stores:
        try:
            sent = store.GetDefaultFolder(5)
            items = sent.Items
            items.Sort("[SentOn]", True)

            for item in items:
                try:
                    if titulo.lower() in item.Subject.lower():
                        print(f"  ✅ E-mail encontrado: {item.Subject[:70]}")
                        return item
                except:
                    continue
        except:
            continue

    return None

# ============================================================
#  FLUXO PRINCIPAL
# ============================================================

def main():
    print("=" * 55)
    print("  REPLY ALL AUTOMÁTICO — Outlook COM + Planner")
    print("=" * 55)
    print(f"\n⚠️  Você tem {TEMPO_INICIAL} segundos para:")
    print("   1. Deixar o card do Planner aberto e visível")
    print("   2. NÃO mover o mouse até o script começar\n")

    for i in range(TEMPO_INICIAL, 0, -1):
        print(f"   Iniciando em {i}...", end='\r')
        time.sleep(1)
    print("\n")

    print("📅 ETAPA 1 — Capturando data do card...")
    data = copiar_campo(POS_DATA, "Data")
    aguardar(0.5)

    print("\n📋 ETAPA 2 — Capturando título do card...")
    titulo = copiar_campo(POS_TITULO, "Título")
    aguardar(0.5)

    print(f"\n✅ Dados capturados:")
    print(f"   Título : {titulo}")
    print(f"   Data   : {data}")
    print(f"   CHG    : {extrair_chg(titulo)}")

    print("\n🔍 ETAPA 3 — Buscando e-mail nos Itens Enviados...")
    email_original = buscar_email_outlook(titulo)

    if not email_original:
        print("\n  ❌ E-mail não encontrado!")
        print("  💡 Verifique se o título do card bate com o assunto do e-mail enviado.")
        return

    print("\n📝 ETAPA 4 — Montando Reply All com formatação completa...")
    reply = email_original.ReplyAll()
    html_novo = montar_html(titulo, data)
    reply.HTMLBody = html_novo + "<br><hr style='border:none;border-top:1px solid #ccc;margin:16pt 0;'>" + reply.HTMLBody

    if ENVIO_DIRETO:
        print("\n📤 ETAPA 5 — Enviando diretamente...")
        reply.Send()
        print("\n" + "=" * 55)
        print("  ✅ E-mail enviado com sucesso!")
    else:
        print("\n📨 ETAPA 5 — Abrindo rascunho para revisão...")
        reply.Display()
        print("\n" + "=" * 55)
        print("  ✅ Rascunho aberto! Revise e envie manualmente.")
        print("  💡 Mude ENVIO_DIRETO = True para envio automático.")

    print("=" * 55)

if __name__ == "__main__":
    main()