#!/usr/bin/env python3
"""
CIURA/PPCIRA 2026 — Criação automática do Microsoft Forms
Compatível com macOS, Linux e Windows.

PRÉ-REQUISITOS:
  Python 3.8+ (já incluído no macOS 12+)
  As dependências são instaladas automaticamente na primeira execução.

COMO USAR (macOS):
  1. Abra o Terminal (Finder → Aplicações → Utilitários → Terminal)
  2. Navegue para a pasta onde guardou este ficheiro:
       cd ~/Downloads
  3. Execute:
       python3 criar-form-ppcira-2026.py
  4. O terminal mostrará um código e um link — abra o link no browser,
     introduza o código e faça login com a sua conta Microsoft 365
  5. Aguarde ~30 segundos — o URL do formulário aparece no terminal

COMO OBTER RESPOSTAS EM EXCEL:
  forms.microsoft.com → formulário → separador "Respostas" → "Abrir no Excel"
"""

import subprocess, sys, json, time

# ── Instalar dependências automaticamente ──────────────────────────────────
def instalar_se_necessario(pacote):
    try:
        __import__(pacote)
    except ImportError:
        print(f"  A instalar {pacote}...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", pacote],
            stdout=subprocess.DEVNULL
        )

print("\n[1/4] A verificar dependências...")
instalar_se_necessario("msal")
instalar_se_necessario("requests")
print("      OK")

import msal        # noqa: E402
import requests    # noqa: E402


# ── Autenticação Microsoft Graph ───────────────────────────────────────────
# Client ID público do "Microsoft Graph Command Line Tools" — não requer
# registo de aplicação no Azure AD.
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES    = ["https://graph.microsoft.com/Forms.ReadWrite"]
BASE_URL  = "https://graph.microsoft.com/beta/me/forms"


def obter_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    # Tentar reutilizar token em cache
    contas = app.get_accounts()
    if contas:
        resultado = app.acquire_token_silent(SCOPES, account=contas[0])
        if resultado and "access_token" in resultado:
            return resultado["access_token"]

    # Device code flow — funciona sem abrir browser automaticamente
    fluxo = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in fluxo:
        raise RuntimeError("Não foi possível iniciar a autenticação: " + str(fluxo))

    print("\n" + "="*60)
    print("AUTENTICAÇÃO MICROSOFT 365")
    print("="*60)
    print(fluxo["message"])
    print("="*60 + "\n")

    resultado = app.acquire_token_by_device_flow(fluxo)
    if "access_token" not in resultado:
        raise RuntimeError(
            "Autenticação falhada: " +
            resultado.get("error_description", resultado.get("error", "erro desconhecido"))
        )
    return resultado["access_token"]


print("\n[2/4] A autenticar no Microsoft 365...")
token = obter_token()
HEADERS = {
    "Authorization": f"Bearer {token}",
    "Content-Type":  "application/json; charset=utf-8",
}
print("      OK — autenticado com sucesso.")


# ── Criar formulário base ──────────────────────────────────────────────────
print("\n[3/4] A criar o formulário no Microsoft Forms...")

descricao = (
    "Promovido pelo PPCIRA — Programa de Prevenção e Controlo de Infeções e de Resistência "
    "aos Antimicrobianos da Direção-Geral da Saúde, em colaboração com o CIURA — Colégio da "
    "Competência em Controlo de Infeção e Uso Racional de Antibióticos da Ordem dos Médicos.\n\n"
    "▸ Destinado a: Diretores/Coordenadores das estruturas locais S/UL-PPCIRA\n"
    "▸ Uma única resposta por instituição\n"
    "▸ Tempo estimado: ~30 minutos\n"
    "▸ Prazo de resposta: 30 de junho de 2026\n"
    "▸ Respostas confidenciais, analisadas em agregado"
)

resp = requests.post(BASE_URL, headers=HEADERS, json={
    "title":       "Inquérito Nacional CIURA/PPCIRA 2026",
    "description": descricao,
    "settings":    {"isAnonymous": True},
})
resp.raise_for_status()
form_id  = resp.json()["id"]
form_url = f"https://forms.microsoft.com/Pages/ResponsePage.aspx?id={form_id}"
print(f"      Formulário base criado (ID: {form_id})")


# ── Funções auxiliares ─────────────────────────────────────────────────────
def pergunta(tipo, titulo, ajuda="", opcoes=None, obrigatorio=False,
             escala_min=1, escala_max=5, rot_min="Não limitante", rot_max="Muito crítico"):
    """Adiciona uma pergunta ao formulário."""
    q = {
        "displayName": titulo,
        "isRequired":  obrigatorio,
        "description": ajuda,
    }

    if tipo == "text":
        q["questionType"] = "text"
    elif tipo == "paragraph":
        q["questionType"] = "text"
        q["textProperties"] = {"isParagraph": True}
    elif tipo in ("radio", "checkbox"):
        q["questionType"] = "multipleChoice"
        q["multipleChoiceProperties"] = {
            "choices":                 [{"displayName": o} for o in (opcoes or [])],
            "allowsMultipleSelection": tipo == "checkbox",
        }
    elif tipo == "dropdown":
        q["questionType"] = "dropdown"
        q["dropdownProperties"] = {
            "choices": [{"displayName": o} for o in (opcoes or [])],
        }
    elif tipo == "scale":
        q["questionType"] = "scale"
        q["scaleProperties"] = {
            "minimum":      escala_min,
            "maximum":      escala_max,
            "minimumLabel": rot_min,
            "maximumLabel": rot_max,
        }

    r = requests.post(
        f"{BASE_URL}/{form_id}/questions",
        headers=HEADERS,
        data=json.dumps(q, ensure_ascii=False).encode("utf-8"),
    )
    if not r.ok:
        print(f"  ⚠  Pergunta ignorada: {titulo[:60]}  ({r.status_code})")
    time.sleep(0.15)   # evitar rate-limiting da API


def secao(titulo, ajuda=""):
    """Adiciona um separador de secção."""
    r = requests.post(
        f"{BASE_URL}/{form_id}/questions",
        headers=HEADERS,
        data=json.dumps({
            "questionType": "sectionHeader",
            "displayName":  titulo,
            "description":  ajuda,
        }, ensure_ascii=False).encode("utf-8"),
    )
    if not r.ok:
        print(f"  ⚠  Secção ignorada: {titulo[:60]}")
    time.sleep(0.15)


# ══════════════════════════════════════════════════════════════════
# SECÇÃO I — CARACTERIZAÇÃO E ENQUADRAMENTO INSTITUCIONAL
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO I — Caracterização e Enquadramento Institucional",
      "Identificação e estrutura da instituição respondente")

pergunta("text",  "1. Nome da instituição", obrigatorio=True)
pergunta("radio", "2. Setor da instituição",
         opcoes=["Pública","Privada","Social","Outra"], obrigatorio=True)
pergunta("checkbox", "3. Tipologia de cuidados (selecione todas as aplicáveis)",
         opcoes=["Hospital","Cuidados de Saúde Primários (CSP)",
                 "Cuidados Continuados","Cuidados Paliativos","Outra"],
         obrigatorio=True)
pergunta("dropdown", "4. Região de Saúde",
         opcoes=["Norte","Centro","Lisboa e Vale do Tejo (LVT)",
                 "Alentejo","Algarve",
                 "Região Autónoma da Madeira (RAM)",
                 "Região Autónoma dos Açores (RAA)"],
         obrigatorio=True)

for cod, txt in [("5a","N.º de camas de internamento agudo"),
                 ("5b","N.º de camas de UCI"),
                 ("5c","N.º de camas de cuidados continuados"),
                 ("5d","Taxa de ocupação global em 2025 (%)"),
                 ("5e","População inscrita nos CSP — n.º utentes (apenas ULS)")]:
    pergunta("text", f"{cod}. {txt}", ajuda="Deixe em branco se não aplicável")

pergunta("radio", "6. O S/UL-PPCIRA está formalmente nomeado pelo órgão máximo de gestão?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("radio", "6.1 Tipo de estrutura formal do S/UL-PPCIRA",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 6',
         opcoes=["Serviço autónomo","Unidade Funcional Autónoma",
                 "Unidade Funcional Integrada","Estrutura de Apoio Técnico",
                 "Comissão","Outro"])

pergunta("radio", "7. O S/UL-PPCIRA possui Regulamento aprovado?",
         opcoes=["Sim","Em elaboração","Não"], obrigatorio=True)
pergunta("dropdown", "8. O S/UL-PPCIRA reporta diretamente a:",
         opcoes=["Conselho de Administração","Direção Clínica",
                 "Direção Clínica e Direção de Enfermagem",
                 "Direção de Enfermagem","Departamento","Outro"],
         obrigatorio=True)
pergunta("radio", "9. É contratualizado anualmente um Plano de Atividades do S/UL-PPCIRA?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("radio", "10. Existe relatório anual de atividades do S/UL-PPCIRA?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("radio", "11. Os objetivos estratégicos institucionais contemplam indicadores PCIRA?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("paragraph", "11.1 Especifique os indicadores PCIRA nos objetivos estratégicos",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 11')
pergunta("radio", "12. O S/UL-PPCIRA é centro de custos autónomo?",
         opcoes=["Sim","Não"], obrigatorio=True)


# ══════════════════════════════════════════════════════════════════
# SECÇÃO II — RECURSOS HUMANOS E ESTRUTURA DE FUNCIONAMENTO
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO II — Recursos Humanos e Estrutura de Funcionamento",
      "Composição da equipa, nomeação e organização interna")

for cat in ["Diretor(a)","Coordenador(a)","Enfermeiro(a) Gestor(a)",
            "Médico(s)","Enfermeiro(s)","Farmacêutico(s)","Outro(s)"]:
    pergunta("text", f"13. {cat} — N.º de elementos")
    pergunta("text", f"13. {cat} — Especialidade / Área profissional")
    pergunta("text", f"13. {cat} — Horas semanais contratualizadas (h)")
    pergunta("text", f"13. {cat} — Horas semanais efetivas (h)")

pergunta("radio", "13.1 O S/UL-PPCIRA tem Diretor ou Coordenador formalmente nomeado?",
         opcoes=["Sim","Não"])
pergunta("radio", "13.2 A nomeação formal abrange todos os elementos da equipa?",
         opcoes=["Sim, todos","Parcialmente","Não"])
pergunta("radio", "13.3 A equipa cumpre os requisitos mínimos do Despacho n.º 10901/2022?",
         opcoes=["Sim","Não"])
pergunta("checkbox", "13.3.1 Identifique o(s) défice(s) existente(s)",
         ajuda='↳ Responda apenas se respondeu "Não" à questão 13.3',
         opcoes=["Défice de tempo médico","Défice de tempo de enfermagem",
                 "Ausência de microbiologista","Ausência de farmacêutico",
                 "Ausência de horário completo do Diretor",
                 "Ausência de competências em qualidade","Outro"])

pergunta("radio", "14. Existem elementos dinamizadores/elos PCIRA nos serviços formalmente designados?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("dropdown", "14.1 Percentagem de serviços com elos PCIRA designados",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 14',
         opcoes=["0–25%","26–50%","51–75%","76–100%"])

for cat in ["Médicos","Enfermeiros","Outros"]:
    pergunta("text", f"15. {cat} — % horas semanais dedicadas a PCI")
    pergunta("text", f"15. {cat} — % horas semanais dedicadas a PAPA")

pergunta("dropdown", "16. Periodicidade de reunião da equipa nuclear do S/UL-PPCIRA",
         opcoes=["Semanal","Quinzenal","Mensal","Trimestral",
                 "Sem periodicidade definida","Outra"])
pergunta("dropdown", "17. Periodicidade de reunião com o órgão máximo de gestão",
         opcoes=["Mensal","Trimestral","Semestral","Anual","Só quando solicitado","Outra"])
pergunta("dropdown", "18. Periodicidade de reunião com os principais serviços clínicos",
         opcoes=["Mensal","Trimestral","Semestral","Anual","Só quando solicitado","Outra"])
pergunta("radio", "19. É realizada a contratualização de objetivos PPCIRA com os serviços clínicos?",
         opcoes=["Sim, em todos os serviços","Sim, em alguns serviços","Não"])
pergunta("dropdown", "19.1 Percentagem de serviços com objetivos PPCIRA contratualizados",
         ajuda='↳ Responda apenas se respondeu "Sim, em alguns serviços" à questão 19',
         opcoes=["0–25%","26–50%","51–75%","76–100%"])


# ══════════════════════════════════════════════════════════════════
# SECÇÃO III — ATIVIDADES DE VIGILÂNCIA EPIDEMIOLÓGICA
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO III — Atividades de Vigilância Epidemiológica",
      "Programas de VE, PAPA e participação em projetos nacionais/internacionais")

pergunta("radio", "20. Existem programas estruturados de VE nas dimensões PCIRA?",
         opcoes=["Sim, em todas as dimensões","Parcialmente","Não"], obrigatorio=True)
pergunta("checkbox", "20.1 Dimensões abrangidas pelos programas de VE",
         ajuda='↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20',
         opcoes=["PBCI","Uso de luvas","Higiene das mãos","ILC",
                 "ILC — cirurgia específica (ECDC)","ICSRCVC","Bacteriémias",
                 "Pneumonia associada a intubação (PAI)","ITUACV",
                 "Microrganismos multirresistentes (MMR)","Clostridioides difficile",
                 "Consumo de antimicrobianos","Qualidade da prescrição (PAPA)","Outros"])
pergunta("dropdown", "20.2 Percentagem de serviços abrangidos pela VE",
         ajuda='↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20',
         opcoes=["0–25%","26–50%","51–75%","76–100%"])
pergunta("checkbox", "20.3 Sistema de informação utilizado para a VE",
         ajuda='↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20',
         opcoes=["Plataforma DGS","Hepic®","Meliora®","Outro"])
pergunta("radio", "20.4 Existe feedback estruturado e regular dos dados de VE para os serviços?",
         opcoes=["Sim","Não"])
pergunta("paragraph", "20.4.1 Frequência e método do feedback dos dados de VE",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 20.4')

pergunta("radio", "21. O PAPA — Programa de Apoio à Prescrição de Antibióticos está em funcionamento?",
         opcoes=["Sim, em todos os serviços","Parcialmente","Não"], obrigatorio=True)
pergunta("checkbox", "21.1 Em que valências/contextos está implementado o PAPA?",
         ajuda='↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21',
         opcoes=["CSP","Internamento hospitalar",
                 "Ambulatório / Hospital de dia / Consultas","UCCI","SMI","Urgência"])
pergunta("checkbox", "21.2 Componentes do PAPA implementadas",
         ajuda='↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21',
         opcoes=["Revisão/restrição de prescrições até 72h",
                 "Protocolos institucionais de antibióticos",
                 "Formação contínua de profissionais",
                 "Retorno de métricas por unidade/serviço/prescritor",
                 "Projetos de intervenção comportamental (nudge)","Outros"])
pergunta("dropdown", "22. Frequência do feedback dos dados PAPA aos prescritores",
         opcoes=["Mensal","Trimestral","Semestral","Anual","Não é realizado"])
pergunta("radio", "23. Existe plano institucional de avaliação e resposta a surtos?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("radio", "24. Participa em projeto de melhoria da qualidade nacional (STOP-Infeção, Drive AMS, ITUCCI...)?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("paragraph", "24.1 Identifique os projetos nacionais em que participa",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 24')
pergunta("radio", "25. Participa em outros grupos de trabalho ou projetos nacionais/internacionais de PCIRA?",
         opcoes=["Sim","Não"])
pergunta("paragraph", "25.1 Identifique os grupos de trabalho/projetos",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 25')


# ══════════════════════════════════════════════════════════════════
# SECÇÃO IV — FORMAÇÃO, COMPETÊNCIAS E DESENVOLVIMENTO PROFISSIONAL
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO IV — Formação, Competências e Desenvolvimento Profissional",
      "Capacitação da equipa e certificações profissionais")

pergunta("radio", "26. Existem programas formais de capacitação e treino em PCIRA?",
         opcoes=["Sim, com carácter obrigatório","Sim, com carácter opcional","Não"],
         obrigatorio=True)
pergunta("dropdown", "26.1 Periodicidade da formação PCIRA para os profissionais da instituição",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 26',
         opcoes=["Anual","De 2 em 2 anos","De 3 em 3 anos","Sem periodicidade definida"])
pergunta("radio", "27. A formação contínua e progressiva da equipa S/UL-PPCIRA está assegurada?",
         opcoes=["Sim","Não"])
pergunta("paragraph", "27.1 Como é assegurada a formação contínua da equipa S/UL-PPCIRA?",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 27')

for nivel in ["Doutoramento em PCIRA","Mestrado em PCIRA","Pós-Graduação em PCIRA",
              "Certificação europeia PCI ou PAPA (EUCIC/ESCMID)",
              "Outros cursos de formação diferenciada"]:
    for cat in ["Médicos","Enfermeiros","Outros profissionais"]:
        pergunta("text", f"28. {nivel} — {cat} (n.º)")

pergunta("radio", "29. Existem profissionais com Competência pela Ordem Profissional em PCI e/ou PAPA?",
         opcoes=["Sim","Não"])
pergunta("text", "29.1a N.º de médicos com CCIURA (Ordem dos Médicos)",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 29')
pergunta("text", "29.1b N.º de enfermeiros com CADEPCI (Ordem dos Enfermeiros)",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 29')
pergunta("radio", "30. O plano de desenvolvimento profissional inclui investigação e publicações em PCIRA?",
         opcoes=["Sim","Não"])


# ══════════════════════════════════════════════════════════════════
# SECÇÃO V — ENVOLVIMENTO INSTITUCIONAL E INFRAESTRUTURAS
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO V — Envolvimento Institucional e Infraestruturas",
      "Espaço físico, apoio administrativo, comissões e laboratório")

pergunta("radio", "31. O S/UL-PPCIRA dispõe de espaço físico próprio (gabinete dedicado)?",
         opcoes=["Sim","Não"], obrigatorio=True)
pergunta("radio", "32. O S/UL-PPCIRA dispõe de apoio administrativo alocado?",
         opcoes=["Sim, exclusivo","Sim, partilhado","Não"], obrigatorio=True)
pergunta("checkbox", "33. O S/UL-PPCIRA integra formalmente as seguintes comissões:",
         opcoes=["Comissão de Farmácia e Terapêutica (CFT)",
                 "Comissão/Unidade de Qualidade e Segurança do Doente",
                 "Outras comissões ou grupos de trabalho",
                 "Não integra nenhuma comissão formalmente"])
pergunta("radio", "34. O S/UL-PPCIRA dispõe de laboratório de microbiologia interno?",
         opcoes=["Sim, sem descontinuidade","Sim, com interrupção/cobertura parcial",
                 "Não (laboratório externo)"],
         obrigatorio=True)
pergunta("radio", "35. Os CSP trabalham com que tipo de laboratório de microbiologia?",
         ajuda="Responda apenas se a sua instituição inclui CSP",
         opcoes=["Laboratório interno","Laboratório externo","Não aplicável"])
pergunta("radio", "36. Os CSP utilizam testes de diagnóstico rápido microbiológico (Point-of-Care)?",
         ajuda="Responda apenas se a sua instituição inclui CSP",
         opcoes=["Sim","Não","Não aplicável"])
pergunta("paragraph", "36.1 Especifique os testes Point-of-Care utilizados nos CSP",
         ajuda='↳ Responda apenas se respondeu "Sim" à questão 36')
pergunta("dropdown", "37a. Carta microbiológica hospitalar — periodicidade de atualização",
         opcoes=["Semestral","Anual","De 2 em 2 anos","Não existe","Não aplicável"])
pergunta("dropdown", "37b. Carta microbiológica comunitária (CSP) — periodicidade de atualização",
         opcoes=["Semestral","Anual","De 2 em 2 anos","Não existe","Não aplicável"])


# ══════════════════════════════════════════════════════════════════
# SECÇÃO VI — PERCEÇÃO PRÁTICA
# ══════════════════════════════════════════════════════════════════
secao("SECÇÃO VI — Perceção Prática",
      "Avaliação global e sugestões para o PPCIRA/DGS")

pergunta("radio", "38. Classificação global das condições institucionais para implementação do PPCIRA?",
         opcoes=["Elevadas","Moderadas","Baixas"], obrigatorio=True)
pergunta("paragraph", "38.1 Justificação da classificação atribuída",
         ajuda="Descreva os principais motivos para a classificação escolhida")

secao("39. Fatores limitantes — grau de criticidade para a sua instituição",
      "Para cada fator, classifique de 1 (não é limitante) a 5 (fator muito crítico)")

for fator in [
    "Falta de priorização pelo Conselho de Administração",
    "Falta de priorização pelos Responsáveis de Serviço",
    "Falta de recursos humanos com tempo protegido para PCIRA",
    "Falta de sistemas de informação adequados",
    "Falta de formação contínua da equipa S/UL-PPCIRA",
    "Falta de formação dos demais profissionais de saúde",
    "Resistência à mudança / baixa adesão clínica",
    "Inexistência de infraestrutura laboratorial com resposta adequada",
]:
    pergunta("scale", f"39. {fator}",
             ajuda="1 = Não é fator limitante   |   5 = Fator muito crítico",
             escala_min=1, escala_max=5,
             rot_min="Não limitante", rot_max="Muito crítico")

pergunta("text", "39. Outro fator limitante — especificação",
         ajuda="Se identificou outro fator relevante, descreva-o aqui")

secao("40. Sugestões para o PPCIRA/DGS — próximos 3 anos",
      "Indique até três intervenções ou iniciativas que considera prioritárias")

pergunta("paragraph", "40. Sugestão 1 (obrigatória)", obrigatorio=True)
pergunta("paragraph", "40. Sugestão 2 (opcional)")
pergunta("paragraph", "40. Sugestão 3 (opcional)")
pergunta("paragraph", "Observações adicionais",
         ajuda="Comentários ou informações adicionais que considere relevantes")


# ── Resultado final ────────────────────────────────────────────────────────
print("\n[4/4] Concluído!")
print()
print("=" * 62)
print("  FORMULÁRIO CRIADO COM SUCESSO")
print("=" * 62)
print()
print("  LINK PARA PARTILHAR COM AS INSTITUIÇÕES:")
print(f"  {form_url}")
print()
print("  EDITAR O FORMULÁRIO:")
print("  forms.microsoft.com  →  Os meus formulários")
print()
print("  EXPORTAR RESPOSTAS PARA EXCEL:")
print("  Formulário → separador Respostas → Abrir no Excel")
print("=" * 62)
print()

# Copiar URL para área de transferência (macOS)
try:
    subprocess.run(["pbcopy"], input=form_url.encode(), check=True)
    print("  O link foi copiado para a área de transferência (Command+V para colar).")
except Exception:
    pass
