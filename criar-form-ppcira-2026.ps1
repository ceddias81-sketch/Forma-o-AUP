#Requires -Version 5.1
<#
.SYNOPSIS
    CIURA/PPCIRA 2026 — Criação automática do Microsoft Forms via Microsoft Graph API

.DESCRIPTION
    Cria automaticamente o Inquérito Nacional CIURA/PPCIRA 2026 completo no Microsoft Forms
    da conta Microsoft 365 do utilizador que executar o script.

.PRE-REQUISITOS
    1. Windows com PowerShell 5.1+ ou PowerShell 7+
    2. Conta Microsoft 365 (conta institucional ou pessoal com Forms)
    3. Ligação à internet
    O módulo Microsoft.Graph será instalado automaticamente se não estiver presente.

.COMO USAR
    1. Abra o PowerShell como Administrador (ou Utilizador, com -Scope CurrentUser)
    2. Navegue para a pasta onde guardou este ficheiro:  cd C:\caminho\do\ficheiro
    3. Execute:  .\criar-form-ppcira-2026.ps1
    4. Uma janela do browser abrirá para autenticação Microsoft — faça login com
       a conta Microsoft 365 onde quer criar o formulário
    5. Aguarde ~30 segundos — o URL do formulário aparecerá no terminal

.COMO OBTER RESPOSTAS EM EXCEL
    1. Aceda ao formulário criado em forms.microsoft.com
    2. Separador "Respostas" → botão "Abrir no Excel"
    3. Um ficheiro .xlsx é transferido com 1 linha por resposta e 1 coluna por pergunta

.NOTAS
    A API do Microsoft Forms (Graph beta) está em pré-visualização. Em caso de erro,
    consulte o guia de criação manual no final deste ficheiro.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# 1. INSTALAR / IMPORTAR MÓDULO MICROSOFT.GRAPH
# ─────────────────────────────────────────────
Write-Host "`n[1/4] A verificar módulo Microsoft.Graph..." -ForegroundColor Cyan

$modulo = 'Microsoft.Graph'
if (-not (Get-Module -ListAvailable -Name $modulo -ErrorAction SilentlyContinue)) {
    Write-Host "     Módulo não encontrado. A instalar (pode demorar ~1-2 minutos)..." -ForegroundColor Yellow
    Install-Module $modulo -Scope CurrentUser -Force -AllowClobber
}
Import-Module $modulo -ErrorAction Stop
Write-Host "     OK — módulo disponível." -ForegroundColor Green


# ─────────────────────────────────────────────
# 2. AUTENTICAÇÃO
# ─────────────────────────────────────────────
Write-Host "`n[2/4] A autenticar no Microsoft 365..." -ForegroundColor Cyan
Write-Host "     Uma janela do browser abrirá para login — use a conta onde quer criar o formulário." -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes "Forms.ReadWrite" -NoWelcome -ErrorAction Stop
} catch {
    # Fallback para scope alternativo
    Connect-MgGraph -Scopes "Forms.ReadWrite.All" -NoWelcome -ErrorAction Stop
}
Write-Host "     OK — autenticado com sucesso." -ForegroundColor Green


# ─────────────────────────────────────────────
# 3. CRIAR FORMULÁRIO
# ─────────────────────────────────────────────
Write-Host "`n[3/4] A criar o formulário no Microsoft Forms..." -ForegroundColor Cyan

$BASE_URL = "https://graph.microsoft.com/beta/me/forms"

$descricao = @"
Promovido pelo PPCIRA — Programa de Prevenção e Controlo de Infeções e de Resistência aos Antimicrobianos da Direção-Geral da Saúde, em colaboração com o CIURA — Colégio da Competência em Controlo de Infeção e Uso Racional de Antibióticos da Ordem dos Médicos.

▸ Destinado a: Diretores/Coordenadores das estruturas locais S/UL-PPCIRA
▸ Uma única resposta por instituição
▸ Tempo estimado: ~30 minutos
▸ Prazo de resposta: 30 de junho de 2026
▸ Respostas confidenciais, analisadas em agregado
"@

$formBody = @{
    title       = "Inquérito Nacional CIURA/PPCIRA 2026"
    description = $descricao
    settings    = @{
        isAnonymous          = $true
        isResponseCountHidden = $false
    }
} | ConvertTo-Json -Depth 3

$form = Invoke-MgGraphRequest -Method POST -Uri $BASE_URL `
        -Body $formBody -ContentType "application/json; charset=utf-8"

$formId  = $form.id
$formUrl = "https://forms.microsoft.com/Pages/ResponsePage.aspx?id=$formId"
Write-Host "     Formulário base criado  (ID: $formId)" -ForegroundColor DarkGray

# ─── Função auxiliar para adicionar perguntas ───────────────────────────────
function Add-FormQuestion {
    param(
        [string]$Type,          # text | paragraph | multiChoice | checkbox | dropdown | scale
        [string]$Title,
        [string]$Help = '',
        [string[]]$Choices = @(),
        [bool]$Required = $false,
        [int]$ScaleMin = 1,
        [int]$ScaleMax = 5,
        [string]$ScaleMinLabel = '',
        [string]$ScaleMaxLabel = ''
    )

    $q = @{
        displayName  = $Title
        isRequired   = $Required
        description  = $Help
    }

    switch ($Type) {
        'text'      { $q.questionType = 'text' }
        'paragraph' { $q.questionType = 'text'; $q.textProperties = @{ isParagraph = $true } }
        'multiChoice' {
            $q.questionType = 'multipleChoice'
            $q.multipleChoiceProperties = @{
                choices                = $Choices | ForEach-Object { @{ displayName = $_ } }
                allowsMultipleSelection = $false
            }
        }
        'checkbox' {
            $q.questionType = 'multipleChoice'
            $q.multipleChoiceProperties = @{
                choices                = $Choices | ForEach-Object { @{ displayName = $_ } }
                allowsMultipleSelection = $true
            }
        }
        'dropdown' {
            $q.questionType = 'dropdown'
            $q.dropdownProperties = @{
                choices = $Choices | ForEach-Object { @{ displayName = $_ } }
            }
        }
        'scale' {
            $q.questionType = 'scale'
            $q.scaleProperties = @{
                minimum      = $ScaleMin
                maximum      = $ScaleMax
                minimumLabel = $ScaleMinLabel
                maximumLabel = $ScaleMaxLabel
            }
        }
    }

    $body = $q | ConvertTo-Json -Depth 5 -EscapeHandling EscapeNonAscii
    try {
        Invoke-MgGraphRequest -Method POST `
            -Uri "$BASE_URL/$formId/questions" `
            -Body $body -ContentType "application/json; charset=utf-8" | Out-Null
    } catch {
        Write-Warning "Pergunta não adicionada: $Title`n  Erro: $($_.Exception.Message)"
    }
}

# ─── Função auxiliar para adicionar separador de secção ─────────────────────
function Add-Section {
    param([string]$Title, [string]$Help = '')
    $body = @{
        questionType = 'sectionHeader'
        displayName  = $Title
        description  = $Help
    } | ConvertTo-Json
    try {
        Invoke-MgGraphRequest -Method POST `
            -Uri "$BASE_URL/$formId/questions" `
            -Body $body -ContentType "application/json; charset=utf-8" | Out-Null
    } catch {
        Write-Warning "Separador de secção não adicionado: $Title"
    }
}


# ════════════════════════════════════════════════════════════════
# SECÇÃO I — CARACTERIZAÇÃO E ENQUADRAMENTO INSTITUCIONAL
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO I — Caracterização e Enquadramento Institucional" `
            -Help  "Identificação e estrutura da instituição respondente"

Add-FormQuestion -Type 'text' -Title '1. Nome da instituição' -Required $true

Add-FormQuestion -Type 'multiChoice' -Title '2. Setor da instituição' -Required $true `
    -Choices 'Pública','Privada','Social','Outra'

Add-FormQuestion -Type 'checkbox' -Title '3. Tipologia de cuidados (selecione todas as aplicáveis)' -Required $true `
    -Choices 'Hospital','Cuidados de Saúde Primários (CSP)','Cuidados Continuados','Cuidados Paliativos','Outra'

Add-FormQuestion -Type 'dropdown' -Title '4. Região de Saúde' -Required $true `
    -Choices 'Norte','Centro','Lisboa e Vale do Tejo (LVT)','Alentejo','Algarve',`
             'Região Autónoma da Madeira (RAM)','Região Autónoma dos Açores (RAA)'

Add-FormQuestion -Type 'text' -Title '5a. N.º de camas de internamento agudo (2025)' `
    -Help 'Deixe em branco se não aplicável'
Add-FormQuestion -Type 'text' -Title '5b. N.º de camas de UCI (2025)' `
    -Help 'Deixe em branco se não aplicável'
Add-FormQuestion -Type 'text' -Title '5c. N.º de camas de cuidados continuados (2025)' `
    -Help 'Deixe em branco se não aplicável'
Add-FormQuestion -Type 'text' -Title '5d. Taxa de ocupação global em 2025 (%)' `
    -Help 'Deixe em branco se não aplicável'
Add-FormQuestion -Type 'text' -Title '5e. População inscrita nos CSP — n.º utentes (apenas ULS)' `
    -Help 'Deixe em branco se não aplicável'

Add-FormQuestion -Type 'multiChoice' -Title '6. O S/UL-PPCIRA está formalmente nomeado pelo órgão máximo de gestão?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'multiChoice' -Title '6.1 Tipo de estrutura formal do S/UL-PPCIRA' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 6' `
    -Choices 'Serviço autónomo','Unidade Funcional Autónoma','Unidade Funcional Integrada',`
             'Estrutura de Apoio Técnico','Comissão','Outro'

Add-FormQuestion -Type 'multiChoice' -Title '7. O S/UL-PPCIRA possui Regulamento aprovado?' -Required $true `
    -Choices 'Sim','Em elaboração','Não'

Add-FormQuestion -Type 'dropdown' -Title '8. O S/UL-PPCIRA reporta diretamente a:' -Required $true `
    -Choices 'Conselho de Administração','Direção Clínica',`
             'Direção Clínica e Direção de Enfermagem','Direção de Enfermagem','Departamento','Outro'

Add-FormQuestion -Type 'multiChoice' -Title '9. É contratualizado anualmente um Plano de Atividades do S/UL-PPCIRA?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'multiChoice' -Title '10. Existe relatório anual de atividades do S/UL-PPCIRA?' -Required $true `
    -Choices 'Sim','Não'

Add-FormQuestion -Type 'multiChoice' -Title '11. Os objetivos estratégicos institucionais contemplam indicadores PCIRA?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'paragraph' -Title '11.1 Especifique os indicadores PCIRA contemplados nos objetivos estratégicos' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 11'

Add-FormQuestion -Type 'multiChoice' -Title '12. O S/UL-PPCIRA é centro de custos autónomo?' -Required $true `
    -Choices 'Sim','Não'


# ════════════════════════════════════════════════════════════════
# SECÇÃO II — RECURSOS HUMANOS E ESTRUTURA DE FUNCIONAMENTO
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO II — Recursos Humanos e Estrutura de Funcionamento" `
            -Help  "Composição da equipa, nomeação e organização interna"

$equipaCategorias = 'Diretor(a)','Coordenador(a)','Enfermeiro(a) Gestor(a)',`
                    'Médico(s)','Enfermeiro(s)','Farmacêutico(s)','Outro(s)'

foreach ($cat in $equipaCategorias) {
    Add-FormQuestion -Type 'text' -Title "13. $cat — N.º de elementos"
    Add-FormQuestion -Type 'text' -Title "13. $cat — Especialidade / Área profissional"
    Add-FormQuestion -Type 'text' -Title "13. $cat — Horas semanais contratualizadas (h)"
    Add-FormQuestion -Type 'text' -Title "13. $cat — Horas semanais efetivas (h)"
}

Add-FormQuestion -Type 'multiChoice' -Title '13.1 O S/UL-PPCIRA tem Diretor ou Coordenador formalmente nomeado?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'multiChoice' -Title '13.2 A nomeação formal abrange todos os elementos da equipa?' `
    -Choices 'Sim, todos','Parcialmente','Não'
Add-FormQuestion -Type 'multiChoice' -Title '13.3 A equipa cumpre os requisitos mínimos do Despacho n.º 10901/2022?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'checkbox' -Title '13.3.1 Identifique o(s) défice(s) existente(s)' `
    -Help '↳ Responda apenas se respondeu "Não" à questão 13.3' `
    -Choices 'Défice de tempo médico','Défice de tempo de enfermagem','Ausência de microbiologista',`
             'Ausência de farmacêutico','Ausência de horário completo do Diretor',`
             'Ausência de competências em qualidade','Outro'

Add-FormQuestion -Type 'multiChoice' -Title '14. Existem elementos dinamizadores/elos PCIRA nos serviços formalmente designados?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'dropdown' -Title '14.1 Percentagem de serviços com elos PCIRA designados' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 14' `
    -Choices '0–25%','26–50%','51–75%','76–100%'

foreach ($cat in 'Médicos','Enfermeiros','Outros') {
    Add-FormQuestion -Type 'text' -Title "15. $cat — % horas semanais dedicadas a PCI"
    Add-FormQuestion -Type 'text' -Title "15. $cat — % horas semanais dedicadas a PAPA"
}

Add-FormQuestion -Type 'dropdown' -Title '16. Periodicidade de reunião da equipa nuclear do S/UL-PPCIRA' `
    -Choices 'Semanal','Quinzenal','Mensal','Trimestral','Sem periodicidade definida','Outra'
Add-FormQuestion -Type 'dropdown' -Title '17. Periodicidade de reunião com o órgão máximo de gestão' `
    -Choices 'Mensal','Trimestral','Semestral','Anual','Só quando solicitado','Outra'
Add-FormQuestion -Type 'dropdown' -Title '18. Periodicidade de reunião com os principais serviços clínicos' `
    -Choices 'Mensal','Trimestral','Semestral','Anual','Só quando solicitado','Outra'

Add-FormQuestion -Type 'multiChoice' -Title '19. É realizada a contratualização de objetivos PPCIRA com os serviços clínicos?' `
    -Choices 'Sim, em todos os serviços','Sim, em alguns serviços','Não'
Add-FormQuestion -Type 'dropdown' -Title '19.1 Percentagem de serviços com objetivos PPCIRA contratualizados' `
    -Help '↳ Responda apenas se respondeu "Sim, em alguns serviços" à questão 19' `
    -Choices '0–25%','26–50%','51–75%','76–100%'


# ════════════════════════════════════════════════════════════════
# SECÇÃO III — ATIVIDADES DE VIGILÂNCIA EPIDEMIOLÓGICA
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO III — Atividades de Vigilância Epidemiológica" `
            -Help  "Programas de VE, PAPA e participação em projetos nacionais/internacionais"

Add-FormQuestion -Type 'multiChoice' -Title '20. Existem programas estruturados de vigilância epidemiológica nas dimensões PCIRA?' -Required $true `
    -Choices 'Sim, em todas as dimensões','Parcialmente','Não'

Add-FormQuestion -Type 'checkbox' -Title '20.1 Dimensões abrangidas pelos programas de VE (selecione todas as aplicáveis)' `
    -Help '↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20' `
    -Choices 'Precauções Básicas de Controlo de Infeção (PBCI)','Uso de luvas','Higiene das mãos',`
             'Infeção de Local Cirúrgico (ILC)','ILC — cirurgia específica (ECDC)',`
             'Infeção corrente sanguínea relacionada com CVC (ICSRCVC)','Bacteriémias',`
             'Pneumonia associada a intubação (PAI)','ITUACV',`
             'Microrganismos multirresistentes (MMR)','Clostridioides difficile',`
             'Consumo de antimicrobianos','Qualidade da prescrição (PAPA)','Outros'

Add-FormQuestion -Type 'dropdown' -Title '20.2 Percentagem de serviços abrangidos pela VE' `
    -Help '↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20' `
    -Choices '0–25%','26–50%','51–75%','76–100%'

Add-FormQuestion -Type 'checkbox' -Title '20.3 Sistema de informação utilizado para a VE' `
    -Help '↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20' `
    -Choices 'Plataforma DGS','Hepic®','Meliora®','Outro'

Add-FormQuestion -Type 'multiChoice' -Title '20.4 Existe feedback estruturado e regular dos dados de VE para os serviços?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'paragraph' -Title '20.4.1 Frequência e método do feedback dos dados de VE' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 20.4'

Add-FormQuestion -Type 'multiChoice' -Title '21. O PAPA — Programa de Apoio à Prescrição de Antibióticos está em funcionamento?' -Required $true `
    -Choices 'Sim, em todos os serviços','Parcialmente','Não'

Add-FormQuestion -Type 'checkbox' -Title '21.1 Em que valências/contextos está implementado o PAPA?' `
    -Help '↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21' `
    -Choices 'Cuidados de Saúde Primários (CSP)','Internamento hospitalar',`
             'Ambulatório / Hospital de dia / Consultas','UCCI','SMI','Urgência'

Add-FormQuestion -Type 'checkbox' -Title '21.2 Componentes do PAPA implementadas (selecione todas as aplicáveis)' `
    -Help '↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21' `
    -Choices 'Revisão/restrição de prescrições até 72h','Protocolos institucionais de antibióticos',`
             'Formação contínua de profissionais','Retorno de métricas por unidade/serviço/prescritor',`
             'Projetos de intervenção comportamental (nudge, etc.)','Outros'

Add-FormQuestion -Type 'dropdown' -Title '22. Frequência do feedback dos dados PAPA aos prescritores' `
    -Choices 'Mensal','Trimestral','Semestral','Anual','Não é realizado'

Add-FormQuestion -Type 'multiChoice' -Title '23. Existe plano institucional de avaliação e resposta a surtos?' -Required $true `
    -Choices 'Sim','Não'

Add-FormQuestion -Type 'multiChoice' -Title '24. Participa em projeto de melhoria da qualidade nacional (STOP-Infeção, Drive AMS, ITUCCI...)?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'paragraph' -Title '24.1 Identifique os projetos nacionais em que participa' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 24'

Add-FormQuestion -Type 'multiChoice' -Title '25. Participa em outros grupos de trabalho ou projetos nacionais/internacionais de PCIRA?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'paragraph' -Title '25.1 Identifique os grupos de trabalho/projetos' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 25'


# ════════════════════════════════════════════════════════════════
# SECÇÃO IV — FORMAÇÃO, COMPETÊNCIAS E DESENVOLVIMENTO PROFISSIONAL
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO IV — Formação, Competências e Desenvolvimento Profissional" `
            -Help  "Capacitação da equipa e certificações profissionais"

Add-FormQuestion -Type 'multiChoice' -Title '26. Existem programas formais de capacitação e treino em PCIRA (acolhimento + formação contínua)?' -Required $true `
    -Choices 'Sim, com carácter obrigatório','Sim, com carácter opcional','Não'
Add-FormQuestion -Type 'dropdown' -Title '26.1 Periodicidade da formação PCIRA para os profissionais da instituição' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 26' `
    -Choices 'Anual','De 2 em 2 anos','De 3 em 3 anos','Sem periodicidade definida'

Add-FormQuestion -Type 'multiChoice' -Title '27. A formação contínua e progressiva da equipa S/UL-PPCIRA está assegurada?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'paragraph' -Title '27.1 Como é assegurada a formação contínua da equipa S/UL-PPCIRA?' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 27'

$formacaoNiveis = @(
    'Doutoramento em PCIRA',
    'Mestrado em PCIRA',
    'Pós-Graduação em PCIRA',
    'Certificação europeia PCI ou PAPA (EUCIC / ESCMID)',
    'Outros cursos de formação diferenciada'
)
foreach ($nivel in $formacaoNiveis) {
    foreach ($cat in 'Médicos','Enfermeiros','Outros profissionais') {
        Add-FormQuestion -Type 'text' -Title "28. $nivel — $cat (n.º)"
    }
}

Add-FormQuestion -Type 'multiChoice' -Title '29. Existem profissionais com Competência pela Ordem Profissional em PCI e/ou PAPA?' `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'text' -Title '29.1a N.º de médicos com CCIURA (Ordem dos Médicos)' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 29'
Add-FormQuestion -Type 'text' -Title '29.1b N.º de enfermeiros com CADEPCI (Ordem dos Enfermeiros)' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 29'

Add-FormQuestion -Type 'multiChoice' -Title '30. O plano de desenvolvimento profissional inclui investigação e publicações em PCIRA?' `
    -Choices 'Sim','Não'


# ════════════════════════════════════════════════════════════════
# SECÇÃO V — ENVOLVIMENTO INSTITUCIONAL E INFRAESTRUTURAS
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO V — Envolvimento Institucional e Infraestruturas" `
            -Help  "Espaço físico, apoio administrativo, comissões e laboratório"

Add-FormQuestion -Type 'multiChoice' -Title '31. O S/UL-PPCIRA dispõe de espaço físico próprio (gabinete dedicado)?' -Required $true `
    -Choices 'Sim','Não'
Add-FormQuestion -Type 'multiChoice' -Title '32. O S/UL-PPCIRA dispõe de apoio administrativo alocado?' -Required $true `
    -Choices 'Sim, exclusivo','Sim, partilhado','Não'

Add-FormQuestion -Type 'checkbox' -Title '33. O S/UL-PPCIRA integra formalmente as seguintes comissões/grupos de trabalho:' `
    -Choices 'Comissão de Farmácia e Terapêutica (CFT)',`
             'Comissão/Unidade de Qualidade e Segurança do Doente',`
             'Outras comissões ou grupos de trabalho',`
             'Não integra nenhuma comissão formalmente'

Add-FormQuestion -Type 'multiChoice' -Title '34. O S/UL-PPCIRA dispõe de laboratório de microbiologia interno?' -Required $true `
    -Choices 'Sim, sem descontinuidade de serviço','Sim, com interrupção ou cobertura parcial',`
             'Não (laboratório externo)'

Add-FormQuestion -Type 'multiChoice' -Title '35. Os CSP trabalham com que tipo de laboratório de microbiologia?' `
    -Help 'Responda apenas se a sua instituição inclui CSP' `
    -Choices 'Laboratório interno','Laboratório externo','Não aplicável'

Add-FormQuestion -Type 'multiChoice' -Title '36. Os CSP utilizam testes de diagnóstico rápido microbiológico (Point-of-Care)?' `
    -Help 'Responda apenas se a sua instituição inclui CSP' `
    -Choices 'Sim','Não','Não aplicável'
Add-FormQuestion -Type 'paragraph' -Title '36.1 Especifique os testes Point-of-Care microbiológicos utilizados nos CSP' `
    -Help '↳ Responda apenas se respondeu "Sim" à questão 36'

Add-FormQuestion -Type 'dropdown' -Title '37a. Carta microbiológica hospitalar — periodicidade de atualização' `
    -Choices 'Semestral','Anual','De 2 em 2 anos','Não existe','Não aplicável'
Add-FormQuestion -Type 'dropdown' -Title '37b. Carta microbiológica comunitária (CSP) — periodicidade de atualização' `
    -Choices 'Semestral','Anual','De 2 em 2 anos','Não existe','Não aplicável'


# ════════════════════════════════════════════════════════════════
# SECÇÃO VI — PERCEÇÃO PRÁTICA
# ════════════════════════════════════════════════════════════════
Add-Section -Title "SECÇÃO VI — Perceção Prática" `
            -Help  "Avaliação global das condições institucionais e sugestões para o PPCIRA/DGS"

Add-FormQuestion -Type 'multiChoice' -Title '38. Classificação global das condições institucionais para implementação do PPCIRA?' -Required $true `
    -Choices 'Elevadas','Moderadas','Baixas'
Add-FormQuestion -Type 'paragraph' -Title '38.1 Justificação da classificação atribuída' `
    -Help 'Descreva os principais motivos para a classificação escolhida'

Add-Section -Title "39. Fatores limitantes — grau de criticidade para a sua instituição" `
            -Help  "Para cada fator, classifique de 1 (não é limitante) a 5 (fator muito crítico)"

$fatores = @(
    'Falta de priorização pelo Conselho de Administração',
    'Falta de priorização pelos Responsáveis de Serviço',
    'Falta de recursos humanos com tempo protegido para PCIRA',
    'Falta de sistemas de informação adequados',
    'Falta de formação contínua da equipa S/UL-PPCIRA',
    'Falta de formação dos demais profissionais de saúde',
    'Resistência à mudança / baixa adesão clínica',
    'Inexistência de infraestrutura laboratorial com resposta adequada'
)
foreach ($fator in $fatores) {
    Add-FormQuestion -Type 'scale' -Title "39. $fator" `
        -Help '1 = Não é fator limitante   |   5 = Fator muito crítico' `
        -ScaleMin 1 -ScaleMax 5 `
        -ScaleMinLabel 'Não limitante' -ScaleMaxLabel 'Muito crítico'
}
Add-FormQuestion -Type 'text' -Title '39. Outro fator limitante — especificação' `
    -Help 'Se identificou outro fator relevante, descreva-o aqui'

Add-Section -Title "40. Sugestões para o PPCIRA/DGS — próximos 3 anos" `
            -Help  "Indique até três intervenções ou iniciativas que considera prioritárias"

Add-FormQuestion -Type 'paragraph' -Title '40. Sugestão 1 (obrigatória)' -Required $true
Add-FormQuestion -Type 'paragraph' -Title '40. Sugestão 2 (opcional)'
Add-FormQuestion -Type 'paragraph' -Title '40. Sugestão 3 (opcional)'

Add-FormQuestion -Type 'paragraph' -Title 'Observações adicionais' `
    -Help 'Comentários, esclarecimentos ou informações adicionais que considere relevantes'


# ─────────────────────────────────────────────
# 4. RESULTADO FINAL
# ─────────────────────────────────────────────
Write-Host "`n[4/4] Concluído!" -ForegroundColor Green
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  FORMULÁRIO CRIADO COM SUCESSO                                   ║" -ForegroundColor Cyan
Write-Host "╠══════════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
Write-Host "║                                                                  ║" -ForegroundColor Cyan
Write-Host "║  LINK PARA PARTILHAR COM AS INSTITUIÇÕES:                        ║" -ForegroundColor Cyan
Write-Host "║  $formUrl" -ForegroundColor White
Write-Host "║                                                                  ║" -ForegroundColor Cyan
Write-Host "║  EDITAR O FORMULÁRIO:                                            ║" -ForegroundColor Cyan
Write-Host "║  forms.microsoft.com  →  Os meus formulários                    ║" -ForegroundColor White
Write-Host "║                                                                  ║" -ForegroundColor Cyan
Write-Host "║  EXPORTAR RESPOSTAS PARA EXCEL:                                  ║" -ForegroundColor Cyan
Write-Host "║  Formulário → separador Respostas → botão Abrir no Excel         ║" -ForegroundColor White
Write-Host "║                                                                  ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Copiar URL para a área de transferência (Windows)
try {
    $formUrl | Set-Clipboard
    Write-Host "  O link foi copiado para a área de transferência." -ForegroundColor Green
} catch {}

Disconnect-MgGraph -ErrorAction SilentlyContinue

<#
════════════════════════════════════════════════════════════════════
GUIA DE CRIAÇÃO MANUAL (caso o script não funcione)
════════════════════════════════════════════════════════════════════

Se o script falhar (ex: permissões da organização não permitem
acesso à API), crie o formulário manualmente em forms.microsoft.com:

1. Clique em "+ Novo Formulário"
2. Adicione as perguntas usando os tipos:
   - Escolha  → para radio buttons (escolha única)
   - Escolha  → marque "Múltiplas respostas" para checkboxes
   - Texto    → para resposta curta / números
   - Texto    → marque "Resposta longa" para parágrafos
   - Lista pendente → para dropdowns
   - Avaliação (escala) → para as perguntas 39 (fatores)
3. Use "Adicionar nova secção" para separar as 6 secções
4. Ative "Ramificação" nas perguntas Sim/Não para mostrar
   sub-perguntas condicionalmente

EXPORTAR PARA EXCEL:
  Separador "Respostas" → "Abrir no Excel"
  → Ficheiro .xlsx com 1 linha por instituição

════════════════════════════════════════════════════════════════════
#>
