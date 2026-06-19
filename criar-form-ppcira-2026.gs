/**
 * CIURA/PPCIRA 2026 — Criação automática do Google Forms
 *
 * Como usar:
 *  1. Aceda a script.google.com com a conta Google institucional
 *  2. Clique em "Novo projeto"
 *  3. Apague o código existente e cole TODO este script
 *  4. Clique em "Executar" > selecione "criarFormPPCIRA2026" > confirme permissões
 *  5. Aguarde ~10 segundos — o formulário aparece no seu Google Drive
 *  6. Copie o link de partilha e envie às instituições
 *  7. As respostas são exportáveis para Excel: separador "Respostas" > ícone Google Sheets
 */

function criarFormPPCIRA2026() {

  // ---- CRIAR FORMULÁRIO ----
  var form = FormApp.create('Inquérito Nacional CIURA/PPCIRA 2026');

  form.setTitle('Inquérito Nacional de Caracterização dos S/UL-PPCIRA — CIURA/PPCIRA 2026');

  form.setDescription(
    'Promovido pelo PPCIRA — Programa de Prevenção e Controlo de Infeções e de Resistência aos ' +
    'Antimicrobianos da Direção-Geral da Saúde, em colaboração com o CIURA — Colégio da Competência ' +
    'em Controlo de Infeção e Uso Racional de Antibióticos da Ordem dos Médicos.\n\n' +
    '▸ Destinado a: Diretores/Coordenadores das estruturas locais S/UL-PPCIRA\n' +
    '▸ Uma única resposta por instituição\n' +
    '▸ Tempo estimado: ~30 minutos\n' +
    '▸ Prazo de resposta: 30 de junho de 2026\n' +
    '▸ Respostas confidenciais, analisadas em agregado pelo PPCIRA/DGS e pelo CIURA'
  );

  form.setConfirmationMessage(
    'Obrigado pela sua participação no Inquérito Nacional CIURA/PPCIRA 2026.\n' +
    'A sua resposta foi registada com sucesso e será analisada de forma confidencial.'
  );

  form.setProgressBar(true);
  form.setCollectEmail(false);   // Altere para true se pretender registar o email do respondente
  form.setLimitOneResponsePerUser(false);
  form.setAllowResponseEdits(true);
  form.setPublishingSummary(false);

  // ============================================================
  // SECÇÃO I — CARACTERIZAÇÃO E ENQUADRAMENTO INSTITUCIONAL
  // ============================================================

  form.addSectionHeaderItem()
    .setTitle('SECÇÃO I — Caracterização e Enquadramento Institucional')
    .setHelpText('Identificação e estrutura da instituição respondente');

  // Q1
  form.addTextItem()
    .setTitle('1. Nome da instituição')
    .setRequired(true);

  // Q2
  form.addMultipleChoiceItem()
    .setTitle('2. Setor da instituição')
    .setChoiceValues(['Pública', 'Privada', 'Social', 'Outra'])
    .setRequired(true);

  // Q3
  form.addCheckboxItem()
    .setTitle('3. Tipologia de cuidados (selecione todas as aplicáveis)')
    .setChoiceValues([
      'Hospital',
      'Cuidados de Saúde Primários (CSP)',
      'Cuidados Continuados',
      'Cuidados Paliativos',
      'Outra'
    ])
    .setRequired(true);

  // Q4
  form.addListItem()
    .setTitle('4. Região de Saúde')
    .setChoiceValues([
      'Norte',
      'Centro',
      'Lisboa e Vale do Tejo (LVT)',
      'Alentejo',
      'Algarve',
      'Região Autónoma da Madeira (RAM)',
      'Região Autónoma dos Açores (RAA)'
    ])
    .setRequired(true);

  // Q5 — Dimensão assistencial
  form.addSectionHeaderItem()
    .setTitle('5. Dimensão assistencial (referente a 2025)')
    .setHelpText('Preencha os campos aplicáveis à sua instituição; deixe em branco os que não se aplicam.');

  var numValidation = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .build();

  form.addTextItem()
    .setTitle('5a. N.º de camas de internamento agudo')
    .setValidation(numValidation);

  form.addTextItem()
    .setTitle('5b. N.º de camas de UCI')
    .setValidation(numValidation);

  form.addTextItem()
    .setTitle('5c. N.º de camas de cuidados continuados')
    .setValidation(numValidation);

  form.addTextItem()
    .setTitle('5d. Taxa de ocupação global em 2025 (%)')
    .setValidation(
      FormApp.createTextValidation()
        .requireNumberBetween(0, 100)
        .build()
    );

  form.addTextItem()
    .setTitle('5e. População inscrita nos CSP — n.º de utentes (apenas ULS)')
    .setValidation(numValidation);

  // Q6
  form.addMultipleChoiceItem()
    .setTitle('6. O S/UL-PPCIRA está formalmente nomeado pelo órgão máximo de gestão?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('6.1 Tipo de estrutura formal do S/UL-PPCIRA')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 6')
    .setChoiceValues([
      'Serviço autónomo',
      'Unidade Funcional Autónoma',
      'Unidade Funcional Integrada',
      'Estrutura de Apoio Técnico',
      'Comissão',
      'Outro'
    ]);

  // Q7
  form.addMultipleChoiceItem()
    .setTitle('7. O S/UL-PPCIRA possui Regulamento aprovado?')
    .setChoiceValues(['Sim', 'Em elaboração', 'Não'])
    .setRequired(true);

  // Q8
  form.addListItem()
    .setTitle('8. O S/UL-PPCIRA reporta diretamente a:')
    .setChoiceValues([
      'Conselho de Administração',
      'Direção Clínica',
      'Direção Clínica e Direção de Enfermagem',
      'Direção de Enfermagem',
      'Departamento',
      'Outro'
    ])
    .setRequired(true);

  // Q9
  form.addMultipleChoiceItem()
    .setTitle('9. É contratualizado anualmente um Plano de Atividades do S/UL-PPCIRA?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  // Q10
  form.addMultipleChoiceItem()
    .setTitle('10. Existe relatório anual de atividades do S/UL-PPCIRA?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  // Q11
  form.addMultipleChoiceItem()
    .setTitle('11. Os objetivos estratégicos institucionais contemplam indicadores PCIRA?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('11.1 Especifique os indicadores PCIRA contemplados nos objetivos estratégicos')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 11');

  // Q12
  form.addMultipleChoiceItem()
    .setTitle('12. O S/UL-PPCIRA é centro de custos autónomo?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  // ============================================================
  // SECÇÃO II — RECURSOS HUMANOS E ESTRUTURA DE FUNCIONAMENTO
  // ============================================================

  form.addPageBreakItem()
    .setTitle('SECÇÃO II — Recursos Humanos e Estrutura de Funcionamento')
    .setHelpText('Composição da equipa, nomeação e organização interna');

  // Q13 — Equipa (tabela: 7 categorias × 4 campos)
  form.addSectionHeaderItem()
    .setTitle('13. Composição da equipa do S/UL-PPCIRA')
    .setHelpText(
      'Preencha os dados de cada categoria profissional. ' +
      'Deixe em branco as categorias não aplicáveis à sua estrutura.'
    );

  var equipaCategorias = [
    'Diretor(a)',
    'Coordenador(a)',
    'Enfermeiro(a) Gestor(a)',
    'Médico(s)',
    'Enfermeiro(s)',
    'Farmacêutico(s)',
    'Outro(s)'
  ];

  equipaCategorias.forEach(function(cat) {
    form.addTextItem()
      .setTitle('13. ' + cat + ' — N.º de elementos')
      .setValidation(numValidation);
    form.addTextItem()
      .setTitle('13. ' + cat + ' — Especialidade / Área profissional');
    form.addTextItem()
      .setTitle('13. ' + cat + ' — Horas semanais contratualizadas (h)')
      .setValidation(numValidation);
    form.addTextItem()
      .setTitle('13. ' + cat + ' — Horas semanais efetivas (h)')
      .setValidation(numValidation);
  });

  // Q13.1
  form.addMultipleChoiceItem()
    .setTitle('13.1 O S/UL-PPCIRA tem Diretor ou Coordenador formalmente nomeado?')
    .setChoiceValues(['Sim', 'Não']);

  // Q13.2
  form.addMultipleChoiceItem()
    .setTitle('13.2 A nomeação formal abrange todos os elementos da equipa?')
    .setChoiceValues(['Sim, todos', 'Parcialmente', 'Não']);

  // Q13.3
  form.addMultipleChoiceItem()
    .setTitle('13.3 A equipa cumpre os requisitos mínimos do Despacho n.º 10901/2022?')
    .setChoiceValues(['Sim', 'Não']);

  form.addCheckboxItem()
    .setTitle('13.3.1 Identifique o(s) défice(s) existente(s)')
    .setHelpText('↳ Responda apenas se respondeu "Não" à questão 13.3')
    .setChoiceValues([
      'Défice de tempo médico',
      'Défice de tempo de enfermagem',
      'Ausência de microbiologista',
      'Ausência de farmacêutico',
      'Ausência de horário completo do Diretor',
      'Ausência de competências em qualidade',
      'Outro'
    ]);

  // Q14
  form.addMultipleChoiceItem()
    .setTitle('14. Existem elementos dinamizadores/elos PCIRA nos serviços formalmente designados?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  form.addListItem()
    .setTitle('14.1 Percentagem de serviços com elos PCIRA designados')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 14')
    .setChoiceValues(['0–25%', '26–50%', '51–75%', '76–100%']);

  // Q15 — Distribuição de horas PCI/PAPA
  form.addSectionHeaderItem()
    .setTitle('15. Distribuição percentual das horas semanais entre PCI e PAPA')
    .setHelpText('Indique a percentagem aproximada das horas semanais de cada categoria dedicadas a PCI vs. PAPA');

  var pctValidation = FormApp.createTextValidation()
    .requireNumberBetween(0, 100)
    .build();

  ['Médicos', 'Enfermeiros', 'Outros'].forEach(function(cat) {
    form.addTextItem()
      .setTitle('15. ' + cat + ' — % horas dedicadas a PCI')
      .setValidation(pctValidation);
    form.addTextItem()
      .setTitle('15. ' + cat + ' — % horas dedicadas a PAPA')
      .setValidation(pctValidation);
  });

  // Q16
  form.addListItem()
    .setTitle('16. Periodicidade de reunião da equipa nuclear do S/UL-PPCIRA')
    .setChoiceValues(['Semanal', 'Quinzenal', 'Mensal', 'Trimestral', 'Sem periodicidade definida', 'Outra']);

  // Q17
  form.addListItem()
    .setTitle('17. Periodicidade de reunião com o órgão máximo de gestão')
    .setChoiceValues(['Mensal', 'Trimestral', 'Semestral', 'Anual', 'Só quando solicitado', 'Outra']);

  // Q18
  form.addListItem()
    .setTitle('18. Periodicidade de reunião com os principais serviços clínicos')
    .setChoiceValues(['Mensal', 'Trimestral', 'Semestral', 'Anual', 'Só quando solicitado', 'Outra']);

  // Q19
  form.addMultipleChoiceItem()
    .setTitle('19. É realizada a contratualização de objetivos PPCIRA com os serviços clínicos?')
    .setChoiceValues(['Sim, em todos os serviços', 'Sim, em alguns serviços', 'Não']);

  form.addListItem()
    .setTitle('19.1 Percentagem de serviços com objetivos PPCIRA contratualizados')
    .setHelpText('↳ Responda apenas se respondeu "Sim, em alguns serviços" à questão 19')
    .setChoiceValues(['0–25%', '26–50%', '51–75%', '76–100%']);

  // ============================================================
  // SECÇÃO III — ATIVIDADES DE VIGILÂNCIA EPIDEMIOLÓGICA
  // ============================================================

  form.addPageBreakItem()
    .setTitle('SECÇÃO III — Atividades de Vigilância Epidemiológica')
    .setHelpText('Programas de VE, PAPA e participação em projetos nacionais/internacionais');

  // Q20
  form.addMultipleChoiceItem()
    .setTitle('20. Existem programas estruturados de vigilância epidemiológica (VE) nas diferentes dimensões PCIRA?')
    .setChoiceValues(['Sim, em todas as dimensões', 'Parcialmente', 'Não'])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('20.1 Dimensões abrangidas pelos programas de VE (selecione todas as aplicáveis)')
    .setHelpText('↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20')
    .setChoiceValues([
      'Precauções Básicas de Controlo de Infeção (PBCI)',
      'Uso de luvas',
      'Higiene das mãos',
      'Infeção de Local Cirúrgico (ILC)',
      'ILC — cirurgia específica (ECDC)',
      'Infeção da corrente sanguínea relacionada com CVC (ICSRCVC)',
      'Bacteriémias',
      'Pneumonia associada a intubação (PAI)',
      'Infeção do trato urinário associada a cateter vesical (ITUACV)',
      'Microrganismos multirresistentes (MMR)',
      'Clostridioides difficile',
      'Consumo de antimicrobianos',
      'Qualidade da prescrição (PAPA)',
      'Outros'
    ]);

  form.addListItem()
    .setTitle('20.2 Percentagem de serviços abrangidos pela VE')
    .setHelpText('↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20')
    .setChoiceValues(['0–25%', '26–50%', '51–75%', '76–100%']);

  form.addCheckboxItem()
    .setTitle('20.3 Sistema de informação utilizado para a VE')
    .setHelpText('↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 20')
    .setChoiceValues(['Plataforma DGS', 'Hepic®', 'Meliora®', 'Outro']);

  form.addMultipleChoiceItem()
    .setTitle('20.4 Existe feedback estruturado e regular dos dados de VE para os serviços?')
    .setChoiceValues(['Sim', 'Não']);

  form.addParagraphTextItem()
    .setTitle('20.4.1 Frequência e método do feedback dos dados de VE')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 20.4 — descreva a frequência e o método utilizado');

  // Q21
  form.addMultipleChoiceItem()
    .setTitle('21. O Programa de Apoio à Prescrição de Antibióticos (PAPA) está em funcionamento?')
    .setChoiceValues(['Sim, em todos os serviços', 'Parcialmente', 'Não'])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('21.1 Em que valências/contextos está implementado o PAPA?')
    .setHelpText('↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21')
    .setChoiceValues([
      'Cuidados de Saúde Primários (CSP)',
      'Internamento hospitalar',
      'Ambulatório / Hospital de dia / Consultas externas',
      'Unidades de Cuidados Continuados Integrados (UCCI)',
      'Serviço de Medicina Intensiva (SMI)',
      'Urgência'
    ]);

  form.addCheckboxItem()
    .setTitle('21.2 Componentes do PAPA implementadas (selecione todas as aplicáveis)')
    .setHelpText('↳ Responda apenas se respondeu "Sim" ou "Parcialmente" à questão 21')
    .setChoiceValues([
      'Revisão/restrição de prescrições até 72 horas',
      'Protocolos institucionais de utilização de antibióticos',
      'Formação contínua de profissionais de saúde sobre antibióticos',
      'Retorno de métricas por unidade/serviço/prescritor',
      'Projetos de intervenção comportamental (nudge, etc.)',
      'Outros'
    ]);

  // Q22
  form.addListItem()
    .setTitle('22. Com que frequência é realizado o feedback dos dados PAPA aos prescritores?')
    .setChoiceValues(['Mensal', 'Trimestral', 'Semestral', 'Anual', 'Não é realizado']);

  // Q23
  form.addMultipleChoiceItem()
    .setTitle('23. Existe plano institucional de avaliação e resposta a surtos?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  // Q24
  form.addMultipleChoiceItem()
    .setTitle('24. Participa ativamente em algum projeto de melhoria da qualidade a nível nacional (ex: STOP-Infeção, Drive AMS, ITUCCI)?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('24.1 Identifique os projetos nacionais de melhoria da qualidade em que participa')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 24');

  // Q25
  form.addMultipleChoiceItem()
    .setTitle('25. Participa ativamente noutros grupos de trabalho ou projetos nacionais/internacionais de PCIRA?')
    .setChoiceValues(['Sim', 'Não']);

  form.addParagraphTextItem()
    .setTitle('25.1 Identifique os grupos de trabalho/projetos nacionais ou internacionais')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 25');

  // ============================================================
  // SECÇÃO IV — FORMAÇÃO, COMPETÊNCIAS E DESENVOLVIMENTO PROFISSIONAL
  // ============================================================

  form.addPageBreakItem()
    .setTitle('SECÇÃO IV — Formação, Competências e Desenvolvimento Profissional')
    .setHelpText('Capacitação da equipa e certificações profissionais em PCIRA');

  // Q26
  form.addMultipleChoiceItem()
    .setTitle('26. Existem programas formais de capacitação e treino em PCIRA (acolhimento + formação contínua)?')
    .setChoiceValues(['Sim, com carácter obrigatório', 'Sim, com carácter opcional', 'Não'])
    .setRequired(true);

  form.addListItem()
    .setTitle('26.1 Periodicidade da formação PCIRA para os profissionais da instituição')
    .setHelpText('↳ Responda apenas se respondeu "Sim" (obrigatório ou opcional) à questão 26')
    .setChoiceValues(['Anual', 'De 2 em 2 anos', 'De 3 em 3 anos', 'Sem periodicidade definida']);

  // Q27
  form.addMultipleChoiceItem()
    .setTitle('27. A formação contínua e progressiva da equipa S/UL-PPCIRA está assegurada?')
    .setChoiceValues(['Sim', 'Não']);

  form.addParagraphTextItem()
    .setTitle('27.1 Como é assegurada a formação contínua da equipa S/UL-PPCIRA?')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 27 — descreva modalidades e recursos utilizados');

  // Q28 — Formação diferenciada
  form.addSectionHeaderItem()
    .setTitle('28. Formação diferenciada em PPCIRA — n.º de profissionais da equipa S/UL-PPCIRA')
    .setHelpText('Indique o número de profissionais com cada nível de formação, por categoria profissional. Coloque 0 se não aplicável.');

  var formacaoNiveis = [
    'Doutoramento em PCIRA',
    'Mestrado em PCIRA',
    'Pós-Graduação em PCIRA',
    'Certificação europeia PCI ou PAPA (EUCIC / ESCMID)',
    'Outros cursos de formação diferenciada'
  ];

  formacaoNiveis.forEach(function(nivel) {
    ['Médicos', 'Enfermeiros', 'Outros profissionais'].forEach(function(cat) {
      form.addTextItem()
        .setTitle('28. ' + nivel + ' — ' + cat + ' (n.º)')
        .setValidation(numValidation);
    });
  });

  // Q29
  form.addMultipleChoiceItem()
    .setTitle('29. Existem profissionais com Competência atribuída pela respetiva Ordem Profissional em PCI e/ou PAPA?')
    .setChoiceValues(['Sim', 'Não']);

  form.addTextItem()
    .setTitle('29.1a N.º de médicos com CCIURA (Ordem dos Médicos)')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 29')
    .setValidation(numValidation);

  form.addTextItem()
    .setTitle('29.1b N.º de enfermeiros com CADEPCI (Ordem dos Enfermeiros)')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 29')
    .setValidation(numValidation);

  // Q30
  form.addMultipleChoiceItem()
    .setTitle('30. O plano de desenvolvimento profissional da equipa inclui investigação e publicações em PCIRA?')
    .setChoiceValues(['Sim', 'Não']);

  // ============================================================
  // SECÇÃO V — ENVOLVIMENTO INSTITUCIONAL E INFRAESTRUTURAS
  // ============================================================

  form.addPageBreakItem()
    .setTitle('SECÇÃO V — Envolvimento Institucional e Infraestruturas')
    .setHelpText('Espaço físico, apoio administrativo, integração em comissões e suporte laboratorial');

  // Q31
  form.addMultipleChoiceItem()
    .setTitle('31. O S/UL-PPCIRA dispõe de espaço físico próprio (gabinete dedicado)?')
    .setChoiceValues(['Sim', 'Não'])
    .setRequired(true);

  // Q32
  form.addMultipleChoiceItem()
    .setTitle('32. O S/UL-PPCIRA dispõe de apoio administrativo alocado?')
    .setChoiceValues(['Sim, exclusivo', 'Sim, partilhado', 'Não'])
    .setRequired(true);

  // Q33
  form.addCheckboxItem()
    .setTitle('33. O S/UL-PPCIRA integra formalmente as seguintes comissões/grupos de trabalho institucionais:')
    .setChoiceValues([
      'Comissão de Farmácia e Terapêutica (CFT)',
      'Comissão/Unidade de Qualidade e Segurança do Doente',
      'Outras comissões ou grupos de trabalho institucionais',
      'Não integra nenhuma comissão formalmente'
    ]);

  // Q34
  form.addMultipleChoiceItem()
    .setTitle('34. O S/UL-PPCIRA dispõe de laboratório de microbiologia interno?')
    .setChoiceValues([
      'Sim, sem descontinuidade de serviço',
      'Sim, com interrupção ou cobertura parcial',
      'Não (recorre a laboratório externo)'
    ])
    .setRequired(true);

  // Q35
  form.addMultipleChoiceItem()
    .setTitle('35. Os CSP trabalham com que tipo de laboratório de microbiologia?')
    .setHelpText('Responda apenas se a sua instituição inclui CSP')
    .setChoiceValues(['Laboratório interno', 'Laboratório externo', 'Não aplicável']);

  // Q36
  form.addMultipleChoiceItem()
    .setTitle('36. Os CSP utilizam testes de diagnóstico rápido microbiológico (Point-of-Care)?')
    .setHelpText('Responda apenas se a sua instituição inclui CSP')
    .setChoiceValues(['Sim', 'Não', 'Não aplicável']);

  form.addParagraphTextItem()
    .setTitle('36.1 Especifique os testes Point-of-Care microbiológicos utilizados nos CSP')
    .setHelpText('↳ Responda apenas se respondeu "Sim" à questão 36');

  // Q37
  form.addListItem()
    .setTitle('37a. Carta microbiológica hospitalar — periodicidade de atualização')
    .setChoiceValues(['Semestral', 'Anual', 'De 2 em 2 anos', 'Não existe', 'Não aplicável']);

  form.addListItem()
    .setTitle('37b. Carta microbiológica comunitária (CSP) — periodicidade de atualização')
    .setChoiceValues(['Semestral', 'Anual', 'De 2 em 2 anos', 'Não existe', 'Não aplicável']);

  // ============================================================
  // SECÇÃO VI — PERCEÇÃO PRÁTICA
  // ============================================================

  form.addPageBreakItem()
    .setTitle('SECÇÃO VI — Perceção Prática')
    .setHelpText('Avaliação global das condições institucionais e sugestões para o PPCIRA/DGS');

  // Q38
  form.addMultipleChoiceItem()
    .setTitle('38. Qual a classificação global das condições institucionais para a implementação do PPCIRA?')
    .setChoiceValues(['Elevadas', 'Moderadas', 'Baixas'])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('38.1 Justificação da classificação atribuída')
    .setHelpText('Descreva os principais motivos para a classificação escolhida');

  // Q39 — Fatores limitantes (escala de criticidade 1–5 por fator)
  form.addSectionHeaderItem()
    .setTitle('39. Fatores limitantes — grau de criticidade para a sua instituição')
    .setHelpText(
      'Para cada fator abaixo, classifique o grau de criticidade na sua instituição:\n' +
      '1 = Não é um fator limitante   |   5 = Fator muito crítico / limitante'
    );

  var fatores = [
    'Falta de priorização pelo Conselho de Administração',
    'Falta de priorização pelos Responsáveis de Serviço',
    'Falta de recursos humanos com tempo protegido para PCIRA',
    'Falta de sistemas de informação adequados',
    'Falta de formação contínua da equipa S/UL-PPCIRA',
    'Falta de formação dos demais profissionais de saúde',
    'Resistência à mudança / baixa adesão clínica',
    'Inexistência de infraestrutura laboratorial com resposta adequada'
  ];

  fatores.forEach(function(fator) {
    form.addScaleItem()
      .setTitle('39. ' + fator)
      .setBounds(1, 5)
      .setLabels('Não limitante', 'Muito crítico');
  });

  form.addTextItem()
    .setTitle('39. Outro fator limitante — especificação')
    .setHelpText('Se identificou outro fator limitante relevante, descreva-o aqui');

  // Q40 — Sugestões
  form.addSectionHeaderItem()
    .setTitle('40. Sugestões para o PPCIRA/DGS — próximos 3 anos')
    .setHelpText('Indique até três intervenções ou iniciativas que considera prioritárias para o PPCIRA/DGS implementar nos próximos 3 anos');

  form.addParagraphTextItem()
    .setTitle('40. Sugestão 1 (obrigatória)')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('40. Sugestão 2 (opcional)');

  form.addParagraphTextItem()
    .setTitle('40. Sugestão 3 (opcional)');

  // Observações gerais
  form.addParagraphTextItem()
    .setTitle('Observações adicionais')
    .setHelpText(
      'Utilize este espaço para quaisquer comentários, esclarecimentos ou informações adicionais ' +
      'que considere relevantes para o PPCIRA/DGS'
    );

  // ============================================================
  // SAÍDA — URLs e instruções
  // ============================================================

  var urlResposta = form.getPublishedUrl();
  var urlEdicao   = form.getEditUrl();
  var formId      = form.getId();

  Logger.log('=== FORMULÁRIO CRIADO COM SUCESSO ===');
  Logger.log('');
  Logger.log('LINK PARA PARTILHAR COM AS INSTITUIÇÕES (para resposta):');
  Logger.log(urlResposta);
  Logger.log('');
  Logger.log('LINK DE EDIÇÃO (apenas para o gestor do formulário):');
  Logger.log(urlEdicao);
  Logger.log('');
  Logger.log('ID DO FORMULÁRIO: ' + formId);
  Logger.log('');
  Logger.log('COMO EXPORTAR RESPOSTAS PARA EXCEL:');
  Logger.log('  1. Abra o formulário > separador "Respostas"');
  Logger.log('  2. Clique no ícone do Google Sheets (folha verde)');
  Logger.log('  3. Cria automaticamente uma Folha de Cálculo com 1 linha por resposta');
  Logger.log('  4. No Google Sheets: Ficheiro > Transferir > Microsoft Excel (.xlsx)');

  // Mostrar URLs no ecrã (se executado a partir de uma folha de cálculo ligada)
  try {
    Browser.msgBox(
      'Formulário criado com sucesso!',
      'Link para partilhar:\\n' + urlResposta + '\\n\\nLink de edição:\\n' + urlEdicao,
      Browser.Buttons.OK
    );
  } catch(e) {
    // Ignorar se não houver UI disponível (execução standalone)
  }

  return {
    urlResposta: urlResposta,
    urlEdicao:   urlEdicao,
    formId:      formId
  };
}
