
import * as XLSX from 'xlsx';

const generateBtn = document.getElementById('generate-btn');
const readBtn = document.getElementById('read-btn');
const fileInput = document.getElementById('file-input');
const cpfInput = document.getElementById('cpf');
const sendEmailBtn = document.getElementById('send-email-btn');

// --- Declaração de todas as constantes de elementos HTML ---

// Header e Dados Pessoais do Titular (Seção "CADASTRO DE SÓCIO PARTICIPANTE")
const proposalNumberUniao = document.getElementById('proposal_number_uniao');
const orgaoInput = document.getElementById('orgao');
const proponenteInput = document.getElementById('proponente');
const nomeInput = document.getElementById('nome');
const nascInput = document.getElementById('nasc');
const sexoInput = document.getElementById('sexo');
const estCivilInput = document.getElementById('est_civil');
const rgInput = document.getElementById('rg');
const expInput = document.getElementById('exp');
const maeInput = document.getElementById('mae');
const emailInput = document.getElementById('email');
const bancoInput = document.getElementById('banco');
const agenciaInput = document.getElementById('agencia');
const contaCorrenteInput = document.getElementById('conta_corrente');
const conjugeInput = document.getElementById('conjuge');
const nascConjugeInput = document.getElementById('nasc_conjuge');
const sexoConjugeInput = document.getElementById('sexo_conjuge');
const endInput = document.getElementById('end');
const numInput = document.getElementById('num');
const complInput = document.getElementById('compl');
const bairroInput = document.getElementById('bairro');
const cepInput = document.getElementById('cep');
const cidadeInput = document.getElementById('cidade');
const estInput = document.getElementById('est');
const telInput = document.getElementById('tel');
const celularInput = document.getElementById('celular');
const orgaoFuncionalInput = document.getElementById('orgao_funcional');
const matFuncionalInput = document.getElementById('mat_funcional');
const funcaoInput = document.getElementById('funcao');
const unidadeInput = document.getElementById('unidade');

// Benefícios Contratados
const MensalidadeSocial = document.getElementById('mensalidade_social');
const planoSaudeAmb = document.getElementById('plano_saude_amb');
const PlanoSaudeComp = document.getElementById('plano_saude_comp');
const OrientacaoJuridica = document.getElementById('orientacao_juridica');
const PlanoOdonto = document.getElementById('plano_odonto');
const SeguroVida = document.getElementById('seguro_vida');
const SeguroVidaSim = document.querySelector('input[name="seguro_vida_sim"]');
const SeguroVidaNao = document.querySelector('input[name="seguro_vida_nao"]');
const AuxilioNatalidade = document.getElementById('auxilio_natalidade');
const AssistenciaFuneral = document.getElementById('assistencia_funeral');
const Convenios = document.getElementById('convenios');

// Relação de Dependentes (União de Benefícios) - dep1 a dep8
// Não precisamos declarar todas as constantes aqui individualmente se vamos iterar sobre elas.
// Mas as referências para os loops serão feitas diretamente no loop.

// Termos e Declaração
const declarationRsValue = document.getElementById('declaration_rs_value');
const declarationLocal = document.getElementById('declaration_local');
const declarationDia = document.getElementById('declaration_dia');
const declarationMes = document.getElementById('declaration_mes');
const declarationAno = document.getElementById('declaration_ano');

// Declaração do Agenciador/Consultor
const agentName = document.getElementById('agent_name');
const agentRegistro = document.getElementById('agent_registro');

// Proposta de Contratação - Dados do Vendedor
const contratacaoProposalNumber = document.getElementById('contratacao_proposal_number');
const contratacaoVendedorCpf = document.getElementById('contratacao_vendedor_cpf');
const contratacaoVendedorNome = document.getElementById('contratacao_vendedor_nome');
const contratacaoVendedorTelefone = document.getElementById('contratacao_vendedor_telefone');
const contratacaoCnpjCorretora = document.getElementById('contratacao_cnpj_corretora');
const contratacaoNomeCorretora = document.getElementById('contratacao_nome_corretora');

// Proposta de Contratação - Dados cadastrais do Contratante
const contratacaoContratanteResponsavel = document.getElementById('contratacao_contratante_responsavel');
const contratacaoContratanteCpf = document.getElementById('contratacao_contratante_cpf');
const contratacaoContratanteNome = document.getElementById('contratacao_contratante_nome');
const contratacaoContratanteSexo = document.getElementById('contratacao_contratante_sexo');
const contratacaoContratanteRg = document.getElementById('contratacao_contratante_rg');
const contratacaoContratanteDataEmissaoRg = document.getElementById('contratacao_contratante_data_emissao_rg');
const contratacaoContratanteOrgaoEmissor = document.getElementById('contratacao_contratante_orgao_emissor');
const contratacaoContratanteEstadoCivil = document.getElementById('contratacao_contratante_estado_civil');
const contratacaoContratanteDataNascimento = document.getElementById('contratacao_contratante_data_nascimento');
const contratacaoContratanteTelefoneFixo = document.getElementById('contratacao_contratante_telefone_fixo');
const contratacaoContratanteCelular = document.getElementById('contratacao_contratante_celular');
const contratacaoContratanteWhatsapp = document.getElementById('contratacao_contratante_whatsapp');
const contratacaoContratanteCartaoNacionalSaude = document.getElementById('contratacao_contratante_cartao_nacional_saude');
const contratacaoContratanteEmail = document.getElementById('contratacao_contratante_email');
const contratacaoContratanteNomeMae = document.getElementById('contratacao_contratante_nome_mae');
const contratacaoContratantePlanoAnteriorSim = document.getElementById('contratacao_contratante_plano_anterior_sim');
const contratacaoContratantePlanoAnteriorNao = document.getElementById('contratacao_contratante_plano_anterior_nao');
const contratacaoContratanteQualPlano = document.getElementById('contratacao_contratante_qual_plano');
const contratacaoContratanteDataInicio = document.getElementById('contratacao_contratante_data_inicio');
const contratacaoContratanteDataUltimoPagamento = document.getElementById('contratacao_contratante_data_ultimo_pagamento');
const contratacaoContratanteCep = document.getElementById('contratacao_contratante_cep');
const contratacaoContratanteLogradouro = document.getElementById('contratacao_contratante_logradouro');
const contratacaoContratanteNumero = document.getElementById('contratacao_contratante_numero');
const contratacaoContratanteComplemento = document.getElementById('contratacao_contratante_complemento');
const contratacaoContratanteBairro = document.getElementById('contratacao_contratante_bairro');
const contratacaoContratanteCidade = document.getElementById('contratacao_contratante_cidade');
const contratacaoContratanteEstado = document.getElementById('contratacao_contratante_estado');

// Documentação entregue (Contratante)
const docEntregueCpfRgCnh = document.getElementById('doc_entregue_cpf_rg_cnh');
const docEntregueFotoseIfie = document.getElementById('doc_entregue_fotoseIfie');
const docEntregueComprovanteResidencia = document.getElementById('doc_entregue_comprovante_residencia');
const docEntregueComprovanteVendaBeneficiario = document.getElementById('doc_entregue_comprovante_venda_beneficiario');
const docEntregueOutros = document.getElementById('doc_entregue_outros');
const docEntregueDeclaracaoTempoPermanencia = document.getElementById('doc_entregue_declaracao_tempo_permanencia');

// Proposta de Contratação - Dependentes (contratacao_depX)
// Serão acessados via loop, não precisam de constantes individuais aqui.

// Responsável pelo Contrato
const contratacaoProposalNumberResponsavel = document.getElementById('contratacao_proposal_number_responsavel');
const contratacaoResponsavelCpf = document.getElementById('contratacao_responsavel_cpf');
const contratacaoResponsavelNome = document.getElementById('contratacao_responsavel_nome');
const contratacaoResponsavelSexoMasculino = document.getElementById('contratacao_responsavel_sexo_masculino');
const contratacaoResponsavelSexoFeminino = document.getElementById('contratacao_responsavel_sexo_feminino');
const contratacaoResponsavelEmail = document.getElementById('contratacao_responsavel_email');
const contratacaoResponsavelNomeMae = document.getElementById('contratacao_responsavel_nome_mae');
const contratacaoResponsavelCelular = document.getElementById('contratacao_responsavel_celular');
const contratacaoResponsavelDataNascimento = document.getElementById('contratacao_responsavel_data_nascimento');
const contratacaoResponsavelEstadoCivil = document.getElementById('contratacao_responsavel_estado_civil');
const contratacaoResponsavelRgCnh = document.getElementById('contratacao_responsavel_rg_cnh');
const contratacaoResponsavelDataEmissaoRgCnh = document.getElementById('contratacao_responsavel_data_emissao_rg_cnh');
const contratacaoResponsavelCep = document.getElementById('contratacao_responsavel_cep');
const contratacaoResponsavelLogradouro = document.getElementById('contratacao_responsavel_logradouro');
const contratacaoResponsavelNumero = document.getElementById('contratacao_responsavel_numero');
const contratacaoResponsavelComplemento = document.getElementById('contratacao_responsavel_complemento');
const contratacaoResponsavelBairro = document.getElementById('contratacao_responsavel_bairro');
const contratacaoResponsavelCidade = document.getElementById('contratacao_responsavel_cidade');
const contratacaoResponsavelEstado = document.getElementById('contratacao_responsavel_estado');

// Documentação entregue (Responsável)
const docEntregueRespCpfRgCnh = document.getElementById('doc_entregue_resp_cpf_rg_cnh');
const docEntregueRespComprovanteResidencia = document.getElementById('doc_entregue_resp_comprovante_residencia');
const docEntregueRespOutros = document.getElementById('doc_entregue_resp_outros');

// Resumo da Contratação
const contratacaoResumoDataProposta = document.getElementById('contratacao_resumo_data_proposta');
const contratacaoResumoProvavelVigencia = document.getElementById('contratacao_resumo_provavel_vigencia');
const contratacaoResumoCpfContratante = document.getElementById('contratacao_resumo_cpf_contratante');
const contratacaoResumoTipo = document.getElementById('contratacao_resumo_tipo');
const contratacaoResumoSegmentacao = document.getElementById('contratacao_resumo_segmentacao');
const contratacaoResumoPlano = document.getElementById('contratacao_resumo_plano');
const contratacaoResumoRegistroAns = document.getElementById('contratacao_resumo_registro_ans');
const acomodacaoQc = document.getElementById('acomodacao_qc');
const acomodacaoQp = document.getElementById('acomodacao_qp');
const contratacaoResumoAbrangencia = document.getElementById('contratacao_resumo_abrangencia');
const coparticipacaoNao = document.getElementById('coparticipacao_nao');
const contratacaoResumoBeneficiarios = document.getElementById('contratacao_resumo_beneficiarios');
const contratacaoResumoValorTotal = document.getElementById('contratacao_resumo_valor_total');


// Declaração de Saúde - Itens 1-21
const declaracaoSaudeProposalNumber = document.getElementById('declaracao_saude_proposal_number');
const declaracaoSaudeProposalNumberCont = document.getElementById('declaracao_saude_proposal_number_cont');
// Itens da declaração de saúde serão acessados via loop.

// Declaração de Saúde - IMC
const bmiTitularPeso = document.getElementById('bmi_titular_peso');
const bmiDep1Peso = document.getElementById('bmi_dep1_peso');
const bmiDep2Peso = document.getElementById('bmi_dep2_peso');
const bmiDep3Peso = document.getElementById('bmi_dep3_peso');
const bmiDep4Peso = document.getElementById('bmi_dep4_peso');
const bmiTitularAltura = document.getElementById('bmi_titular_altura');
const bmiDep1Altura = document.getElementById('bmi_dep1_altura');
const bmiDep2Altura = document.getElementById('bmi_dep2_altura');
const bmiDep3Altura = document.getElementById('bmi_dep3_altura');
const bmiDep4Altura = document.getElementById('bmi_dep4_altura');

// Informações Complementares
const complementaryInfoProposalNumber = document.getElementById('complementary_info_proposal_number');
// Campos de informações complementares serão acessados via loop.

// Declaração de Saúde Final
const declaracaoSaudeFinalProposalNumber = document.getElementById('declaracao_saude_final_proposal_number');
const entrevistaQualificadaOpcao1Chk = document.getElementById('entrevista_qualificada_opcao1_chk');
const entrevistaQualificadaOpcao2Chk = document.getElementById('entrevista_qualificada_opcao2_chk');
const entrevistaQualificadaOpcao3Chk = document.getElementById('entrevista_qualificada_opcao3_chk');

// Carta de Orientação ao Beneficiário - Assinaturas
const cartaBeneficiarioNomeSig = document.getElementById('carta_beneficiario_nome_sig');
const cartaBeneficiarioLocalSig = document.getElementById('carta_beneficiario_local_sig');
const cartaBeneficiarioDataSig = document.getElementById('carta_beneficiario_data_sig');
const cartaIntermediarioNomeSig = document.getElementById('carta_intermediario_nome_sig');
const cartaIntermediarioCpfSig = document.getElementById('carta_intermediario_cpf_sig');
const cartaIntermediarioLocalSig = document.getElementById('carta_intermediario_local_sig');
const cartaIntermediarioDataSig = document.getElementById('carta_intermediario_data_sig');

// Termo Único de Promoções
const termoPromocoesProposalNumber = document.getElementById('termo_promocoes_proposal_number');

// Declaração de Recebimento e Posse
const declaracaoRecebimentoProposalNumber = document.getElementById('declaracao_recebimento_proposal_number');

// Termo de Consentimento
const termoConsentimentoProposalNumber = document.getElementById('termo_consentimento_proposal_number');

// Termo de Adesão Planos Odontológicos - Titular
const odontoTitularNome = document.getElementById('odonto_titular_nome');
const odontoTitularCpf = document.getElementById('odonto_titular_cpf');
const odontoTitularDataNascimento = document.getElementById('odonto_titular_data_nascimento');
const odontoTitularRg = document.getElementById('odonto_titular_rg');
const odontoTitularSexo = document.getElementById('odonto_titular_sexo');
const odontoTitularEstadoCivil = document.getElementById('odonto_titular_estado_civil');
const odontoTitularEndereco = document.getElementById('odonto_titular_endereco');
const odontoTitularNumero = document.getElementById('odonto_titular_numero');
const odontoTitularBairro = document.getElementById('odonto_titular_bairro');
const odontoTitularCep = document.getElementById('odonto_titular_cep');
const odontoTitularCidade = document.getElementById('odonto_titular_cidade');
const odontoTitularTelefone = document.getElementById('odonto_titular_telefone');
const odontoTitularMae = document.getElementById('odonto_titular_mae');
const odontoTitularEmail = document.getElementById('odonto_titular_email');

// Termo de Adesão Planos Odontológicos - Dependentes 1-4
// Serão acessados via loop.

// Autorização para Desconto
const nomeDescontoAutoriza = document.getElementById('nomeDescontoAutoriza');
const cpfDescontoAutoriza = document.getElementById('cpfDescontoAutoriza');
const matFuncionalDescontoAutoriza = document.getElementById('matFuncionalDescontoAutoriza');
const identidadeDescontoAutoriza = document.getElementById('identidadeDescontoAutoriza');
const diaDescontoAutoriza = document.getElementById('diaDescontoAutoriza');
const mesDescontoAutoriza = document.getElementById('mesDescontoAutoriza');
const anosDescontoAutoriza = document.getElementById('anosDescontoAutoriza');
const valorDescontoAutoriza = document.getElementById('valorDescontoAutoriza');
const totalDescontoAutoriza = document.getElementById('totalDescontoAutoriza');
const mensalidadeSocialDescontoAutoriza = document.getElementById('mensalidadeSocialDescontoAutoriza');
const planoSaudeAmbulatorialDescontoAutoriza = document.getElementById('planoSaudeAmbulatorialDescontoAutoriza');
const orientacaoJuridicaDescontoAutoriza = document.getElementById('orientacaoJuridicaDescontoAutoriza');
const segurosDescontoAutoriza = document.getElementById('segurosDescontoAutoriza');
const auxilioNatalidadeDescontoAutoriza = document.getElementById('auxilioNatalidadeDescontoAutoriza');
const planoSaudeCompletoDescontoAutoriza = document.getElementById('planoSaudeCompletoDescontoAutoriza');
const atdDomiciliarDescontoAutoriza = document.getElementById('atdDomiciliarDescontoAutoriza');
const assistenciaFuneralDescontoAutoriza = document.getElementById('assistenciaFuneralDescontoAutoriza');
const planoOdontologicoDescontoAutoriza = document.getElementById('planoOdontologicoDescontoAutoriza');
const conveniosDescontoAutoriza = document.getElementById('conveniosDescontoAutoriza');
const bancoDescontoAutoriza = document.getElementById('bancoDescontoAutoriza');
const agenciaDescontoAutoriza = document.getElementById('agenciaDescontoAutoriza');
const contaCorrenteDescontoAutoriza = document.getElementById('contaCorrenteDescontoAutoriza');
const valorCobrancaDescontoAutoriza = document.getElementById('valorCobrancaDescontoAutoriza');
const despesaBancoDescontoAutoriza = document.getElementById('despesaBancoDescontoAutoriza');
const valorTotalDescontoAutoriza = document.getElementById('valorTotalDescontoAutoriza');
const dataAssDescontoAutoriza = document.getElementById('dataAssDescontoAutoriza');
const assinaturaDescontoAutoriza = document.getElementById('assinaturaDescontoAutoriza');


// --- FUNÇÕES AUXILIARES ---

// Helper function to check if a value should enable a checkbox
const shouldBeChecked = (value) => {
    return value === true || String(value).toUpperCase().trim() === 'SIM' || (typeof value === 'string' && value.trim() !== '' && value.toUpperCase().trim() !== 'NÃO');
};

// --- Funções de Atualização (Populating HTML from other HTML fields) ---

/**
 * Coleta os dados de um dependente de uma seção específica.
 * @param {number} depIndex - O índice do dependente (1 a 8).
 * @param {string} prefix - O prefixo para os IDs/nomes dos campos (ex: 'dep', 'contratacao_dep').
 * @returns {object} Um objeto com os dados do dependente.
 */
function getDependentDataFromSection(depIndex, prefix) {
    const depData = {};
    const nameInput = document.getElementById(`${prefix}${depIndex}_nome`);
    if (!nameInput || !nameInput.value.trim()) {
        return null; // Retorna null se o nome do dependente não estiver preenchido
    }

    depData.nome = nameInput.value.trim();
    depData.nasc = document.getElementById(`${prefix}${depIndex}_nasc`)?.value || '';
    depData.parentesco = document.getElementById(`${prefix}${depIndex}_parentesco`)?.value || '';

    // Campos específicos para 'contratacao_depX'
    if (prefix === 'contratacao_dep') {
        depData.cpf = document.getElementById(`${prefix}${depIndex}_cpf`)?.value || '';
        depData.sexo = document.getElementById(`${prefix}${depIndex}_sexo`)?.value || '';
        depData.rg = document.getElementById(`${prefix}${depIndex}_rg`)?.value || '';
        depData.dataEmissaoRg = document.getElementById(`${prefix}${depIndex}_data_emissao_rg`)?.value || '';
        depData.orgaoEmissor = document.getElementById(`${prefix}${depIndex}_orgao_emissor`)?.value || '';
        depData.estadoCivil = document.getElementById(`${prefix}${depIndex}_estado_civil`)?.value || '';
        depData.cartaoNacionalSaude = document.getElementById(`${prefix}${depIndex}_cartao_nacional_saude`)?.value || '';
        depData.nomeMae = document.getElementById(`${prefix}${depIndex}_nome_mae`)?.value || '';
        // Radio buttons para plano anterior
        const planoAnteriorSim = document.getElementById(`${prefix}${depIndex}_plano_anterior_sim`);
        const planoAnteriorNao = document.getElementById(`${prefix}${depIndex}_plano_anterior_nao`);
        if (planoAnteriorSim?.checked) {
            depData.planoAnterior = planoAnteriorSim.value;
        } else if (planoAnteriorNao?.checked) {
            depData.planoAnterior = planoAnteriorNao.value;
        } else {
            depData.planoAnterior = '';
        }
        depData.qualPlano = document.getElementById(`${prefix}${depIndex}_qual_plano`)?.value || '';
        depData.dataInicio = document.getElementById(`${prefix}${depIndex}_data_inicio`)?.value || '';
        depData.dataUltimoPagamento = document.getElementById(`${prefix}${depIndex}_data_ultimo_pagamento`)?.value || '';
        depData.qualOperadoraAnterior = document.getElementById(`${prefix}${depIndex}_qual_operadora_anterior`)?.value || '';
        depData.registroProdutoAnterior = document.getElementById(`${prefix}${depIndex}_registro_produto_anterior`)?.value || '';

        // Checkboxes de documentação (específicos do Dep1 no HTML, mas aqui generalizado)
        depData.docCpfRgCnh = document.getElementById(`doc_entregue_${prefix}${depIndex}_cpf_rg_cnh`)?.checked || false;
        if (depIndex === 1) { // Apenas para o Dep1, pois o HTML só tem esses IDs para ele
            depData.docFotoseIfie = document.getElementById(`doc_entregue_${prefix}${depIndex}_fotoseIfie`)?.checked || false;
            depData.docComprovanteResidencia = document.getElementById(`doc_entregue_${prefix}${depIndex}_comprovante_residencia`)?.checked || false;
            depData.docComprovanteVendaBeneficiario = document.getElementById(`doc_entregue_${prefix}${depIndex}_comprovante_venda_beneficiario`)?.checked || false;
            depData.docOutros = document.getElementById(`doc_entregue_${prefix}${depIndex}_outros`)?.checked || false;
            depData.docDeclaracaoTempoPermanencia = document.getElementById(`doc_entregue_${prefix}${depIndex}_declaracao_tempo_permanencia`)?.checked || false;
        } else { // Para os outros, garantir que as chaves existam mesmo que os elementos não
            depData.docFotoseIfie = false;
            depData.docComprovanteResidencia = false;
            depData.docComprovanteVendaBeneficiario = false;
            depData.docOutros = false;
            depData.docDeclaracaoTempoPermanencia = false;
        }

    } else if (prefix === 'dep') { // Campos de plano para 'depX' (seção "Relação de Dependente")
        depData.planoAmb = document.getElementById(`${prefix}${depIndex}_plano_amb`)?.checked || false;
        depData.planoComp = document.getElementById(`${prefix}${depIndex}_plano_comp`)?.checked || false;
        depData.planoOdonto = document.getElementById(`${prefix}${depIndex}_plano_odonto`)?.checked || false;
        depData.assistFuneral = document.getElementById(`${prefix}${depIndex}_assist_funeral`)?.checked || false;
    }

    return depData;
}


/**
 * Preenche os campos HTML de um dependente em uma seção específica.
 * @param {object} depData - O objeto com os dados do dependente.
 * @param {number} depIndex - O índice do dependente (1 a 8).
 * @param {string} prefix - O prefixo para os IDs/nomes dos campos (ex: 'dep', 'contratacao_dep').
 */
function setDependentDataToSection(depData, depIndex, prefix) {
    const nameInput = document.getElementById(`${prefix}${depIndex}_nome`);
    if (!nameInput) return; // Se o campo principal não existir, não faz nada

    nameInput.value = depData?.nome || '';
    const nascInput = document.getElementById(`${prefix}${depIndex}_nasc`);
    if (nascInput) { // Lida com campos de data
        const dateValue = depData?.nasc;
        if (typeof dateValue === 'string' && dateValue.includes('/')) { // dd/mm/yyyy
            const parts = dateValue.split('/');
            nascInput.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
        } else {
            nascInput.value = dateValue || '';
        }
    }
    const parentescoInput = document.getElementById(`${prefix}${depIndex}_parentesco`);
    if (parentescoInput) parentescoInput.value = depData?.parentesco || '';

    // Campos específicos para 'contratacao_depX'
    if (prefix === 'contratacao_dep') {
        const cpfInput = document.getElementById(`${prefix}${depIndex}_cpf`);
        if (cpfInput) cpfInput.value = depData?.cpf || '';
        const sexoInput = document.getElementById(`${prefix}${depIndex}_sexo`);
        if (sexoInput) sexoInput.value = depData?.sexo || '';
        const rgInput = document.getElementById(`${prefix}${depIndex}_rg`);
        if (rgInput) rgInput.value = depData?.rg || '';
        const dataEmissaoRgInput = document.getElementById(`${prefix}${depIndex}_data_emissao_rg`);
        if (dataEmissaoRgInput) { // Lida com campos de data
            const dateValue = depData?.dataEmissaoRg;
            if (typeof dateValue === 'string' && dateValue.includes('/')) { // dd/mm/yyyy
                const parts = dateValue.split('/');
                dataEmissaoRgInput.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            } else {
                dataEmissaoRgInput.value = dateValue || '';
            }
        }
        const orgaoEmissorInput = document.getElementById(`${prefix}${depIndex}_orgao_emissor`);
        if (orgaoEmissorInput) orgaoEmissorInput.value = depData?.orgaoEmissor || '';
        const estadoCivilInput = document.getElementById(`${prefix}${depIndex}_estado_civil`);
        if (estadoCivilInput) estadoCivilInput.value = depData?.estadoCivil || '';
        const cartaoNacionalSaudeInput = document.getElementById(`${prefix}${depIndex}_cartao_nacional_saude`);
        if (cartaoNacionalSaudeInput) cartaoNacionalSaudeInput.value = depData?.cartaoNacionalSaude || '';
        const nomeMaeInput = document.getElementById(`${prefix}${depIndex}_nome_mae`);
        if (nomeMaeInput) nomeMaeInput.value = depData?.nomeMae || '';

        // Radio buttons para plano anterior
        const planoAnteriorSim = document.getElementById(`${prefix}${depIndex}_plano_anterior_sim`);
        const planoAnteriorNao = document.getElementById(`${prefix}${depIndex}_plano_anterior_nao`);
        if (planoAnteriorSim) planoAnteriorSim.checked = (depData?.planoAnterior === 'Sim');
        if (planoAnteriorNao) planoAnteriorNao.checked = (depData?.planoAnterior === 'Não');

        const qualPlanoInput = document.getElementById(`${prefix}${depIndex}_qual_plano`);
        if (qualPlanoInput) qualPlanoInput.value = depData?.qualPlano || '';
        const dataInicioInput = document.getElementById(`${prefix}${depIndex}_data_inicio`);
        if (dataInicioInput) { // Lida com campos de data
            const dateValue = depData?.dataInicio;
            if (typeof dateValue === 'string' && dateValue.includes('/')) { // dd/mm/yyyy
                const parts = dateValue.split('/');
                dataInicioInput.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            } else {
                dataInicioInput.value = dateValue || '';
            }
        }
        const dataUltimoPagamentoInput = document.getElementById(`${prefix}${depIndex}_data_ultimo_pagamento`);
        if (dataUltimoPagamentoInput) { // Lida com campos de data
            const dateValue = depData?.dataUltimoPagamento;
            if (typeof dateValue === 'string' && dateValue.includes('/')) { // dd/mm/yyyy
                const parts = dateValue.split('/');
                dataUltimoPagamentoInput.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            } else {
                dataUltimoPagamentoInput.value = dateValue || '';
            }
        }
        const qualOperadoraAnteriorInput = document.getElementById(`${prefix}${depIndex}_qual_operadora_anterior`);
        if (qualOperadoraAnteriorInput) qualOperadoraAnteriorInput.value = depData?.qualOperadoraAnterior || '';
        const registroProdutoAnteriorInput = document.getElementById(`${prefix}${depIndex}_registro_produto_anterior`);
        if (registroProdutoAnteriorInput) registroProdutoAnteriorInput.value = depData?.registroProdutoAnterior || '';

        // Checkboxes de documentação
        const docCpfRgCnh = document.getElementById(`doc_entregue_${prefix}${depIndex}_cpf_rg_cnh`);
        if (docCpfRgCnh) docCpfRgCnh.checked = shouldBeChecked(depData?.docCpfRgCnh);

        if (depIndex === 1) { // Apenas para o Dep1
            const docFotoseIfie = document.getElementById(`doc_entregue_${prefix}${depIndex}_fotoseIfie`);
            if (docFotoseIfie) docFotoseIfie.checked = shouldBeChecked(depData?.docFotoseIfie);
            const docComprovanteResidencia = document.getElementById(`doc_entregue_${prefix}${depIndex}_comprovante_residencia`);
            if (docComprovanteResidencia) docComprovanteResidencia.checked = shouldBeChecked(depData?.docComprovanteResidencia);
            const docComprovanteVendaBeneficiario = document.getElementById(`doc_entregue_${prefix}${depIndex}_comprovante_venda_beneficiario`);
            if (docComprovanteVendaBeneficiario) docComprovanteVendaBeneficiario.checked = shouldBeChecked(depData?.docComprovanteVendaBeneficiario);
            const docOutros = document.getElementById(`doc_entregue_${prefix}${depIndex}_outros`);
            if (docOutros) docOutros.checked = shouldBeChecked(depData?.docOutros);
            const docDeclaracaoTempoPermanencia = document.getElementById(`doc_entregue_${prefix}${depIndex}_declaracao_tempo_permanencia`);
            if (docDeclaracaoTempoPermanencia) docDeclaracaoTempoPermanencia.checked = shouldBeChecked(depData?.docDeclaracaoTempoPermanencia);
        }

    } else if (prefix === 'dep') { // Campos de plano para 'depX' (seção "Relação de Dependente")
        const planoAmb = document.getElementById(`${prefix}${depIndex}_plano_amb`);
        if (planoAmb) planoAmb.checked = shouldBeChecked(depData?.planoAmb);
        const planoComp = document.getElementById(`${prefix}${depIndex}_plano_comp`);
        if (planoComp) planoComp.checked = shouldBeChecked(depData?.planoComp);
        const planoOdonto = document.getElementById(`${prefix}${depIndex}_plano_odonto`);
        if (planoOdonto) planoOdonto.checked = shouldBeChecked(depData?.planoOdonto);
        const assistFuneral = document.getElementById(`${prefix}${depIndex}_assist_funeral`);
        if (assistFuneral) assistFuneral.checked = shouldBeChecked(depData?.assistFuneral);
    } else if (prefix === 'odonto_dep') { // Campos específicos para 'odonto_depX'
        const cpfInput = document.getElementById(`${prefix}${depIndex}_cpf`);
        if (cpfInput) cpfInput.value = depData?.cpf || '';
        const sexoInput = document.getElementById(`${prefix}${depIndex}_sexo`);
        if (sexoInput) {
            const mappedValue = depData?.sexo === 'Masculino' ? 'M' : (depData?.sexo === 'Feminino' ? 'F' : '');
            sexoInput.value = mappedValue || '';
        }
        const estadoCivilInput = document.getElementById(`${prefix}${depIndex}_estado_civil`);
        if (estadoCivilInput) estadoCivilInput.value = depData?.estadoCivil || '';
        const maeInput = document.getElementById(`${prefix}${depIndex}_mae`);
        if (maeInput) maeInput.value = depData?.nomeMae || ''; // Mapeia 'nomeMae' para 'mae' no odonto_dep
    }
}


// Função principal para coletar todos os dados dos dependentes de suas seções primárias
function collectAllDependentData() {
    const allDependentsData = {};

    // Prioriza a seção "Proposta de Contratação - Dependentes" para dados mais completos
    for (let i = 1; i <= 4; i++) { // Max 4 dependentes na Proposta de Contratação
        const depData = getDependentDataFromSection(i, 'contratacao_dep');
        if (depData) {
            allDependentsData[i] = depData;
        }
    }

    // Complementa com dados da seção "Relação de Dependente" se não estiverem na primeira
    for (let i = 1; i <= 8; i++) { // Max 8 dependentes na Relação de Dependente
        const depDataUniao = getDependentDataFromSection(i, 'dep');
        if (depDataUniao) {
            // Se o dependente não foi coletado em 'contratacao_dep' ou tem dados adicionais aqui
            if (!allDependentsData[i]) {
                allDependentsData[i] = {};
            }
            // Copia/sobrescreve campos de nome, nasc, parentesco da seção 'dep'
            // Assumimos que 'dep' é a fonte principal para nome/nasc/parentesco
            allDependentsData[i].nome = depDataUniao.nome || allDependentsData[i].nome;
            allDependentsData[i].nasc = depDataUniao.nasc || allDependentsData[i].nasc;
            allDependentsData[i].parentesco = depDataUniao.parentesco || allDependentsData[i].parentesco;

            // Copia campos de plano da seção 'dep'
            allDependentsData[i].planoAmb = depDataUniao.planoAmb || allDependentsData[i].planoAmb || false;
            allDependentsData[i].planoComp = depDataUniao.planoComp || allDependentsData[i].planoComp || false;
            allDependentsData[i].planoOdonto = depDataUniao.planoOdonto || allDependentsData[i].planoOdonto || false;
            allDependentsData[i].assistFuneral = depDataUniao.assistFuneral || allDependentsData[i].assistFuneral || false;
        }
    }

    return allDependentsData;
}


// Função para distribuir os dados dos dependentes para todas as seções relevantes
function updateAllDependentFields() {
    const allDependentsData = collectAllDependentData();

    // Atualiza Relação de Dependentes (União de Benefícios) - dep1 a dep8
    for (let i = 1; i <= 8; i++) {
        setDependentDataToSection(allDependentsData[i], i, 'dep');
    }

    // Atualiza Proposta de Contratação - Dependentes (contratacao_dep1 a contratacao_dep4)
    for (let i = 1; i <= 4; i++) {
        setDependentDataToSection(allDependentsData[i], i, 'contratacao_dep');
    }

    // Atualiza Termo de Adesão Planos Odontológicos - Dependentes (odonto_dep1 a odonto_dep4)
    for (let i = 1; i <= 4; i++) {
        setDependentDataToSection(allDependentsData[i], i, 'odonto_dep');
    }

    // Atualiza Declaração de Saúde (Itens 1-21) - Colunas dos Dependentes (se necessário)
    // Se você tiver dados de "S/N" ou outros campos dos dependentes nessa tabela que precisam ser espelhados,
    // a lógica precisaria ser adicionada aqui, buscando de allDependentsData[i].sexo, allDependentsData[i].doenca etc.
    // Exemplo (apenas para ilustrar, o HTML atual não tem esses campos para dependentes na declaração):
    for (let item = 1; item <= 21; item++) {
        for (let depNum = 1; depNum <= 4; depNum++) {
            const selectElement = document.getElementById(`item${item}_dep${depNum}`);
            if (selectElement && allDependentsData[depNum]) {
                // Aqui você precisaria de uma chave correspondente em allDependentsData[depNum]
                // Ex: allDependentsData[depNum][`item${item}_resposta`]
                // Por agora, apenas um exemplo genérico:
                // selectElement.value = allDependentsData[depNum][`doenca_item${item}`] || '';
            }
        }
    }


    // Atualiza Informações Complementares (comp_info_X_depY)
    for (let compItem = 1; compItem <= 11; compItem++) {
        for (let depNum = 1; depNum <= 4; depNum++) {
            const checkbox = document.getElementById(`comp_info_${compItem}_dep${depNum}`);
            if (checkbox && allDependentsData[depNum]) {
                // Exemplo: Se allDependentsData[depNum][`comp_info_item${compItem}_checked`] for true
                // checkbox.checked = shouldBeChecked(allDependentsData[depNum][`comp_info_item${compItem}_checked`]);
            }
        }
    }
}


// Function to update Contratante fields based on titular data (União de Benefícios)
function updateContratanteFields() {
    if (contratacaoContratanteCpf) contratacaoContratanteCpf.value = cpfInput ? cpfInput.value : '';
    if (contratacaoContratanteNome) contratacaoContratanteNome.value = nomeInput ? nomeInput.value : '';
    if (contratacaoContratanteRg) contratacaoContratanteRg.value = rgInput ? rgInput.value : '';
    if (contratacaoContratanteOrgaoEmissor) contratacaoContratanteOrgaoEmissor.value = expInput ? expInput.value : '';

    if (contratacaoContratanteSexo) {
        const sourceValue = sexoInput ? sexoInput.value : '';
        const mappedValue = sourceValue === 'M' ? 'Masculino' : (sourceValue === 'F' ? 'Feminino' : '');
        const optionExists = Array.from(contratacaoContratanteSexo.options).some(option => option.value === mappedValue);
        contratacaoContratanteSexo.value = optionExists ? mappedValue : '';
        if (sourceValue !== '' && mappedValue === '' && !optionExists) {
            console.warn(`Sex value "${sourceValue}" from source field could not be mapped or is not a valid option in the target select.`);
        }
    }
    if (contratacaoContratanteEstadoCivil) contratacaoContratanteEstadoCivil.value = estCivilInput ? estCivilInput.value : '';
    if (contratacaoContratanteCelular) contratacaoContratanteCelular.value = celularInput ? celularInput.value : '';
    if (contratacaoContratanteWhatsapp) contratacaoContratanteWhatsapp.value = celularInput ? celularInput.value : '';
    if (contratacaoContratanteEmail) contratacaoContratanteEmail.value = emailInput ? emailInput.value : '';
    if (contratacaoContratanteNomeMae) contratacaoContratanteNomeMae.value = maeInput ? maeInput.value : '';
    if (contratacaoContratanteCep) contratacaoContratanteCep.value = cepInput ? cepInput.value : '';
    if (contratacaoContratanteLogradouro) contratacaoContratanteLogradouro.value = endInput ? endInput.value : '';
    if (contratacaoContratanteNumero) contratacaoContratanteNumero.value = numInput ? numInput.value : '';
    if (contratacaoContratanteComplemento) contratacaoContratanteComplemento.value = complInput ? complInput.value : '';
    if (contratacaoContratanteBairro) contratacaoContratanteBairro.value = bairroInput ? bairroInput.value : '';
    if (contratacaoContratanteCidade) contratacaoContratanteCidade.value = cidadeInput ? cidadeInput.value : '';
    if (contratacaoContratanteEstado) contratacaoContratanteEstado.value = estInput ? estInput.value : '';
    if (contratacaoContratanteDataNascimento) contratacaoContratanteDataNascimento.value = nascInput ? nascInput.value : '';

    if (contratacaoResumoCpfContratante) {
        contratacaoResumoCpfContratante.value = cpfInput ? cpfInput.value : '';
    }



if(odontoTitularNome) odontoTitularNome.value = nomeInput ? nomeInput.value : '';
if(odontoTitularCpf) odontoTitularCpf.value = cpfInput ? cpfInput.value : '';
if(odontoTitularDataNascimento) odontoTitularDataNascimento.value = nascInput ? nascInput.value : '';
if(odontoTitularRg) odontoTitularRg.value = rgInput ? rgInput.value : '';
if(odontoTitularSexo) odontoTitularSexo.value = sexoInput ? sexoInput.value : '';
if(odontoTitularEstadoCivil) odontoTitularEstadoCivil.value = estCivilInput ? estCivilInput.value : '';
if(odontoTitularEndereco) odontoTitularEndereco.value = endInput ? endInput.value : '';
if(odontoTitularNumero) odontoTitularNumero.value = numInput ? numInput.value : '';
if(odontoTitularBairro) odontoTitularBairro.value = bairroInput ? bairroInput.value : '';
if(odontoTitularCep) odontoTitularCep.value = cepInput ? cepInput.value : '';
if(odontoTitularCidade) odontoTitularCidade.value = cidadeInput ? cidadeInput.value : '';
if(odontoTitularTelefone) odontoTitularTelefone.value = celularInput ? celularInput.value : '';
if(odontoTitularMae) odontoTitularMae.value = maeInput ? maeInput.value : '';
if(odontoTitularEmail) odontoTitularEmail.value = emailInput ? emailInput.value : '';
}


// Function to update the Responsible fields based on titular data
function updateContratacaoResponsavelFields() {
    if (contratacaoResponsavelCpf) contratacaoResponsavelCpf.value = cpfInput ? cpfInput.value : '';
    if (contratacaoResponsavelNome) contratacaoResponsavelNome.value = nomeInput ? nomeInput.value : '';
    if (contratacaoResponsavelEmail) contratacaoResponsavelEmail.value = emailInput ? emailInput.value : '';
    if (contratacaoResponsavelNomeMae) contratacaoResponsavelNomeMae.value = maeInput ? maeInput.value : '';
    if (contratacaoResponsavelCelular) contratacaoResponsavelCelular.value = celularInput ? celularInput.value : '';
    if (contratacaoResponsavelDataNascimento) contratacaoResponsavelDataNascimento.value = nascInput ? nascInput.value : '';
    if (contratacaoResponsavelEstadoCivil) contratacaoResponsavelEstadoCivil.value = estCivilInput ? estCivilInput.value : '';
    if (contratacaoResponsavelRgCnh) contratacaoResponsavelRgCnh.value = rgInput ? rgInput.value : '';
    if (contratacaoResponsavelCep) contratacaoResponsavelCep.value = cepInput ? cepInput.value : '';
    if (contratacaoResponsavelLogradouro) contratacaoResponsavelLogradouro.value = endInput ? endInput.value : '';
    if (contratacaoResponsavelNumero) contratacaoResponsavelNumero.value = numInput ? numInput.value : '';
    if (contratacaoResponsavelComplemento) contratacaoResponsavelComplemento.value = complInput ? complInput.value : '';
    if (contratacaoResponsavelBairro) contratacaoResponsavelBairro.value = bairroInput ? bairroInput.value : '';
    if (contratacaoResponsavelCidade) contratacaoResponsavelCidade.value = cidadeInput ? cidadeInput.value : '';
    if (contratacaoResponsavelEstado) contratacaoResponsavelEstado.value = estInput ? estInput.value : '';

    if (sexoInput && contratacaoResponsavelSexoMasculino && contratacaoResponsavelSexoFeminino) {
        const sourceValue = sexoInput.value;
        if (sourceValue === 'M') {
            contratacaoResponsavelSexoMasculino.checked = true;
            contratacaoResponsavelSexoFeminino.checked = false;
        } else if (sourceValue === 'F') {
            contratacaoResponsavelSexoMasculino.checked = false;
            contratacaoResponsavelSexoFeminino.checked = true;
        } else {
            contratacaoResponsavelSexoMasculino.checked = false;
            contratacaoResponsavelSexoFeminino.checked = false;
        }
    }
}

// Function to set the current date in the "Data da proposta" field
function setContratacaoResumoDate() {
    if (contratacaoResumoDataProposta) {
        const today = new Date();
        const year = today.getFullYear();
        const month = ('0' + (today.getMonth() + 1)).slice(-2);
        const day = ('0' + today.getDate()).slice(-2);
        contratacaoResumoDataProposta.value = `${year}-${month}-${day}`;
    }
}

// --- Event Listeners para acionar as funções de atualização ---

const fieldsToWatch = [
    nomeInput, cpfInput, nascInput, sexoInput, estCivilInput, rgInput, expInput,
    maeInput, emailInput, celularInput, telInput, endInput, numInput, complInput,
    bairroInput, cepInput, cidadeInput, estInput
];

// Adiciona os campos de todas as seções de dependentes à lista para serem monitorados
// Isso garante que qualquer alteração em qualquer campo de dependente dispare a atualização
for (let i = 1; i <= 8; i++) {
    fieldsToWatch.push(document.getElementById(`dep${i}_nome`));
    fieldsToWatch.push(document.getElementById(`dep${i}_nasc`));
    fieldsToWatch.push(document.getElementById(`dep${i}_parentesco`));
    fieldsToWatch.push(document.getElementById(`dep${i}_plano_amb`));
    fieldsToWatch.push(document.getElementById(`dep${i}_plano_comp`));
    fieldsToWatch.push(document.getElementById(`dep${i}_plano_odonto`));
    fieldsToWatch.push(document.getElementById(`dep${i}_assist_funeral`));

    // Campos da Proposta de Contratação - Dependentes (até o 4º)
    if (i <= 4) {
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_cpf`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_nome`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_sexo`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_rg`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_data_emissao_rg`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_orgao_emissor`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_estado_civil`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_data_nascimento`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_cartao_nacional_saude`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_parentesco`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_nome_mae`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_plano_anterior_sim`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_plano_anterior_nao`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_qual_plano`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_data_inicio`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_data_ultimo_pagamento`));
        fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_cpf_rg_cnh`));
        if (i === 1) { // Checkboxes específicos do Dep1
            fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_fotoseIfie`));
            fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_comprovante_residencia`));
            fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_comprovante_venda_beneficiario`));
            fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_outros`));
            fieldsToWatch.push(document.getElementById(`doc_entregue_dep${i}_declaracao_tempo_permanencia`));
        }
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_qual_operadora_anterior`));
        fieldsToWatch.push(document.getElementById(`contratacao_dep${i}_registro_produto_anterior`));
    }
    
    // Campos do Termo de Adesão Odontológicos - Dependentes (até o 4º)
    if (i <= 4) {
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_nome`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_cpf`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_data_nascimento`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_sexo`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_estado_civil`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_mae`));
        fieldsToWatch.push(document.getElementById(`odonto_dep${i}_parentesco`));
    }

    // Campos da Declaração de Saúde (itens da tabela)
    for (let item = 1; item <= 21; item++) {
        fieldsToWatch.push(document.getElementById(`item${item}_titular`));
        for (let depNum = 1; depNum <= 4; depNum++) {
            fieldsToWatch.push(document.getElementById(`item${item}_dep${depNum}`));
        }
    }
    // Campos IMC
    fieldsToWatch.push(bmiTitularPeso);
    fieldsToWatch.push(bmiDep1Peso);
    fieldsToWatch.push(bmiDep2Peso);
    fieldsToWatch.push(bmiDep3Peso);
    fieldsToWatch.push(bmiDep4Peso);
    fieldsToWatch.push(bmiTitularAltura);
    fieldsToWatch.push(bmiDep1Altura);
    fieldsToWatch.push(bmiDep2Altura);
    fieldsToWatch.push(bmiDep3Altura);
    fieldsToWatch.push(bmiDep4Altura);

    // Campos Informações Complementares
    for (let compItem = 1; compItem <= 11; compItem++) {
        fieldsToWatch.push(document.getElementById(`comp_info_item_num_${compItem}`));
        fieldsToWatch.push(document.getElementById(`comp_info_description_${compItem}`));
        fieldsToWatch.push(document.getElementById(`comp_info_${compItem}_titular`));
        fieldsToWatch.push(document.getElementById(`comp_info_year_${compItem}`));
        for (let depNum = 1; depNum <= 4; depNum++) {
            fieldsToWatch.push(document.getElementById(`comp_info_${compItem}_dep${depNum}`));
        }
    }
}


fieldsToWatch.forEach(field => {
    if (field) {
        field.addEventListener('input', () => {
            updateContratanteFields(); // Atualiza campos do contratante
            updateContratacaoResponsavelFields(); // Atualiza campos do responsável
            updateAllDependentFields(); // **Chama a função que coordena a atualização de todos os dependentes**
        });
        field.addEventListener('change', () => { // Para selects e radios
            updateContratanteFields();
            updateContratacaoResponsavelFields();
            updateAllDependentFields();
        });
    }
});


// Listener específico para CPF para validação no Gerar Arquivo
cpfInput?.addEventListener('blur', () => {
    if (cpfInput.value && !validateCPF(cpfInput.value)) {
        cpfInput.classList.add('invalid');
    } else {
        cpfInput.classList.remove('invalid');
    }
});


// --- FUNÇÃO PARA PEGAR TODOS OS DADOS DO FORMULÁRIO (HTML -> JS `data` object) ---
function getFormData() {
    const data = {};

    // Helper para coletar valor de input/select
    const getValue = (element) => element ? element.value || '' : '';
    // Helper para coletar estado de checkbox
    const getChecked = (element) => element ? element.checked : false;

    // Cabeçalho e Dados Pessoais do Titular
    data['proposal_number_uniao'] = getValue(proposalNumberUniao);
    data['orgao'] = getValue(orgaoInput);
    data['proponente'] = getValue(proponenteInput);
    data['nome'] = getValue(nomeInput);
    data['nasc'] = getValue(nascInput);
    data['sexo'] = getValue(sexoInput);
    data['est_civil'] = getValue(estCivilInput);
    data['rg'] = getValue(rgInput);
    data['exp'] = getValue(expInput);
    data['cpf'] = getValue(cpfInput);
    data['mae'] = getValue(maeInput);
    data['email'] = getValue(emailInput);
    data['banco'] = getValue(bancoInput);
    data['agencia'] = getValue(agenciaInput);
    data['conta_corrente'] = getValue(contaCorrenteInput);
    data['conjuge'] = getValue(conjugeInput);
    data['nasc_conjuge'] = getValue(nascConjugeInput);
    data['sexo_conjuge'] = getValue(sexoConjugeInput);
    data['end'] = getValue(endInput);
    data['num'] = getValue(numInput);
    data['compl'] = getValue(complInput);
    data['bairro'] = getValue(bairroInput);
    data['cep'] = getValue(cepInput);
    data['cidade'] = getValue(cidadeInput);
    data['est'] = getValue(estInput);
    data['tel'] = getValue(telInput);
    data['celular'] = getValue(celularInput);
    data['orgao_funcional'] = getValue(orgaoFuncionalInput);
    data['mat_funcional'] = getValue(matFuncionalInput);
    data['funcao'] = getValue(funcaoInput);
    data['unidade'] = getValue(unidadeInput);

    // Benefícios Contratados
    data['mensalidade_social'] = getChecked(MensalidadeSocial);
    data['plano_saude_amb'] = getChecked(planoSaudeAmb);
    data['plano_saude_comp'] = getChecked(PlanoSaudeComp);
    data['orientacao_juridica'] = getChecked(OrientacaoJuridica);
    data['plano_odonto'] = getChecked(PlanoOdonto);
    data['seguro_vida'] = getChecked(SeguroVida);
    // Para radio buttons de Seguro de Vida, verifica qual está marcado
    if (getChecked(SeguroVidaSim)) {
        data['seguro_vida_opcao'] = SeguroVidaSim.value;
    } else if (getChecked(SeguroVidaNao)) {
        data['seguro_vida_opcao'] = SeguroVidaNao.value;
    } else {
        data['seguro_vida_opcao'] = '';
    }
    data['auxilio_natalidade'] = getChecked(AuxilioNatalidade);
    data['assistencia_funeral'] = getChecked(AssistenciaFuneral);
    data['convenios'] = getChecked(Convenios);

    // Relação de Dependentes (União de Benefícios)
    for (let i = 1; i <= 8; i++) {
        const depName = document.getElementById(`dep${i}_nome`);
        if (depName && depName.value.trim() !== '') { // Coleta dados apenas se o nome do dependente for preenchido
            data[`dep${i}_nome`] = getValue(depName);
            data[`dep${i}_nasc`] = getValue(document.getElementById(`dep${i}_nasc`));
            data[`dep${i}_parentesco`] = getValue(document.getElementById(`dep${i}_parentesco`));
            data[`dep${i}_plano_amb`] = getChecked(document.getElementById(`dep${i}_plano_amb`));
            data[`dep${i}_plano_comp`] = getChecked(document.getElementById(`dep${i}_plano_comp`));
            data[`dep${i}_plano_odonto`] = getChecked(document.getElementById(`dep${i}_plano_odonto`));
            data[`dep${i}_assist_funeral`] = getChecked(document.getElementById(`dep${i}_assist_funeral`));
        } else { // Se o nome do dependente estiver vazio, limpa dados anteriores para esse dependente
            data[`dep${i}_nome`] = '';
            data[`dep${i}_nasc`] = '';
            data[`dep${i}_parentesco`] = '';
            data[`dep${i}_plano_amb`] = false;
            data[`dep${i}_plano_comp`] = false;
            data[`dep${i}_plano_odonto`] = false;
            data[`dep${i}_assist_funeral`] = false;
        }
    }

    // Termos e Declaração
    data['declaration_rs_value'] = getValue(declarationRsValue);
    data['declaration_local'] = getValue(declarationLocal);
    data['declaration_dia'] = getValue(declarationDia);
    data['declaration_mes'] = getValue(declarationMes);
    data['declaration_ano'] = getValue(declarationAno);
    data['agent_name'] = getValue(agentName);
    data['agent_registro'] = getValue(agentRegistro);

    // Proposta de Contratação - Dados do Vendedor
    data['contratacao_proposal_number'] = getValue(contratacaoProposalNumber);
    data['contratacao_vendedor_cpf'] = getValue(contratacaoVendedorCpf);
    data['contratacao_vendedor_nome'] = getValue(contratacaoVendedorNome);
    data['contratacao_vendedor_telefone'] = getValue(contratacaoVendedorTelefone);
    data['contratacao_cnpj_corretora'] = getValue(contratacaoCnpjCorretora);
    data['contratacao_nome_corretora'] = getValue(contratacaoNomeCorretora);

    // Proposta de Contratação - Dados cadastrais do Contratante
    data['contratacao_contratante_responsavel'] = getChecked(contratacaoContratanteResponsavel);
    data['contratacao_contratante_cpf'] = getValue(contratacaoContratanteCpf);
    data['contratacao_contratante_nome'] = getValue(contratacaoContratanteNome);
    data['contratacao_contratante_sexo'] = getValue(contratacaoContratanteSexo);
    data['contratacao_contratante_rg'] = getValue(contratacaoContratanteRg);
    data['contratacao_contratante_data_emissao_rg'] = getValue(contratacaoContratanteDataEmissaoRg);
    data['contratacao_contratante_orgao_emissor'] = getValue(contratacaoContratanteOrgaoEmissor);
    data['contratacao_contratante_estado_civil'] = getValue(contratacaoContratanteEstadoCivil);
    data['contratacao_contratante_data_nascimento'] = getValue(contratacaoContratanteDataNascimento);
    data['contratacao_contratante_telefone_fixo'] = getValue(contratacaoContratanteTelefoneFixo);
    data['contratacao_contratante_celular'] = getValue(contratacaoContratanteCelular);
    data['contratacao_contratante_whatsapp'] = getValue(contratacaoContratanteWhatsapp);
    data['contratacao_contratante_cartao_nacional_saude'] = getValue(contratacaoContratanteCartaoNacionalSaude);
    data['contratacao_contratante_email'] = getValue(contratacaoContratanteEmail);
    data['contratacao_contratante_nome_mae'] = getValue(contratacaoContratanteNomeMae);
    // Radio buttons: Contratante - Possui Plano Anterior?
    if (getChecked(contratacaoContratantePlanoAnteriorSim)) {
        data['contratacao_contratante_plano_anterior'] = contratacaoContratantePlanoAnteriorSim.value;
    } else if (getChecked(contratacaoContratantePlanoAnteriorNao)) {
        data['contratacao_contratante_plano_anterior'] = contratacaoContratantePlanoAnteriorNao.value;
    } else {
        data['contratacao_contratante_plano_anterior'] = '';
    }
    data['contratacao_contratante_qual_plano'] = getValue(contratacaoContratanteQualPlano);
    data['contratacao_contratante_data_inicio'] = getValue(contratacaoContratanteDataInicio);
    data['contratacao_contratante_data_ultimo_pagamento'] = getValue(contratacaoContratanteDataUltimoPagamento);
    data['contratacao_contratante_cep'] = getValue(contratacaoContratanteCep);
    data['contratacao_contratante_logradouro'] = getValue(contratacaoContratanteLogradouro);
    data['contratacao_contratante_numero'] = getValue(contratacaoContratanteNumero);
    data['contratacao_contratante_complemento'] = getValue(contratacaoContratanteComplemento);
    data['contratacao_contratante_bairro'] = getValue(contratacaoContratanteBairro);
    data['contratacao_contratante_cidade'] = getValue(contratacaoContratanteCidade);
    data['contratacao_contratante_estado'] = getValue(contratacaoContratanteEstado);

    // Documentação entregue (Contratante)
    data['doc_entregue_cpf_rg_cnh'] = getChecked(docEntregueCpfRgCnh);
    data['doc_entregue_fotoseIfie'] = getChecked(docEntregueFotoseIfie);
    data['doc_entregue_comprovante_residencia'] = getChecked(docEntregueComprovanteResidencia);
    data['doc_entregue_comprovante_venda_beneficiario'] = getChecked(docEntregueComprovanteVendaBeneficiario);
    data['doc_entregue_outros'] = getChecked(docEntregueOutros);
    data['doc_entregue_declaracao_tempo_permanencia'] = getChecked(docEntregueDeclaracaoTempoPermanencia);

    // Proposta de Contratação - Dados cadastrais dos Dependentes (contratacao_depX)
    for (let i = 1; i <= 4; i++) {
        const depCpf = document.getElementById(`contratacao_dep${i}_cpf`);
        if (depCpf && depCpf.value.trim() !== '') { // Só coleta se o CPF do dependente estiver preenchido
            data[`contratacao_dep${i}_cpf`] = getValue(depCpf);
            data[`contratacao_dep${i}_nome`] = getValue(document.getElementById(`contratacao_dep${i}_nome`));
            data[`contratacao_dep${i}_sexo`] = getValue(document.getElementById(`contratacao_dep${i}_sexo`));
            data[`contratacao_dep${i}_rg`] = getValue(document.getElementById(`contratacao_dep${i}_rg`));
            data[`contratacao_dep${i}_data_emissao_rg`] = getValue(document.getElementById(`contratacao_dep${i}_data_emissao_rg`));
            data[`contratacao_dep${i}_orgao_emissor`] = getValue(document.getElementById(`contratacao_dep${i}_orgao_emissor`));
            data[`contratacao_dep${i}_estado_civil`] = getValue(document.getElementById(`contratacao_dep${i}_estado_civil`));
            data[`contratacao_dep${i}_data_nascimento`] = getValue(document.getElementById(`contratacao_dep${i}_data_nascimento`));
            data[`contratacao_dep${i}_cartao_nacional_saude`] = getValue(document.getElementById(`contratacao_dep${i}_cartao_nacional_saude`));
            data[`contratacao_dep${i}_parentesco`] = getValue(document.getElementById(`contratacao_dep${i}_parentesco`));
            data[`contratacao_dep${i}_nome_mae`] = getValue(document.getElementById(`contratacao_dep${i}_nome_mae`));

            // Radio buttons: Dependente X - Possui Plano Anterior?
            const depXPlanoAnteriorSim = document.getElementById(`contratacao_dep${i}_plano_anterior_sim`);
            const depXPlanoAnteriorNao = document.getElementById(`contratacao_dep${i}_plano_anterior_nao`);
            if (getChecked(depXPlanoAnteriorSim)) {
                data[`contratacao_dep${i}_plano_anterior`] = depXPlanoAnteriorSim.value;
            } else if (getChecked(depXPlanoAnteriorNao)) {
                data[`contratacao_dep${i}_plano_anterior`] = depXPlanoAnteriorNao.value;
            } else {
                data[`contratacao_dep${i}_plano_anterior`] = '';
            }

            data[`contratacao_dep${i}_qual_plano`] = getValue(document.getElementById(`contratacao_dep${i}_qual_plano`));
            data[`contratacao_dep${i}_data_inicio`] = getValue(document.getElementById(`contratacao_dep${i}_data_inicio`));
            data[`contratacao_dep${i}_data_ultimo_pagamento`] = getValue(document.getElementById(`contratacao_dep${i}_data_ultimo_pagamento`));

            // Checkboxes: Dependente X - Documentação entregue
            data[`doc_entregue_dep${i}_cpf_rg_cnh`] = getChecked(document.getElementById(`doc_entregue_dep${i}_cpf_rg_cnh`));
            // Apenas Dep1 tem todos os checkboxes no HTML fornecido para 'doc_entregue_depX'.
            // Para outros dep's, eles são false por padrão se o HTML não os tiver.
            if (i === 1) {
                data[`doc_entregue_dep${i}_fotoseIfie`] = getChecked(document.getElementById(`doc_entregue_dep${i}_fotoseIfie`));
                data[`doc_entregue_dep${i}_comprovante_residencia`] = getChecked(document.getElementById(`doc_entregue_dep${i}_comprovante_residencia`));
                data[`doc_entregue_dep${i}_comprovante_venda_beneficiario`] = getChecked(document.getElementById(`doc_entregue_dep${i}_comprovante_venda_beneficiario`));
                data[`doc_entregue_dep${i}_outros`] = getChecked(document.getElementById(`doc_entregue_dep${i}_outros`));
                data[`doc_entregue_dep${i}_declaracao_tempo_permanencia`] = getChecked(document.getElementById(`doc_entregue_dep${i}_declaracao_tempo_permanencia`));
            } else {
                 // Certifique-se de que essas chaves existam mesmo se não houver um elemento correspondente no HTML
                data[`doc_entregue_dep${i}_fotoseIfie`] = false;
                data[`doc_entregue_dep${i}_comprovante_residencia`] = false;
                data[`doc_entregue_dep${i}_comprovante_venda_beneficiario`] = false;
                data[`doc_entregue_dep${i}_outros`] = false;
                data[`doc_entregue_dep${i}_declaracao_tempo_permanencia`] = false;
            }

            data[`contratacao_dep${i}_qual_operadora_anterior`] = getValue(document.getElementById(`contratacao_dep${i}_qual_operadora_anterior`));
            data[`contratacao_dep${i}_registro_produto_anterior`] = getValue(document.getElementById(`contratacao_dep${i}_registro_produto_anterior`));
        } else { // Se o CPF do dependente estiver vazio, limpa todos os dados para esse dependente
            data[`contratacao_dep${i}_cpf`] = '';
            data[`contratacao_dep${i}_nome`] = '';
            data[`contratacao_dep${i}_sexo`] = '';
            data[`contratacao_dep${i}_rg`] = '';
            data[`contratacao_dep${i}_data_emissao_rg`] = '';
            data[`contratacao_dep${i}_orgao_emissor`] = '';
            data[`contratacao_dep${i}_estado_civil`] = '';
            data[`contratacao_dep${i}_data_nascimento`] = '';
            data[`contratacao_dep${i}_cartao_nacional_saude`] = '';
            data[`contratacao_dep${i}_parentesco`] = '';
            data[`contratacao_dep${i}_nome_mae`] = '';
            data[`contratacao_dep${i}_plano_anterior`] = '';
            data[`contratacao_dep${i}_qual_plano`] = '';
            data[`contratacao_dep${i}_data_inicio`] = '';
            data[`contratacao_dep${i}_data_ultimo_pagamento`] = '';
            data[`doc_entregue_dep${i}_cpf_rg_cnh`] = false;
            data[`doc_entregue_dep${i}_fotoseIfie`] = false;
            data[`doc_entregue_dep${i}_comprovante_residencia`] = false;
            data[`doc_entregue_dep${i}_comprovante_venda_beneficiario`] = false;
            data[`doc_entregue_dep${i}_outros`] = false;
            data[`doc_entregue_dep${i}_declaracao_tempo_permanencia`] = false;
            data[`contratacao_dep${i}_qual_operadora_anterior`] = '';
            data[`contratacao_dep${i}_registro_produto_anterior`] = '';
        }
    }


    // Responsável pelo Contrato
    data['contratacao_proposal_number_responsavel'] = getValue(contratacaoProposalNumberResponsavel);
    data['contratacao_responsavel_cpf'] = getValue(contratacaoResponsavelCpf);
    data['contratacao_responsavel_nome'] = getValue(contratacaoResponsavelNome);
    // Radio buttons: Responsável - Sexo
    if (getChecked(contratacaoResponsavelSexoMasculino)) {
        data['contratacao_responsavel_sexo'] = contratacaoResponsavelSexoMasculino.value;
    } else if (getChecked(contratacaoResponsavelSexoFeminino)) {
        data['contratacao_responsavel_sexo'] = contratacaoResponsavelSexoFeminino.value;
    } else {
        data['contratacao_responsavel_sexo'] = '';
    }
    data['contratacao_responsavel_email'] = getValue(contratacaoResponsavelEmail);
    data['contratacao_responsavel_nome_mae'] = getValue(contratacaoResponsavelNomeMae);
    data['contratacao_responsavel_celular'] = getValue(contratacaoResponsavelCelular);
    data['contratacao_responsavel_data_nascimento'] = getValue(contratacaoResponsavelDataNascimento);
    data['contratacao_responsavel_estado_civil'] = getValue(contratacaoResponsavelEstadoCivil);
    data['contratacao_responsavel_rg_cnh'] = getValue(contratacaoResponsavelRgCnh);
    data['contratacao_responsavel_data_emissao_rg_cnh'] = getValue(contratacaoResponsavelDataEmissaoRgCnh);
    data['contratacao_responsavel_cep'] = getValue(contratacaoResponsavelCep);
    data['contratacao_responsavel_logradouro'] = getValue(contratacaoResponsavelLogradouro);
    data['contratacao_responsavel_numero'] = getValue(contratacaoResponsavelNumero);
    data['contratacao_responsavel_complemento'] = getValue(contratacaoResponsavelComplemento);
    data['contratacao_responsavel_bairro'] = getValue(contratacaoResponsavelBairro);
    data['contratacao_responsavel_cidade'] = getValue(contratacaoResponsavelCidade);
    data['contratacao_responsavel_estado'] = getValue(contratacaoResponsavelEstado);

    // Documentação entregue (Responsável)
    data['doc_entregue_resp_cpf_rg_cnh'] = getChecked(docEntregueRespCpfRgCnh);
    data['doc_entregue_resp_comprovante_residencia'] = getChecked(docEntregueRespComprovanteResidencia);
    data['doc_entregue_resp_outros'] = getChecked(docEntregueRespOutros);

    // Resumo da Contratação
    data['contratacao_resumo_data_proposta'] = getValue(contratacaoResumoDataProposta);
    data['contratacao_resumo_provavel_vigencia'] = getValue(contratacaoResumoProvavelVigencia);
    data['contratacao_resumo_cpf_contratante'] = getValue(contratacaoResumoCpfContratante);
    data['contratacao_resumo_tipo'] = getValue(contratacaoResumoTipo);
    data['contratacao_resumo_segmentacao'] = getValue(contratacaoResumoSegmentacao);
    data['contratacao_resumo_plano'] = getValue(contratacaoResumoPlano);
    data['contratacao_resumo_registro_ans'] = getValue(contratacaoResumoRegistroAns);
    // Radio buttons: Acomodação
    if (getChecked(acomodacaoQc)) {
        data['contratacao_resumo_acomodacao'] = acomodacaoQc.value;
    } else if (getChecked(acomodacaoQp)) {
        data['contratacao_resumo_acomodacao'] = acomodacaoQp.value;
    } else {
        data['contratacao_resumo_acomodacao'] = '';
    }
    data['contratacao_resumo_abrangencia'] = getValue(contratacaoResumoAbrangencia);
    data['contratacao_resumo_coparticipacao'] = getChecked(coparticipacaoNao) ? coparticipacaoNao.value : '';
    data['contratacao_resumo_beneficiarios'] = getValue(contratacaoResumoBeneficiarios);
    data['contratacao_resumo_valor_total'] = getValue(contratacaoResumoValorTotal);

    // Declaração de Saúde (Itens 1-21 e IMC)
    data['declaracao_saude_proposal_number'] = getValue(declaracaoSaudeProposalNumber);
    data['declaracao_saude_proposal_number_cont'] = getValue(declaracaoSaudeProposalNumberCont);
    for (let i = 1; i <= 21; i++) {
        data[`item${i}_titular`] = getValue(window[`item${i}Titular`]);
        data[`item${i}_dep1`] = getValue(window[`item${i}Dep1`]);
        data[`item${i}_dep2`] = getValue(window[`item${i}Dep2`]);
        data[`item${i}_dep3`] = getValue(window[`item${i}Dep3`]);
        data[`item${i}_dep4`] = getValue(window[`item${i}Dep4`]);
    }
    data['bmi_titular_peso'] = getValue(bmiTitularPeso);
    data['bmi_dep1_peso'] = getValue(bmiDep1Peso);
    data['bmi_dep2_peso'] = getValue(bmiDep2Peso);
    data['bmi_dep3_peso'] = getValue(bmiDep3Peso);
    data['bmi_dep4_peso'] = getValue(bmiDep4Peso);
    data['bmi_titular_altura'] = getValue(bmiTitularAltura);
    data['bmi_dep1_altura'] = getValue(bmiDep1Altura);
    data['bmi_dep2_altura'] = getValue(bmiDep2Altura);
    data['bmi_dep3_altura'] = getValue(bmiDep3Altura);
    data['bmi_dep4_altura'] = getValue(bmiDep4Altura);

    // Informações Complementares (comp_info_item_num_1 a comp_info_year_11)
    data['complementary_info_proposal_number'] = getValue(complementaryInfoProposalNumber);
    for (let i = 1; i <= 11; i++) {
        data[`comp_info_item_num_${i}`] = getValue(window[`compInfoItemNum${i}`]);
        data[`comp_info_description_${i}`] = getValue(window[`compInfoDescription${i}`]);
        data[`comp_info_${i}_titular`] = getChecked(window[`compInfo${i}Titular`]);
        data[`comp_info_${i}_dep1`] = getChecked(window[`compInfo${i}Dep1`]);
        data[`comp_info_${i}_dep2`] = getChecked(window[`compInfo${i}Dep2`]);
        data[`comp_info_${i}_dep3`] = getChecked(window[`compInfo${i}Dep3`]);
        data[`comp_info_${i}_dep4`] = getChecked(window[`compInfo${i}Dep4`]);
        data[`comp_info_year_${i}`] = getValue(window[`compInfoYear${i}`]);
    }

    // Declaração de Saúde Final
    data['declaracao_saude_final_proposal_number'] = getValue(declaracaoSaudeFinalProposalNumber);
    data['entrevista_qualificada_opcao1_chk'] = getChecked(entrevistaQualificadaOpcao1Chk);
    data['entrevista_qualificada_opcao2_chk'] = getChecked(entrevistaQualificadaOpcao2Chk);
    data['entrevista_qualificada_opcao3_chk'] = getChecked(entrevistaQualificadaOpcao3Chk);

    // Carta de Orientação ao Beneficiário - Assinaturas
    data['carta_beneficiario_nome_sig'] = getValue(cartaBeneficiarioNomeSig);
    data['carta_beneficiario_local_sig'] = getValue(cartaBeneficiarioLocalSig);
    data['carta_beneficiario_data_sig'] = getValue(cartaBeneficiarioDataSig);
    data['carta_intermediario_nome_sig'] = getValue(cartaIntermediarioNomeSig);
    data['carta_intermediario_cpf_sig'] = getValue(cartaIntermediarioCpfSig);
    data['carta_intermediario_local_sig'] = getValue(cartaIntermediarioLocalSig);
    data['carta_intermediario_data_sig'] = getValue(cartaIntermediarioDataSig);

    // Termo Único de Promoções
    data['termo_promocoes_proposal_number'] = getValue(termoPromocoesProposalNumber);

    // Declaração de Recebimento e Posse
    data['declaracao_recebimento_proposal_number'] = getValue(declaracaoRecebimentoProposalNumber);

    // Termo de Consentimento
    data['termo_consentimento_proposal_number'] = getValue(termoConsentimentoProposalNumber);

    // Termo de Adesão Planos Odontológicos - Titular
    data['odonto_titular_nome'] = getValue(odontoTitularNome);
    data['odonto_titular_cpf'] = getValue(odontoTitularCpf);
    data['odonto_titular_data_nascimento'] = getValue(odontoTitularDataNascimento);
    data['odonto_titular_rg'] = getValue(odontoTitularRg);
    data['odonto_titular_sexo'] = getValue(odontoTitularSexo);
    data['odonto_titular_estado_civil'] = getValue(odontoTitularEstadoCivil);
    data['odonto_titular_endereco'] = getValue(odontoTitularEndereco);
    data['odonto_titular_numero'] = getValue(odontoTitularNumero);
    data['odonto_titular_bairro'] = getValue(odontoTitularBairro);
    data['odonto_titular_cep'] = getValue(odontoTitularCep);
    data['odonto_titular_cidade'] = getValue(odontoTitularCidade);
    data['odonto_titular_telefone'] = getValue(odontoTitularTelefone);
    data['odonto_titular_mae'] = getValue(odontoTitularMae);
    data['odonto_titular_email'] = getValue(odontoTitularEmail);

    // Termo de Adesão Planos Odontológicos - Dependentes 1-4
    for (let i = 1; i <= 4; i++) {
        const depOdontoNome = document.getElementById(`odonto_dep${i}_nome`);
        if (depOdontoNome && depOdontoNome.value.trim() !== '') {
            data[`odonto_dep${i}_nome`] = getValue(depOdontoNome);
            data[`odonto_dep${i}_cpf`] = getValue(document.getElementById(`odonto_dep${i}_cpf`));
            data[`odonto_dep${i}_data_nascimento`] = getValue(document.getElementById(`odonto_dep${i}_data_nascimento`));
            data[`odonto_dep${i}_sexo`] = getValue(document.getElementById(`odonto_dep${i}_sexo`));
            data[`odonto_dep${i}_estado_civil`] = getValue(document.getElementById(`odonto_dep${i}_estado_civil`));
            data[`odonto_dep${i}_mae`] = getValue(document.getElementById(`odonto_dep${i}_mae`));
            data[`odonto_dep${i}_parentesco`] = getValue(document.getElementById(`odonto_dep${i}_parentesco`));
        } else {
            // Limpa os dados se o nome do dependente estiver vazio
            data[`odonto_dep${i}_nome`] = '';
            data[`odonto_dep${i}_cpf`] = '';
            data[`odonto_dep${i}_data_nascimento`] = '';
            data[`odonto_dep${i}_sexo`] = '';
            data[`odonto_dep${i}_estado_civil`] = '';
            data[`odonto_dep${i}_mae`] = '';
            data[`odonto_dep${i}_parentesco`] = '';
        }
    }

    // Autorização para Desconto
    data['nomeDescontoAutoriza'] = getValue(nomeDescontoAutoriza);
    data['cpfDescontoAutoriza'] = getValue(cpfDescontoAutoriza);
    data['matFuncionalDescontoAutoriza'] = getValue(matFuncionalDescontoAutoriza);
    data['identidadeDescontoAutoriza'] = getValue(identidadeDescontoAutoriza);
    data['diaDescontoAutoriza'] = getValue(diaDescontoAutoriza);
    data['mesDescontoAutoriza'] = getValue(mesDescontoAutoriza);
    data['anosDescontoAutoriza'] = getValue(anosDescontoAutoriza);
    data['valorDescontoAutoriza'] = getValue(valorDescontoAutoriza);
    data['totalDescontoAutoriza'] = getValue(totalDescontoAutoriza);
    data['mensalidadeSocialDescontoAutoriza'] = getChecked(mensalidadeSocialDescontoAutoriza);
    data['planoSaudeAmbulatorialDescontoAutoriza'] = getChecked(planoSaudeAmbulatorialDescontoAutoriza);
    data['orientacaoJuridicaDescontoAutoriza'] = getChecked(orientacaoJuridicaDescontoAutoriza);
    data['segurosDescontoAutoriza'] = getChecked(segurosDescontoAutoriza);
    data['auxilioNatalidadeDescontoAutoriza'] = getChecked(auxilioNatalidadeDescontoAutoriza);
    data['planoSaudeCompletoDescontoAutoriza'] = getChecked(planoSaudeCompletoDescontoAutoriza);
    data['atdDomiciliarDescontoAutoriza'] = getChecked(atdDomiciliarDescontoAutoriza);
    data['assistenciaFuneralDescontoAutoriza'] = getChecked(assistenciaFuneralDescontoAutoriza);
    data['planoOdontologicoDescontoAutoriza'] = getChecked(planoOdontologicoDescontoAutoriza);
    data['conveniosDescontoAutoriza'] = getChecked(conveniosDescontoAutoriza);
    data['bancoDescontoAutoriza'] = getValue(bancoDescontoAutoriza);
    data['agenciaDescontoAutoriza'] = getValue(agenciaDescontoAutoriza);
    data['contaCorrenteDescontoAutoriza'] = getValue(contaCorrenteDescontoAutoriza);
    data['valorCobrancaDescontoAutoriza'] = getValue(valorCobrancaDescontoAutoriza);
    data['despesaBancoDescontoAutoriza'] = getValue(despesaBancoDescontoAutoriza);
    data['valorTotalDescontoAutoriza'] = getValue(valorTotalDescontoAutoriza);
    data['dataAssDescontoAutoriza'] = getValue(dataAssDescontoAutoriza);
    data['assinaturaDescontoAutoriza'] = getValue(assinaturaDescontoAutoriza);


    return data;
}

// --- FUNÇÃO PARA PREENCHER O FORMULÁRIO (JS `data` object -> HTML) ---
function setFormData(data) {
    // Helper para definir valor de input/select
    const setValue = (element, key) => {
        if (element && data.hasOwnProperty(key)) {
            element.value = data[key] || '';
        }
    };
    // Helper para definir estado de checkbox
    const setChecked = (element, key) => {
        if (element && data.hasOwnProperty(key)) {
            element.checked = shouldBeChecked(data[key]);
        }
    };

    // Cabeçalho e Dados Pessoais do Titular
    setValue(proposalNumberUniao, 'proposal_number_uniao');
    setValue(orgaoInput, 'orgao');
    setValue(proponenteInput, 'proponente');
    setValue(nomeInput, 'nome');
    setValue(nascInput, 'nasc');
    setValue(sexoInput, 'sexo');
    setValue(estCivilInput, 'est_civil');
    setValue(rgInput, 'rg');
    setValue(expInput, 'exp');
    setValue(cpfInput, 'cpf');
    setValue(maeInput, 'mae');
    setValue(emailInput, 'email');
    setValue(bancoInput, 'banco');
    setValue(agenciaInput, 'agencia');
    setValue(contaCorrenteInput, 'conta_corrente');
    setValue(conjugeInput, 'conjuge');
    setValue(nascConjugeInput, 'nasc_conjuge');
    setValue(sexoConjugeInput, 'sexo_conjuge');
    setValue(endInput, 'end');
    setValue(numInput, 'num');
    setValue(complInput, 'compl');
    setValue(bairroInput, 'bairro');
    setValue(cepInput, 'cep');
    setValue(cidadeInput, 'cidade');
    setValue(estInput, 'est');
    setValue(telInput, 'tel');
    setValue(celularInput, 'celular');
    setValue(orgaoFuncionalInput, 'orgao_funcional');
    setValue(matFuncionalInput, 'mat_funcional');
    setValue(funcaoInput, 'funcao');
    setValue(unidadeInput, 'unidade');

    // Benefícios Contratados
    setChecked(MensalidadeSocial, 'mensalidade_social');
    setChecked(planoSaudeAmb, 'plano_saude_amb');
    setChecked(PlanoSaudeComp, 'plano_saude_comp');
    setChecked(OrientacaoJuridica, 'orientacao_juridica');
    setChecked(PlanoOdonto, 'plano_odonto');
    setChecked(SeguroVida, 'seguro_vida');
    // Radio buttons Seguro de Vida
    if (SeguroVidaSim) SeguroVidaSim.checked = (data['seguro_vida_opcao'] === 'sim');
    if (SeguroVidaNao) SeguroVidaNao.checked = (data['seguro_vida_opcao'] === 'nao');
    setChecked(AuxilioNatalidade, 'auxilio_natalidade');
    setChecked(AssistenciaFuneral, 'assistencia_funeral');
    setChecked(Convenios, 'convenios');

    // Relação de Dependentes (União de Benefícios)
    for (let i = 1; i <= 8; i++) {
        setValue(document.getElementById(`dep${i}_nome`), `dep${i}_nome`);
        setValue(document.getElementById(`dep${i}_nasc`), `dep${i}_nasc`);
        setValue(document.getElementById(`dep${i}_parentesco`), `dep${i}_parentesco`);
        setChecked(document.getElementById(`dep${i}_plano_amb`), `dep${i}_plano_amb`);
        setChecked(document.getElementById(`dep${i}_plano_comp`), `dep${i}_plano_comp`);
        setChecked(document.getElementById(`dep${i}_plano_odonto`), `dep${i}_plano_odonto`);
        setChecked(document.getElementById(`dep${i}_assist_funeral`), `dep${i}_assist_funeral`);
    }

    // Termos e Declaração
    setValue(declarationRsValue, 'declaration_rs_value');
    setValue(declarationLocal, 'declaration_local');
    setValue(declarationDia, 'declaration_dia');
    setValue(declarationMes, 'declaration_mes');
    setValue(declarationAno, 'declaration_ano');
    setValue(agentName, 'agent_name');
    setValue(agentRegistro, 'agent_registro');

    // Proposta de Contratação - Dados do Vendedor
    setValue(contratacaoProposalNumber, 'contratacao_proposal_number');
    setValue(contratacaoVendedorCpf, 'contratacao_vendedor_cpf');
    setValue(contratacaoVendedorNome, 'contratacao_vendedor_nome');
    setValue(contratacaoVendedorTelefone, 'contratacao_vendedor_telefone');
    setValue(contratacaoCnpjCorretora, 'contratacao_cnpj_corretora');
    setValue(contratacaoNomeCorretora, 'contratacao_nome_corretora');

    // Proposta de Contratação - Dados cadastrais do Contratante
    setChecked(contratacaoContratanteResponsavel, 'contratacao_contratante_responsavel');
    setValue(contratacaoContratanteCpf, 'contratacao_contratante_cpf');
    setValue(contratacaoContratanteNome, 'contratacao_contratante_nome');
    setValue(contratacaoContratanteSexo, 'contratacao_contratante_sexo');
    setValue(contratacaoContratanteRg, 'contratacao_contratante_rg');
    setValue(contratacaoContratanteDataEmissaoRg, 'contratacao_contratante_data_emissao_rg');
    setValue(contratacaoContratanteOrgaoEmissor, 'contratacao_contratante_orgao_emissor');
    setValue(contratacaoContratanteEstadoCivil, 'contratacao_contratante_estado_civil');
    setValue(contratacaoContratanteDataNascimento, 'contratacao_contratante_data_nascimento');
    setValue(contratacaoContratanteTelefoneFixo, 'contratacao_contratante_telefone_fixo');
    setValue(contratacaoContratanteCelular, 'contratacao_contratante_celular');
    setValue(contratacaoContratanteWhatsapp, 'contratacao_contratante_whatsapp');
    setValue(contratacaoContratanteCartaoNacionalSaude, 'contratacao_contratante_cartao_nacional_saude');
    setValue(contratacaoContratanteEmail, 'contratacao_contratante_email');
    setValue(contratacaoContratanteNomeMae, 'contratacao_contratante_nome_mae');
    // Radio buttons: Contratante - Possui Plano Anterior?
    if (contratacaoContratantePlanoAnteriorSim) contratacaoContratantePlanoAnteriorSim.checked = (String(data['contratacao_contratante_plano_anterior']).toUpperCase() === 'SIM');
    if (contratacaoContratantePlanoAnteriorNao) contratacaoContratantePlanoAnteriorNao.checked = (String(data['contratacao_contratante_plano_anterior']).toUpperCase() === 'NÃO');
    setValue(contratacaoContratanteQualPlano, 'contratacao_contratante_qual_plano');
    setValue(contratacaoContratanteDataInicio, 'contratacao_contratante_data_inicio');
    setValue(contratacaoContratanteDataUltimoPagamento, 'contratacao_contratante_data_ultimo_pagamento');
    setValue(contratacaoContratanteCep, 'contratacao_contratante_cep');
    setValue(contratacaoContratanteLogradouro, 'contratacao_contratante_logradouro');
    setValue(contratacaoContratanteNumero, 'contratacao_contratante_numero');
    setValue(contratacaoContratanteComplemento, 'contratacao_contratante_complemento');
    setValue(contratacaoContratanteBairro, 'contratacao_contratante_bairro');
    setValue(contratacaoContratanteCidade, 'contratacao_contratante_cidade');
    setValue(contratacaoContratanteEstado, 'contratacao_contratante_estado');

    // Documentação entregue (Contratante)
    setChecked(docEntregueCpfRgCnh, 'doc_entregue_cpf_rg_cnh');
    setChecked(docEntregueFotoseIfie, 'doc_entregue_fotoseIfie');
    setChecked(docEntregueComprovanteResidencia, 'doc_entregue_comprovante_residencia');
    setChecked(docEntregueComprovanteVendaBeneficiario, 'doc_entregue_comprovante_venda_beneficiario');
    setChecked(docEntregueOutros, 'doc_entregue_outros');
    setChecked(docEntregueDeclaracaoTempoPermanencia, 'doc_entregue_declaracao_tempo_permanencia');

    // Proposta de Contratação - Dados cadastrais dos Dependentes (contratacao_depX)
    for (let i = 1; i <= 4; i++) {
        setValue(document.getElementById(`contratacao_dep${i}_cpf`), `contratacao_dep${i}_cpf`);
        setValue(document.getElementById(`contratacao_dep${i}_nome`), `contratacao_dep${i}_nome`);
        setValue(document.getElementById(`contratacao_dep${i}_sexo`), `contratacao_dep${i}_sexo`);
        setValue(document.getElementById(`contratacao_dep${i}_rg`), `contratacao_dep${i}_rg`);
        setValue(document.getElementById(`contratacao_dep${i}_data_emissao_rg`), `contratacao_dep${i}_data_emissao_rg`);
        setValue(document.getElementById(`contratacao_dep${i}_orgao_emissor`), `contratacao_dep${i}_orgao_emissor`);
        setValue(document.getElementById(`contratacao_dep${i}_estado_civil`), `contratacao_dep${i}_estado_civil`);
        setValue(document.getElementById(`contratacao_dep${i}_data_nascimento`), `contratacao_dep${i}_data_nascimento`);
        setValue(document.getElementById(`contratacao_dep${i}_cartao_nacional_saude`), `contratacao_dep${i}_cartao_nacional_saude`);
        setValue(document.getElementById(`contratacao_dep${i}_parentesco`), `contratacao_dep${i}_parentesco`);
        setValue(document.getElementById(`contratacao_dep${i}_nome_mae`), `contratacao_dep${i}_nome_mae`);

        // Radio buttons: Dependente X - Possui Plano Anterior?
        const depXPlanoAnteriorSim = document.getElementById(`contratacao_dep${i}_plano_anterior_sim`);
        const depXPlanoAnteriorNao = document.getElementById(`contratacao_dep${i}_plano_anterior_nao`);
        if (depXPlanoAnteriorSim) depXPlanoAnteriorSim.checked = (String(data[`contratacao_dep${i}_plano_anterior`]).toUpperCase() === 'SIM');
        if (depXPlanoAnteriorNao) depXPlanoAnteriorNao.checked = (String(data[`contratacao_dep${i}_plano_anterior`]).toUpperCase() === 'NÃO');

        setValue(document.getElementById(`contratacao_dep${i}_qual_plano`), `contratacao_dep${i}_qual_plano`);
        setValue(document.getElementById(`contratacao_dep${i}_data_inicio`), `contratacao_dep${i}_data_inicio`);
        setValue(document.getElementById(`contratacao_dep${i}_data_ultimo_pagamento`), `contratacao_dep${i}_data_ultimo_pagamento`);

        // Checkboxes: Dependente X - Documentação entregue
        setChecked(document.getElementById(`doc_entregue_dep${i}_cpf_rg_cnh`), `doc_entregue_dep${i}_cpf_rg_cnh`);
        if (i === 1) { // Apenas Dep1 tem todos os checkboxes no HTML fornecido
            setChecked(document.getElementById(`doc_entregue_dep${i}_fotoseIfie`), `doc_entregue_dep${i}_fotoseIfie`);
            setChecked(document.getElementById(`doc_entregue_dep${i}_comprovante_residencia`), `doc_entregue_dep${i}_comprovante_residencia`);
            setChecked(document.getElementById(`doc_entregue_dep${i}_comprovante_venda_beneficiario`), `doc_entregue_dep${i}_comprovante_venda_beneficiario`);
            setChecked(document.getElementById(`doc_entregue_dep${i}_outros`), `doc_entregue_dep${i}_outros`);
            setChecked(document.getElementById(`doc_entregue_dep${i}_declaracao_tempo_permanencia`), `doc_entregue_dep${i}_declaracao_tempo_permanencia`);
        }
        setValue(document.getElementById(`contratacao_dep${i}_qual_operadora_anterior`), `contratacao_dep${i}_qual_operadora_anterior`);
        setValue(document.getElementById(`contratacao_dep${i}_registro_produto_anterior`), `contratacao_dep${i}_registro_produto_anterior`);
    }

    // Responsável pelo Contrato
    setValue(contratacaoProposalNumberResponsavel, 'contratacao_proposal_number_responsavel');
    setValue(contratacaoResponsavelCpf, 'contratacao_responsavel_cpf');
    setValue(contratacaoResponsavelNome, 'contratacao_responsavel_nome');
    // Radio buttons: Responsável - Sexo
    if (contratacaoResponsavelSexoMasculino) contratacaoResponsavelSexoMasculino.checked = (data['contratacao_responsavel_sexo'] === 'Masculino');
    if (contratacaoResponsavelSexoFeminino) contratacaoResponsavelSexoFeminino.checked = (data['contratacao_responsavel_sexo'] === 'Feminino');
    setValue(contratacaoResponsavelEmail, 'contratacao_responsavel_email');
    setValue(contratacaoResponsavelNomeMae, 'contratacao_responsavel_nome_mae');
    setValue(contratacaoResponsavelCelular, 'contratacao_responsavel_celular');
    setValue(contratacaoResponsavelDataNascimento, 'contratacao_responsavel_data_nascimento');
    setValue(contratacaoResponsavelEstadoCivil, 'contratacao_responsavel_estado_civil');
    setValue(contratacaoResponsavelRgCnh, 'contratacao_responsavel_rg_cnh');
    setValue(contratacaoResponsavelDataEmissaoRgCnh, 'contratacao_responsavel_data_emissao_rg_cnh');
    setValue(contratacaoResponsavelCep, 'contratacao_responsavel_cep');
    setValue(contratacaoResponsavelLogradouro, 'contratacao_responsavel_logradouro');
    setValue(contratacaoResponsavelNumero, 'contratacao_responsavel_numero');
    setValue(contratacaoResponsavelComplemento, 'contratacao_responsavel_complemento');
    setValue(contratacaoResponsavelBairro, 'contratacao_responsavel_bairro');
    setValue(contratacaoResponsavelCidade, 'contratacao_responsavel_cidade');
    setValue(contratacaoResponsavelEstado, 'contratacao_responsavel_estado');

    // Documentação entregue (Responsável)
    setChecked(docEntregueRespCpfRgCnh, 'doc_entregue_resp_cpf_rg_cnh');
    setChecked(docEntregueRespComprovanteResidencia, 'doc_entregue_resp_comprovante_residencia');
    setChecked(docEntregueRespOutros, 'doc_entregue_resp_outros');

    // Resumo da Contratação
    setValue(contratacaoResumoDataProposta, 'contratacao_resumo_data_proposta');
    setValue(contratacaoResumoProvavelVigencia, 'contratacao_resumo_provavel_vigencia');
    setValue(contratacaoResumoCpfContratante, 'contratacao_resumo_cpf_contratante');
    setValue(contratacaoResumoTipo, 'contratacao_resumo_tipo');
    setValue(contratacaoResumoSegmentacao, 'contratacao_resumo_segmentacao');
    setValue(contratacaoResumoPlano, 'contratacao_resumo_plano');
    setValue(contratacaoResumoRegistroAns, 'contratacao_resumo_registro_ans');
    // Radio buttons: Acomodação
    if (acomodacaoQc) acomodacaoQc.checked = (data['contratacao_resumo_acomodacao'] === 'QC');
    if (acomodacaoQp) acomodacaoQp.checked = (data['contratacao_resumo_acomodacao'] === 'QP');
    setValue(contratacaoResumoAbrangencia, 'contratacao_resumo_abrangencia');
    // Radio buttons: Coparticipação
    if (coparticipacaoNao) coparticipacaoNao.checked = (data['contratacao_resumo_coparticipacao'] === 'NÃO');
    setValue(contratacaoResumoBeneficiarios, 'contratacao_resumo_beneficiarios');
    setValue(contratacaoResumoValorTotal, 'contratacao_resumo_valor_total');


    // Declaração de Saúde (Itens 1-21 e IMC)
    setValue(declaracaoSaudeProposalNumber, 'declaracao_saude_proposal_number');
    setValue(declaracaoSaudeProposalNumberCont, 'declaracao_saude_proposal_number_cont');
    for (let i = 1; i <= 21; i++) {
        setValue(window[`item${i}Titular`], `item${i}_titular`);
        setValue(window[`item${i}Dep1`], `item${i}_dep1`);
        setValue(window[`item${i}Dep2`], `item${i}_dep2`);
        setValue(window[`item${i}Dep3`], `item${i}_dep3`);
        setValue(window[`item${i}Dep4`], `item${i}_dep4`);
    }
    setValue(bmiTitularPeso, 'bmi_titular_peso');
    setValue(bmiDep1Peso, 'bmi_dep1_peso');
    setValue(bmiDep2Peso, 'bmi_dep2_peso');
    setValue(bmiDep3Peso, 'bmi_dep3_peso');
    setValue(bmiDep4Peso, 'bmi_dep4_peso');
    setValue(bmiTitularAltura, 'bmi_titular_altura');
    setValue(bmiDep1Altura, 'bmi_dep1_altura');
    setValue(bmiDep2Altura, 'bmi_dep2_altura');
    setValue(bmiDep3Altura, 'bmi_dep3_altura');
    setValue(bmiDep4Altura, 'bmi_dep4_altura');

    // Informações Complementares (comp_info_item_num_1 a comp_info_year_11)
    setValue(complementaryInfoProposalNumber, 'complementary_info_proposal_number');
    for (let i = 1; i <= 11; i++) {
        setValue(window[`compInfoItemNum${i}`], `comp_info_item_num_${i}`);
        setValue(window[`compInfoDescription${i}`], `comp_info_description_${i}`);
        setChecked(window[`compInfo${i}Titular`], `comp_info_${i}_titular`);
        setChecked(window[`compInfo${i}Dep1`], `comp_info_${i}_dep1`);
        setChecked(window[`compInfo${i}Dep2`], `comp_info_${i}_dep2`);
        setChecked(window[`compInfo${i}Dep3`], `comp_info_${i}_dep3`);
        setChecked(window[`compInfo${i}Dep4`], `comp_info_${i}_dep4`);
        setValue(window[`compInfoYear${i}`], `comp_info_year_${i}`);
    }

    // Declaração de Saúde Final
    setValue(declaracaoSaudeFinalProposalNumber, 'declaracao_saude_final_proposal_number');
    setChecked(entrevistaQualificadaOpcao1Chk, 'entrevista_qualificada_opcao1_chk');
    setChecked(entrevistaQualificadaOpcao2Chk, 'entrevista_qualificada_opcao2_chk');
    setChecked(entrevistaQualificadaOpcao3Chk, 'entrevista_qualificada_opcao3_chk');

    // Carta de Orientação ao Beneficiário - Assinaturas
    setValue(cartaBeneficiarioNomeSig, 'carta_beneficiario_nome_sig');
    setValue(cartaBeneficiarioLocalSig, 'carta_beneficiario_local_sig');
    setValue(cartaBeneficiarioDataSig, 'carta_beneficiario_data_sig');
    setValue(cartaIntermediarioNomeSig, 'carta_intermediario_nome_sig');
    setValue(cartaIntermediarioCpfSig, 'carta_intermediario_cpf_sig');
    setValue(cartaIntermediarioLocalSig, 'carta_intermediario_local_sig');
    setValue(cartaIntermediarioDataSig, 'carta_intermediario_data_sig');

    // Termo Único de Promoções
    setValue(termoPromocoesProposalNumber, 'termo_promocoes_proposal_number');

    // Declaração de Recebimento e Posse
    setValue(declaracaoRecebimentoProposalNumber, 'declaracao_recebimento_proposal_number');

    // Termo de Consentimento
    setValue(termoConsentimentoProposalNumber, 'termo_consentimento_proposal_number');

    // Termo de Adesão Planos Odontológicos - Titular
    setValue(odontoTitularNome, 'odonto_titular_nome');
    setValue(odontoTitularCpf, 'odonto_titular_cpf');
    setValue(odontoTitularDataNascimento, 'odonto_titular_data_nascimento');
    setValue(odontoTitularRg, 'odonto_titular_rg');
    setValue(odontoTitularSexo, 'odonto_titular_sexo');
    setValue(odontoTitularEstadoCivil, 'odonto_titular_estado_civil');
    setValue(odontoTitularEndereco, 'odonto_titular_endereco');
    setValue(odontoTitularNumero, 'odonto_titular_numero');
    setValue(odontoTitularBairro, 'odonto_titular_bairro');
    setValue(odontoTitularCep, 'odonto_titular_cep');
    setValue(odontoTitularCidade, 'odonto_titular_cidade');
    setValue(odontoTitularTelefone, 'odonto_titular_telefone');
    setValue(odontoTitularMae, 'odonto_titular_mae');
    setValue(odontoTitularEmail, 'odonto_titular_email');

    // Termo de Adesão Planos Odontológicos - Dependentes 1-4
    for (let i = 1; i <= 4; i++) {
        setValue(document.getElementById(`odonto_dep${i}_nome`), `odonto_dep${i}_nome`);
        setValue(document.getElementById(`odonto_dep${i}_cpf`), `odonto_dep${i}_cpf`);
        setValue(document.getElementById(`odonto_dep${i}_data_nascimento`), `odonto_dep${i}_data_nascimento`);
        setValue(document.getElementById(`odonto_dep${i}_sexo`), `odonto_dep${i}_sexo`);
        setValue(document.getElementById(`odonto_dep${i}_estado_civil`), `odonto_dep${i}_estado_civil`);
        setValue(document.getElementById(`odonto_dep${i}_mae`), `odonto_dep${i}_mae`);
        setValue(document.getElementById(`odonto_dep${i}_parentesco`), `odonto_dep${i}_parentesco`);
    }

    // Autorização para Desconto
    setValue(nomeDescontoAutoriza, 'nomeDescontoAutoriza');
    setValue(cpfDescontoAutoriza, 'cpfDescontoAutoriza');
    setValue(matFuncionalDescontoAutoriza, 'matFuncionalDescontoAutoriza');
    setValue(identidadeDescontoAutoriza, 'identidadeDescontoAutoriza');
    setValue(diaDescontoAutoriza, 'diaDescontoAutoriza');
    setValue(mesDescontoAutoriza, 'mesDescontoAutoriza');
    setValue(anosDescontoAutoriza, 'anosDescontoAutoriza');
    setValue(valorDescontoAutoriza, 'valorDescontoAutoriza');
    setValue(totalDescontoAutoriza, 'totalDescontoAutoriza');
    setChecked(mensalidadeSocialDescontoAutoriza, 'mensalidadeSocialDescontoAutoriza');
    setChecked(planoSaudeAmbulatorialDescontoAutoriza, 'planoSaudeAmbulatorialDescontoAutoriza');
    setChecked(orientacaoJuridicaDescontoAutoriza, 'orientacaoJuridicaDescontoAutoriza');
    setChecked(segurosDescontoAutoriza, 'segurosDescontoAutoriza');
    setChecked(auxilioNatalidadeDescontoAutoriza, 'auxilioNatalidadeDescontoAutoriza');
    setChecked(planoSaudeCompletoDescontoAutoriza, 'planoSaudeCompletoDescontoAutoriza');
    setChecked(atdDomiciliarDescontoAutoriza, 'atdDomiciliarDescontoAutoriza');
    setChecked(assistenciaFuneralDescontoAutoriza, 'assistenciaFuneralDescontoAutoriza');
    setChecked(planoOdontologicoDescontoAutoriza, 'planoOdontologicoDescontoAutoriza');
    setChecked(conveniosDescontoAutoriza, 'conveniosDescontoAutoriza');
    setValue(bancoDescontoAutoriza, 'bancoDescontoAutoriza');
    setValue(agenciaDescontoAutoriza, 'agenciaDescontoAutoriza');
    setValue(contaCorrenteDescontoAutoriza, 'contaCorrenteDescontoAutoriza');
    setValue(valorCobrancaDescontoAutoriza, 'valorCobrancaDescontoAutoriza');
    setValue(despesaBancoDescontoAutoriza, 'despesaBancoDescontoAutoriza');
    setValue(valorTotalDescontoAutoriza, 'valorTotalDescontoAutoriza');
    setValue(dataAssDescontoAutoriza, 'dataAssDescontoAutoriza');
    setValue(assinaturaDescontoAutoriza, 'assinaturaDescontoAutoriza');

    // Lógica para preenchimento de campos de data importados do Excel
    const dateInputs = document.querySelectorAll('input[type="date"]');
    dateInputs.forEach(input => {
        const key = input.id || input.name;
        if (data.hasOwnProperty(key)) {
            const dateValue = data[key];
            if (typeof dateValue === 'string' && dateValue) {
                // Tenta converter de dd/mm/aaaa para YYYY-MM-DD
                let parts = dateValue.split('/');
                if (parts.length === 3) {
                    input.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                } else {
                    // Assume que já está em YYYY-MM-DD ou outro formato compatível
                    input.value = dateValue;
                }
            } else if (dateValue instanceof Date && !isNaN(dateValue)) {
                // Se for um objeto Date válido (do XLSX), formata para YYYY-MM-DD
                const year = dateValue.getFullYear();
                const month = (dateValue.getMonth() + 1).toString().padStart(2, '0');
                const day = dateValue.getDate().toString().padStart(2, '0');
                input.value = `${year}-${month}-${day}`;
            } else {
                input.value = ''; // Limpa o campo se o valor não for válido
            }
        } else {
            input.value = ''; // Limpa se não houver dado
        }
    });


    // Dispara blur para validação de CPF após preenchimento
    if (cpfInput) {
        const blurEvent = new Event('blur');
        cpfInput.dispatchEvent(blurEvent);
    }
}


// --- CPF Validation Function ---
function validateCPF(cpf) {
    cpf = String(cpf).replace(/[^\d]/g, '');

    if (cpf.length !== 11) return false;
    if (/^(\d)\1+$/.test(cpf)) return false;

    let sum = 0;
    let remainder;

    for (let i = 1; i <= 9; i++) sum = sum + parseInt(cpf.substring(i - 1, i)) * (11 - i);
    remainder = (sum * 10) % 11;
    if ((remainder === 10) || (remainder === 11)) remainder = 0;
    if (remainder !== parseInt(cpf.substring(9, 10))) return false;

    sum = 0;
    for (let i = 1; i <= 10; i++) sum = sum + parseInt(cpf.substring(i - 1, i)) * (12 - i);
    remainder = (sum * 10) % 11;
    if ((remainder === 10) || (remainder === 11)) remainder = 0;
    if (remainder !== parseInt(cpf.substring(10, 11))) return false;

    return true;
}

// --- Event Listener for "Gerar Arquivo" ---
generateBtn.addEventListener('click', () => {
    try {
        if (!nomeInput.value.trim()) {
            alert("Por Favor, informe o nome antes de gerar o arquivo.");
            nomeInput.focus();
            nomeInput.classList.add('invalid');
            return;
        }

        const cpfValue = cpfInput.value;
        if (!validateCPF(cpfValue)) {
            alert("Por favor, insira um CPF válido antes de gerar o arquivo.");
            cpfInput.focus();
            cpfInput.classList.add('invalid');
            return;
        } else {
            cpfInput.classList.remove('invalid');
        }

        const nomeFile = nomeInput.value.trim();
        const defaultFileName = `${nomeFile.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`; // Clean filename

        const formData = getFormData();

        const excelData = {};
        for (const key in formData) {
            if (formData.hasOwnProperty(key)) {
                if (typeof formData[key] === 'boolean') {
                    excelData[key] = formData[key] ? 'Sim' : 'Não';
                } else if (formData[key] instanceof Date) {
                    // Formata objetos Date para string YYYY-MM-DD para Excel
                    const date = formData[key];
                    excelData[key] = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
                } else {
                    excelData[key] = formData[key];
                }
            }
        }

        const ws = XLSX.utils.json_to_sheet([excelData]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "CadastroSocio");

        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: "application/octet-stream" });

        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", defaultFileName);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);

        link.click();

        document.body.removeChild(link);
        URL.revokeObjectURL(url);

    } catch (error) {
        console.error("Erro ao gerar o arquivo Excel:", error);
        alert("Ocorreu um erro ao tentar gerar o arquivo Excel.");
    }
});

readBtn.addEventListener('click', () => {
    fileInput.click();
});

// --- Event Listener for File Input Change ---
fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) {
        return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            // cellDates: true permite que o SheetJS tente converter strings de data para objetos Date.
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

            if (jsonData.length > 0) {
                const loadedData = jsonData[0];
                const processedData = {};
                for (const key in loadedData) {
                    if (loadedData.hasOwnProperty(key)) {
                        const value = loadedData[key];
                        if (typeof value === 'string') {
                            if (value.toUpperCase().trim() === 'SIM') {
                                processedData[key] = true;
                            } else if (value.toUpperCase().trim() === 'NÃO') {
                                processedData[key] = false;
                            } else {
                                processedData[key] = value.trim();
                            }
                        } else {
                            processedData[key] = value;
                        }
                    }
                }
                setFormData(processedData);
                alert("Dados carregados com sucesso!");
            } else {
                alert("A planilha selecionada está vazia.");
            }
        } catch (error) {
            console.error("Erro ao ler o arquivo:", error);
            alert("Erro ao ler o arquivo. Verifique se é um arquivo Excel válido (.xlsx ou .xls).");
        } finally {
            fileInput.value = null;
        }
    };

    reader.onerror = (e) => {
        console.error("Erro ao ler o arquivo:", e);
        alert("Ocorreu um erro ao tentar ler o arquivo.");
        fileInput.value = null;
    };

    reader.readAsArrayBuffer(file);
});


// Initial update on page load in case fields have pre-filled values
document.addEventListener('DOMContentLoaded', () => {
    updateContratanteFields();
    updateContratacaoResponsavelFields();
    updateAllDependentFields(); // Chamada inicial para preencher os dependentes
    setContratacaoResumoDate();
});

sendEmailBtn.addEventListener('click', () => {
    console.log("Send Email button clicked. Server-side implementation needed.");
    alert("Funcionalidade de envio de e-mail não implementada neste ambiente.");
});
