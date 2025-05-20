import * as XLSX from 'xlsx';

const generateBtn = document.getElementById('generate-btn');
const readBtn = document.getElementById('read-btn');
const fileInput = document.getElementById('file-input');
const cpfInput = document.getElementById('cpf');
const sendEmailBtn = document.getElementById('send-email-btn');

// Get source fields from the first section (Titular)
const nomeInput = document.getElementById('nome');
const nascInput = document.getElementById('nasc');
const sexoInput = document.getElementById('sexo');
const estCivilInput = document.getElementById('est_civil');
const rgInput = document.getElementById('rg');
const expInput = document.getElementById('exp');
const celularInput = document.getElementById('celular');
const emailInput = document.getElementById('email');
const maeInput = document.getElementById('mae');
const cepInput = document.getElementById('cep');
const endInput = document.getElementById('end');
const numInput = document.getElementById('num');
const complInput = document.getElementById('compl');
const bairroInput = document.getElementById('bairro');
const cidadeInput = document.getElementById('cidade');
const estInput = document.getElementById('est');
const telInput = document.getElementById('tel');

// Get source fields from the dependent rows (União de Benefícios section)
const dep1NomeInput = document.getElementById('dep1_nome');
const dep1NascInput = document.getElementById('dep1_nasc');
const dep1ParentescoInput = document.getElementById('dep1_parentesco');
const dep2NomeInput = document.getElementById('dep2_nome');
const dep2NascInput = document.getElementById('dep2_nasc');
const dep2ParentescoInput = document.getElementById('dep2_parentesco');
const dep3NomeInput = document.getElementById('dep3_nome');
const dep3NascInput = document.getElementById('dep3_nasc');
const dep3ParentescoInput = document.getElementById('dep3_parentesco');
const dep4NomeInput = document.getElementById('dep4_nome');
const dep4NascInput = document.getElementById('dep4_nasc');
const dep4ParentescoInput = document.getElementById('dep4_parentesco');

// Get target fields from the last section (Termo de Adesão Planos Odontológicos - Titular)
const odontoTitularNomeInput = document.getElementById('odonto_titular_nome');
const odontoTitularCpfInput = document.getElementById('odonto_titular_cpf');
const odontoTitularDataNascimentoInput = document.getElementById('odonto_titular_data_nascimento');
const odontoTitularRgInput = document.getElementById('odonto_titular_rg');
const odontoTitularEstadoCivilSelect = document.getElementById('odonto_titular_estado_civil');
const odontoTitularSexoSelect = document.getElementById('odonto_titular_sexo');
const odontoTitularMaeInput = document.getElementById('odonto_titular_mae');
const odontoTitularEmailInput = document.getElementById('odonto_titular_email');
const odontoTitularEnderecoInput = document.getElementById('odonto_titular_endereco');
const odontoTitularNumeroInput = document.getElementById('odonto_titular_numero');
const odontoTitularBairroInput = document.getElementById('odonto_titular_bairro');
const odontoTitularCepInput = document.getElementById('odonto_titular_cep');
const odontoTitularCidadeInput = document.getElementById('odonto_titular_cidade');
const odontoTitularTelefoneInput = document.getElementById('odonto_titular_telefone');

// Get target fields from the "Proposta de Contratação" section (Contratante)
const contratacaoContratanteResponsavelCheckbox = document.getElementById('contratacao_contratante_responsavel');
const contratacaoContratanteCpfInput = document.getElementById('contratacao_contratante_cpf');
const contratacaoContratanteNomeInput = document.getElementById('contratacao_contratante_nome');
const contratacaoContratanteSexoSelect = document.getElementById('contratacao_contratante_sexo');
const contratacaoContratanteRgInput = document.getElementById('contratacao_contratante_rg');
const contratacaoContratanteDataEmissaoRgInput = document.getElementById('contratacao_contratante_data_emissao_rg');
const contratacaoContratanteOrgaoEmissorInput = document.getElementById('contratacao_contratante_orgao_emissor');
const contratacaoContratanteEstadoCivilSelect = document.getElementById('contratacao_contratante_estado_civil');
const contratacaoContratanteDataNascimentoInput = document.getElementById('contratacao_contratante_data_nascimento');
const contratacaoContratanteTelefoneFixoInput = document.getElementById('contratacao_contratante_telefone_fixo');
const contratacaoContratanteCelularInput = document.getElementById('contratacao_contratante_celular');
const contratacaoContratanteWhatsappInput = document.getElementById('contratacao_contratante_whatsapp');
const contratacaoContratanteCartaoNacionalSaudeInput = document.getElementById('contratacao_contratante_cartao_nacional_saude');
const contratacaoContratanteEmailInput = document.getElementById('contratacao_contratante_email');
const contratacaoContratanteNomeMaeInput = document.getElementById('contratacao_contratante_nome_mae');
const contratacaoContratantePlanoAnteriorRadios = document.querySelectorAll('input[name="contratacao_contratante_plano_anterior"]');
const contratacaoContratanteQualPlanoInput = document.getElementById('contratacao_contratante_qual_plano');
const contratacaoContratanteDataInicioInput = document.getElementById('contratacao_contratante_data_inicio');
const contratacaoContratanteDataUltimoPagamentoInput = document.getElementById('contratacao_contratante_data_ultimo_pagamento');
const contratacaoContratanteCepInput = document.getElementById('contratacao_contratante_cep');
const contratacaoContratanteLogradouroInput = document.getElementById('contratacao_contratante_logradouro');
const contratacaoContratanteNumeroInput = document.getElementById('contratacao_contratante_numero');
const contratacaoContratanteComplementoInput = document.getElementById('contratacao_contratante_complemento');
const contratacaoContratanteBairroInput = document.getElementById('contratacao_contratante_bairro');
const contratacaoContratanteCidadeInput = document.getElementById('contratacao_contratante_cidade');
const contratacaoContratanteEstadoInput = document.getElementById('contratacao_contratante_estado');

// Get target fields from the "Proposta de Contratação" section (Dependents)
const contratacaoDep1NomeInput = document.getElementById('contratacao_dep1_nome');
const contratacaoDep1DataNascimentoInput = document.getElementById('contratacao_dep1_data_nascimento');
const contratacaoDep1ParentescoInput = document.getElementById('contratacao_dep1_parentesco');
const contratacaoDep2NomeInput = document.getElementById('contratacao_dep2_nome');
const contratacaoDep2DataNascimentoInput = document.getElementById('contratacao_dep2_data_nascimento');
const contratacaoDep2ParentescoInput = document.getElementById('contratacao_dep2_parentesco');
const contratacaoDep3NomeInput = document.getElementById('contratacao_dep3_nome');
const contratacaoDep3DataNascimentoInput = document.getElementById('contratacao_dep3_data_nascimento');
const contratacaoDep3ParentescoInput = document.getElementById('contratacao_dep3_parentesco');
const contratacaoDep4NomeInput = document.getElementById('contratacao_dep4_nome');
const contratacaoDep4DataNascimentoInput = document.getElementById('contratacao_dep4_data_nascimento');
const contratacaoDep4ParentescoInput = document.getElementById('contratacao_dep4_parentesco');

// Get target fields from the "Proposta de Contratação - Responsável pelo Contrato" section
const contratacaoResponsavelCpfInput = document.getElementById('contratacao_responsavel_cpf');
const contratacaoResponsavelNomeInput = document.getElementById('contratacao_responsavel_nome');
const contratacaoResponsavelEmailInput = document.getElementById('contratacao_responsavel_email');
const contratacaoResponsavelNomeMaeInput = document.getElementById('contratacao_responsavel_nome_mae');
const contratacaoResponsavelCelularInput = document.getElementById('contratacao_responsavel_celular');
const contratacaoResponsavelDataNascimentoInput = document.getElementById('contratacao_responsavel_data_nascimento');
const contratacaoResponsavelEstadoCivilSelect = document.getElementById('contratacao_responsavel_estado_civil');
const contratacaoResponsavelRgCnhInput = document.getElementById('contratacao_responsavel_rg_cnh');
const contratacaoResponsavelCepInput = document.getElementById('contratacao_responsavel_cep');
const contratacaoResponsavelLogradouroInput = document.getElementById('contratacao_responsavel_logradouro');
const contratacaoResponsavelNumeroInput = document.getElementById('contratacao_responsavel_numero');
const contratacaoResponsavelComplementoInput = document.getElementById('contratacao_responsavel_complemento');
const contratacaoResponsavelBairroInput = document.getElementById('contratacao_responsavel_bairro');
const contratacaoResponsavelCidadeInput = document.getElementById('contratacao_responsavel_cidade');
const contratacaoResponsavelEstadoInput = document.getElementById('contratacao_responsavel_estado');

// Get target Sexo radio buttons from the "Responsável pelo Contrato" section
const contratacaoResponsavelSexoMasculinoRadio = document.getElementById('contratacao_responsavel_sexo_masculino');
const contratacaoResponsavelSexoFemininoRadio = document.getElementById('contratacao_responsavel_sexo_feminino');

// Get target fields from the "Resumo da contratação" section
const contratacaoResumoCpfContratanteInput = document.getElementById('contratacao_resumo_cpf_contratante');
const contratacaoResumoDataPropostaInput = document.getElementById('contratacao_resumo_data_proposta');


const pesoTitularDeclaracao = document.getElementById("bmi_titular_peso");
const pesoDep1Declaracao = document.getElementById("bmi_dep1_peso");
const pesoDep2Declaracao = document.getElementById("bmi_dep2_peso");
const pesoDep3Declaracao = document.getElementById("bmi_dep3_peso");
const pesoDep4Declaracao = document.getElementById("bmi_dep4_peso");

const alturaTitularDeclaracao = document.getElementById("bmi_titular_altura");
const alturaDep1Declaracao = document.getElementById("bmi_dep1_altura");
const alturaDep2Declaracao = document.getElementById("bmi_dep2_altura");
const alturaDep3Declaracao = document.getElementById("bmi_dep3_altura");
const alturaDep4Declaracao = document.getElementById("bmi_dep4_altura");

//Termo adesão odonto
const dadosNomeTitularTermo = document.getElementById("nomeTitularTermoOdonto");
const dadosDataNascTitularTermo = document.getElementById("dataNascTitularTermoOdonto");
const dadosEnderecoTitularTermo = document.getElementById("endTitularTermoOdonto");
const dadosNumberTitulatTermo = document.getElementById("numberTitularTermoOdonto");
const dadosComplementoTitularTermo = document.getElementById("complTitularTermoOdonto");
const dadosBairroTitularTermo = document.getElementById("bairroTitularTermoOdonto");
const dadosCepTitularTermo = document.getElementById("cepTitularTermoOdonto");
const dadosCidadeTitularTermo = document.getElementById("cidadeTitularTermoOdonto");
const dadosEstadoTitularTermo = document.getElementById("estadoTitularTermoOdonto");
const dadosEmailTitularTermo = document.getElementById("emailTitularTermoOdonto");
const dadosTelefoneTitularTermo = document.getElementById("telefoneTitularTermoOdonto");
const dadosTipoDocumentoTitularTermoOdonto = document.getElementById("tipoDocumentoTitularTermoOdonto");
const dadosNumeroDocumentoTitularTermoOdonto = document.getElementById("numeroDocumentoTitularTermoOdonto");
const dadosNacionalidadeTitularTermoOdonto = document.getElementById("nacionalidadeTitularTermoOdonto");
const dadosPaisEmissaoTitularTermo = document.getElementById("paisEmissaoTitularTermoOdonto");
const dadosResidenciaFiscalTitularTermo = document.getElementById("residenciaFiscalTitularTermoOdonto");
const dadosLocasNascimentoTitularTermo = document.getElementById("localNascimentoTitularTermoOdonto");
const dadosProfissaoTitularTermoOdonto = document.getElementById("profissaoTitularTermoOdonto");
const dadosDetalheOcupacaoTitularTermo = document.getElementById("detalhesOcupacaoTitularTermoOdonto");
const dadosRendaMediaMensalTitularTermo = document.getElementById("rendaMediaTitularTermoOdonto");


const dadosNomeTitularProposta = document.getElementById("nomeTitularProposta");
const dadosdataNascTitularProposta = document.getElementById("datanascTitularProposta");
const dadosEstadoCivilTitularProposta = document.getElementById("estadocivilTitularProposta");
const dadosEndTitularProposta = document.getElementById("endTitularProposta");
const dadosBairroTitularProposta = document.getElementById("bairroTitularProposta");
const dadosCidadeTitularProposta = document.getElementById("cidadeTitularProposta");
const dadosCepTitularProposta = document.getElementById("cepTitularProposta");
const dadosTelTitularProposta = document.getElementById("telTitularProposta");
const dadosNacionalidadeTitularProposta = document.getElementById("nacionalidadeTitularProposta");
const dadosManutencaoTitularProposta = document.getElementById("manutencaoTitularProposta");
const dadosUopTitularProposta = document.getElementById("uopTitularProposta");
const dadosCiaTitularProposta = document.getElementById("ciaTitularProposta");
const dadosSucursalTitularProposta = document.getElementById("sucursalTitularProposta");
const dadosRamoTitularProposta = document.getElementById("ramoTitularProposta");
const dadosApoliceTitularProposta = document.getElementById("apoliceTitularProposta");
const dadosNumeroCertificadoTitularProposta = document.getElementById("numeroCertificadoTitularProposta");
const dadosGrupoTitularProposta = document.getElementById("grupoTitularProposta");
const dadosPlanoTitularProposta = document.getElementById("planoTitularProposta");
const dadosProLaboreTitularProposta = document.getElementById("proLaboreTitularProposta");
const dadosEstipulanteTitularProposta = document.getElementById("estipulanteTitularProposta");
const dadosEstruturaVendaTitularProposta = document.getElementById("estruturaVendaTitularProposta");
const dadosPesoTitularProposta = document.getElementById("pesoTitularProposta");
const dadosAlturaTitularProposta = document.getElementById("alturaTitularProposta");
const dadosCargoTitularProposta = document.getElementById("cargoTitularProposta");
const dadosRendaMensalTitularProposta = document.getElementById("rendaMensalTitularProposta");
const dadosDataAdmissaoTitularProposta = document.getElementById("dataAdmissaoTitularProposta");
const dadosInicioVigenciaTitularProposta = document.getElementById("inicioVigenciaTitularProposta");
const dadosTerminoVigenciaTitularProposta = document.getElementById("terminioVigenciaTitularProposta");
const dadosNomeResponsavelTitularProposta = document.getElementById("nomeResponsavelTitularProposta");
const dadosCpfResponsavelTitularProposta = document.getElementById("cpfResponsavelTitularProposta");
const dadosCusteioTitularProposta = document.getElementById("custeioTitularProposta");
const dadosEmpresaTitularProposta = document.getElementById("empresaTitularProoposta");
const dadosFuncionarioTitularProposta = document.getElementById("FuncionarioTitularProposta");
const dadosSexoTitularProposta = document.getElementById("sexoTitularProposta");

const docTitularTermo = document.getElementById("numeroDocumentoTitularTermo"); 




// Function to update odonto fields based on titular data
function updateOdontoTitularFields() {

    if(contratacaoDep1NomeInput) contratacaoDep1NomeInput.value = dep1NomeInput ? dep1NomeInput.value : '';

    if(contratacaoDep1DataNascimentoInput) contratacaoDep1DataNascimentoInput.value = dep1NascInput ? dep1NascInput.value : '';

    if(contratacaoDep1ParentescoInput) contratacaoDep1ParentescoInput.value = dep1ParentescoInput ? dep1ParentescoInput.value : '';

    if(contratacaoDep2NomeInput) contratacaoDep2NomeInput.value = dep2NomeInput ? dep2NomeInput.value : '';

    if(contratacaoDep2DataNascimentoInput) contratacaoDep2DataNascimentoInput.value = dep2NascInput ? dep2NascInput.value : '';

    if(contratacaoDep2ParentescoInput) contratacaoDep2ParentescoInput.value = dep2ParentescoInput ? dep2ParentescoInput.value : '';

    if(contratacaoDep3NomeInput) contratacaoDep3NomeInput.value = dep3NomeInput ? dep3NomeInput.value : '';

    if(contratacaoDep3DataNascimentoInput) contratacaoDep3DataNascimentoInput.value = dep3NascInput ? dep3NascInput.value : '';

    if(contratacaoDep3ParentescoInput) contratacaoDep3ParentescoInput.value = dep3ParentescoInput ? dep3ParentescoInput.value : '';

    if(contratacaoDep4NomeInput) contratacaoDep4NomeInput.value = dep4NomeInput ? dep4NomeInput.value : '';

    if(contratacaoDep4DataNascimentoInput) contratacaoDep4DataNascimentoInput.value = dep4NascInput ? dep4NascInput.value : '';

    if(contratacaoDep4ParentescoInput) contratacaoDep4ParentescoInput.value = dep4ParentescoInput ? dep4ParentescoInput.value : '';

    if(dadosPesoTitularProposta) dadosPesoTitularProposta.value = pesoTitularDeclaracao ? pesoTitularDeclaracao.value : '';


    if(dadosAlturaTitularProposta) dadosAlturaTitularProposta.value = alturaTitularDeclaracao ? alturaTitularDeclaracao.value : '';

    if(dadosNumeroDocumentoTitularTermoOdonto)
    {
        if(dadosTipoDocumentoTitularTermoOdonto.value == 'cpf')
            {
                dadosNumeroDocumentoTitularTermoOdonto.value = cpfInput ? cpfInput.value : '';
            } 

    }

    if(docTitularTermo) docTitularTermo.value = dadosNumeroDocumentoTitularTermoOdonto ? dadosNumeroDocumentoTitularTermoOdonto.value : '';

    if(dadosNacionalidadeTitularProposta) dadosNacionalidadeTitularProposta.value = dadosNacionalidadeTitularTermoOdonto ? dadosNacionalidadeTitularTermoOdonto.value : '';


    if(dadosNomeTitularTermo) dadosNomeTitularTermo.value = nomeInput ? nomeInput.value : '';
    if(dadosDataNascTitularTermo) dadosDataNascTitularTermo.value = nascInput ? nascInput.value : '';
    if(dadosEnderecoTitularTermo) dadosEnderecoTitularTermo.value = endInput ? endInput.value : '';
    if(dadosNumberTitulatTermo) dadosNumberTitulatTermo.value =  numInput ? numInput.value : '';
    if(dadosComplementoTitularTermo) dadosComplementoTitularTermo.value = complInput ? complInput.value : '';
    if(dadosBairroTitularTermo) dadosBairroTitularTermo.value = bairroInput ? bairroInput.value : '';
    if(dadosCepTitularTermo) dadosCepTitularTermo.value = cepInput ? cepInput.value : '';
    if(dadosCidadeTitularTermo) dadosCidadeTitularTermo.value = cidadeInput ? cidadeInput.value : '';
    if(dadosEstadoTitularTermo) dadosEstadoTitularTermo.value = estInput ? estInput.value : '';
    if(dadosEmailTitularTermo) dadosEmailTitularTermo.value = emailInput ? emailInput.value : '';
    if(dadosTelefoneTitularTermo) dadosTelefoneTitularTermo.value = celularInput ? celularInput.value : '';

    if(dadosNomeTitularProposta) dadosNomeTitularProposta.value = nomeInput ? nomeInput.value : '';
    if(dadosdataNascTitularProposta) dadosdataNascTitularProposta.value = nascInput ? nascInput.value : '';
    if(dadosEstadoCivilTitularProposta) dadosEstadoCivilTitularProposta.value = estCivilInput ? estCivilInput.value : '';

    if(dadosEndTitularProposta) dadosEndTitularProposta.value = endInput ? endInput.value : '';
    if(dadosBairroTitularProposta) dadosBairroTitularProposta.value = bairroInput ? bairroInput.value : '';
    if(dadosCidadeTitularProposta) dadosCidadeTitularProposta.value = cidadeInput ? cidadeInput.value: '';
    if(dadosCepTitularProposta) dadosCepTitularProposta.value = cepInput ? cepInput.value : '';
    if(dadosTelTitularProposta) dadosTelTitularProposta.value = celularInput ? celularInput.value : '';
 


    if (odontoTitularNomeInput) odontoTitularNomeInput.value = nomeInput ? nomeInput.value : '';
    if (odontoTitularCpfInput) odontoTitularCpfInput.value = cpfInput ? cpfInput.value : '';
    if (odontoTitularDataNascimentoInput) odontoTitularDataNascimentoInput.value = nascInput ? nascInput.value : '';
    if (odontoTitularRgInput) odontoTitularRgInput.value = rgInput ? rgInput.value : '';
    if (odontoTitularEstadoCivilSelect) odontoTitularEstadoCivilSelect.value = estCivilInput ? estCivilInput.value : '';

    // Update Sexo select
    if (sexoInput && odontoTitularSexoSelect) {
        const sourceValue = sexoInput.value;
        const optionExists = Array.from(odontoTitularSexoSelect.options).some(option => option.value === sourceValue);
        odontoTitularSexoSelect.value = optionExists ? sourceValue : '';
        if (sourceValue !== '' && !optionExists) {
            console.warn(`Sex value "${sourceValue}" from source field is not a valid option in the target select.`);
        }
    }

    // Update new fields
    if (odontoTitularMaeInput) odontoTitularMaeInput.value = maeInput ? maeInput.value : '';
    if (odontoTitularEmailInput) odontoTitularEmailInput.value = emailInput ? emailInput.value : '';
    if (odontoTitularEnderecoInput) odontoTitularEnderecoInput.value = endInput ? endInput.value : '';
    // Mapping Num and Compl to one field (odonto_titular_numero) - Concatenate Num and Compl
    if (odontoTitularNumeroInput) {
        const numValue = numInput ? numInput.value : '';
        const complValue = complInput ? complInput.value : '';
        if (numValue && complValue) {
            odontoTitularNumeroInput.value = `${numValue}, ${complValue}`;
        } else if (numValue) {
            odontoTitularNumeroInput.value = numValue;
        } else if (complValue) {
            odontoTitularNumeroInput.value = complValue;
        } else {
            odontoTitularNumeroInput.value = '';
        }
    }
    if (odontoTitularBairroInput) odontoTitularBairroInput.value = bairroInput ? bairroInput.value : '';
    if (odontoTitularCepInput) odontoTitularCepInput.value = cepInput ? cepInput.value : '';
    if (odontoTitularCidadeInput) odontoTitularCidadeInput.value = cidadeInput ? cidadeInput.value : '';
    // Mapping Celular to Telefone in Odonto section
    if (odontoTitularTelefoneInput) odontoTitularTelefoneInput.value = celularInput ? celularInput.value : '';
}

// Function to update Contratante fields based on titular data
function updateContratanteFields() {
    // Pre-existing mappings
    if (contratacaoContratanteCpfInput) contratacaoContratanteCpfInput.value = cpfInput ? cpfInput.value : '';
    if (contratacaoContratanteNomeInput) contratacaoContratanteNomeInput.value = nomeInput ? nomeInput.value : '';
    if (contratacaoContratanteRgInput) contratacaoContratanteRgInput.value = rgInput ? rgInput.value : '';
    if (contratacaoContratanteOrgaoEmissorInput) contratacaoContratanteOrgaoEmissorInput.value = expInput ? expInput.value : '';

    // New mappings
    if (contratacaoContratanteSexoSelect) {
        const sourceValue = sexoInput ? sexoInput.value : '';
        // Convert 'M'/'F' to 'Masculino'/'Feminino' for the target select
        const mappedValue = sourceValue === 'M' ? 'Masculino' : (sourceValue === 'F' ? 'Feminino' : '');
        const optionExists = Array.from(contratacaoContratanteSexoSelect.options).some(option => option.value === mappedValue);
        contratacaoContratanteSexoSelect.value = optionExists ? mappedValue : '';
        if (sourceValue !== '' && mappedValue === '' && !optionExists) {
             console.warn(`Sex value "${sourceValue}" from source field could not be mapped or is not a valid option in the target select.`);
        }
    }
    if(dadosSexoTitularProposta) dadosSexoTitularProposta.value = contratacaoContratanteSexoSelect ? contratacaoContratanteSexoSelect.value : '';
    if (contratacaoContratanteEstadoCivilSelect) contratacaoContratanteEstadoCivilSelect.value = estCivilInput ? estCivilInput.value : '';
    if (contratacaoContratanteCelularInput) contratacaoContratanteCelularInput.value = celularInput ? celularInput.value : '';
    if (contratacaoContratanteWhatsappInput) contratacaoContratanteWhatsappInput.value = celularInput ? celularInput.value : ''; // Assuming whatsapp is same as cellular
    if (contratacaoContratanteEmailInput) contratacaoContratanteEmailInput.value = emailInput ? emailInput.value : '';
    if (contratacaoContratanteNomeMaeInput) contratacaoContratanteNomeMaeInput.value = maeInput ? maeInput.value : '';
    if (contratacaoContratanteCepInput) contratacaoContratanteCepInput.value = cepInput ? cepInput.value : '';
    if (contratacaoContratanteLogradouroInput) contratacaoContratanteLogradouroInput.value = endInput ? endInput.value : '';
    if (contratacaoContratanteNumeroInput) contratacaoContratanteNumeroInput.value = numInput ? numInput.value : '';
    if (contratacaoContratanteComplementoInput) contratacaoContratanteComplementoInput.value = complInput ? complInput.value : '';
    if (contratacaoContratanteBairroInput) contratacaoContratanteBairroInput.value = bairroInput ? bairroInput.value : '';
    if (contratacaoContratanteCidadeInput) contratacaoContratanteCidadeInput.value = cidadeInput ? cidadeInput.value : '';
    if (contratacaoContratanteEstadoInput) contratacaoContratanteEstadoInput.value = estInput ? estInput.value : '';
    if (contratacaoContratanteDataNascimentoInput) contratacaoContratanteDataNascimentoInput.value = nascInput ? nascInput.value : '';

    // Update CPF contratante in Resumo section
    if (contratacaoResumoCpfContratanteInput) {
        contratacaoResumoCpfContratanteInput.value = cpfInput ? cpfInput.value : '';
    }
}

// Function to update Contratacao dependent fields based on União de Benefícios dependent data
function updateContratacaoDependentFields() {
    if (contratacaoDep1NomeInput) contratacaoDep1NomeInput.value = dep1NomeInput ? dep1NomeInput.value : '';
    if (contratacaoDep1DataNascimentoInput) contratacaoDep1DataNascimentoInput.value = dep1NascInput ? dep1NascInput.value : '';
    if (contratacaoDep1ParentescoInput) contratacaoDep1ParentescoInput.value = dep1ParentescoInput ? dep1ParentescoInput.value : '';

    if (contratacaoDep2NomeInput) contratacaoDep2NomeInput.value = dep2NomeInput ? dep2NomeInput.value : '';
    if (contratacaoDep2DataNascimentoInput) contratacaoDep2DataNascimentoInput.value = dep2NascInput ? dep2NascInput.value : '';
    if (contratacaoDep2ParentescoInput) contratacaoDep2ParentescoInput.value = dep2ParentescoInput ? dep2ParentescoInput.value : '';

    if (contratacaoDep3NomeInput) contratacaoDep3NomeInput.value = dep3NomeInput ? dep3NomeInput.value : '';
    if (contratacaoDep3DataNascimentoInput) contratacaoDep3DataNascimentoInput.value = dep3NascInput ? dep3NascInput.value : '';
    if (contratacaoDep3ParentescoInput) contratacaoDep3ParentescoInput.value = dep3ParentescoInput ? dep3ParentescoInput.value : '';

    if (contratacaoDep4NomeInput) contratacaoDep4NomeInput.value = dep4NomeInput ? dep4NomeInput.value : '';
    if (contratacaoDep4DataNascimentoInput) contratacaoDep4DataNascimentoInput.value = dep4NascInput ? dep4NascInput.value : '';
    if (contratacaoDep4ParentescoInput) contratacaoDep4ParentescoInput.value = dep4ParentescoInput ? dep4ParentescoInput.value : '';
}

// Function to update the Responsible fields based on titular data
function updateContratacaoResponsavelFields() {
     // Copy CPF from Titular section to Responsavel section
     if (contratacaoResponsavelCpfInput) contratacaoResponsavelCpfInput.value = cpfInput ? cpfInput.value : '';
     // Copy Nome from Titular section to Responsavel section
     if (contratacaoResponsavelNomeInput) contratacaoResponsavelNomeInput.value = nomeInput ? nomeInput.value : '';
     // Copy Email from Titular section to Responsavel section
     if (contratacaoResponsavelEmailInput) contratacaoResponsavelEmailInput.value = emailInput ? emailInput.value : '';
     // Copy Nome da Mãe from Titular section to Responsavel section
     if (contratacaoResponsavelNomeMaeInput) contratacaoResponsavelNomeMaeInput.value = maeInput ? maeInput.value : '';
     // Copy Celular from Titular section to Responsavel section 
     if (contratacaoResponsavelCelularInput) contratacaoResponsavelCelularInput.value = celularInput ? celularInput.value : '';
     // New mappings
     if (contratacaoResponsavelDataNascimentoInput) contratacaoResponsavelDataNascimentoInput.value = nascInput ? nascInput.value : '';
     if (contratacaoResponsavelEstadoCivilSelect) contratacaoResponsavelEstadoCivilSelect.value = estCivilInput ? estCivilInput.value : '';
     if (contratacaoResponsavelRgCnhInput) contratacaoResponsavelRgCnhInput.value = rgInput ? rgInput.value : '';
     if (contratacaoResponsavelCepInput) contratacaoResponsavelCepInput.value = cepInput ? cepInput.value : '';
     if (contratacaoResponsavelLogradouroInput) contratacaoResponsavelLogradouroInput.value = endInput ? endInput.value : '';
     if (contratacaoResponsavelNumeroInput) contratacaoResponsavelNumeroInput.value = numInput ? numInput.value : '';
     if (contratacaoResponsavelComplementoInput) contratacaoResponsavelComplementoInput.value = complInput ? complInput.value : '';
     if (contratacaoResponsavelBairroInput) contratacaoResponsavelBairroInput.value = bairroInput ? bairroInput.value : ''; 
     if (contratacaoResponsavelCidadeInput) contratacaoResponsavelCidadeInput.value = cidadeInput ? cidadeInput.value : '';
     if (contratacaoResponsavelEstadoInput) contratacaoResponsavelEstadoInput.value = estInput ? estInput.value : '';

     // Update Sexo radio buttons based on source Sexo select
     if (sexoInput && contratacaoResponsavelSexoMasculinoRadio && contratacaoResponsavelSexoFemininoRadio) {
         const sourceValue = sexoInput.value;
         if (sourceValue === 'M') {
             contratacaoResponsavelSexoMasculinoRadio.checked = true;
             contratacaoResponsavelSexoFemininoRadio.checked = false;
         } else if (sourceValue === 'F') {
             contratacaoResponsavelSexoMasculinoRadio.checked = false;
             contratacaoResponsavelSexoFemininoRadio.checked = true;
         } else {
             // If source is empty or other value, uncheck both
             contratacaoResponsavelSexoMasculinoRadio.checked = false;
             contratacaoResponsavelSexoFemininoRadio.checked = false;
         }
     }
}

// Function to set the current date in the "Data da proposta" field
function setContratacaoResumoDate() {
    if (contratacaoResumoDataPropostaInput) {
        const today = new Date();
        const year = today.getFullYear();
        const month = ('0' + (today.getMonth() + 1)).slice(-2); // Months are 0-indexed
        const day = ('0' + today.getDate()).slice(-2);
        contratacaoResumoDataPropostaInput.value = `${year}-${month}-${day}`;
    }
}

// Add event listeners to the source fields to trigger updates

pesoTitularDeclaracao?.addEventListener('input', () => {
updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});

alturaTitularDeclaracao?.addEventListener('input', () => {
updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});

nomeInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
cpfInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
nascInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
rgInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
expInput?.addEventListener('input', updateContratanteFields);
estCivilInput?.addEventListener('change', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
// Add event listener for the source Sexo select
sexoInput?.addEventListener('change', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields(); 
});
maeInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
emailInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
celularInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
endInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
numInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
complInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
bairroInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields(); 
});
cidadeInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});
estInput?.addEventListener('input', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoResponsavelFields();
});

// Add event listeners for dependent fields
dep1NomeInput?.addEventListener('input', updateContratacaoDependentFields);
dep1NascInput?.addEventListener('input', updateContratacaoDependentFields);
dep1ParentescoInput?.addEventListener('input', updateContratacaoDependentFields);
dep2NomeInput?.addEventListener('input', updateContratacaoDependentFields);
dep2NascInput?.addEventListener('input', updateContratacaoDependentFields);
dep2ParentescoInput?.addEventListener('input', updateContratacaoDependentFields);
dep3NomeInput?.addEventListener('input', updateContratacaoDependentFields);
dep3NascInput?.addEventListener('input', updateContratacaoDependentFields);
dep3ParentescoInput?.addEventListener('input', updateContratacaoDependentFields);
dep4NomeInput?.addEventListener('input', updateContratacaoDependentFields);
dep4NascInput?.addEventListener('input', updateContratacaoDependentFields);
dep4ParentescoInput?.addEventListener('input', updateContratacaoDependentFields);

// Initial update on page load in case fields have pre-filled values
document.addEventListener('DOMContentLoaded', () => {
    updateOdontoTitularFields();
    updateContratanteFields();
    updateContratacaoDependentFields();
    updateContratacaoResponsavelFields();
    setContratacaoResumoDate(); 
});

// --- Helper Function to get form data ---
function getFormData() {
    const data = {};
    // Select all inputs except checkbox/radio, and all selects, within the form container
    const inputs = document.querySelectorAll('.form-container input:not([type="checkbox"]):not([type="radio"]), .form-container select');
    inputs.forEach(input => {
        const key = input.id || input.name; // Use ID or name as the key
        if (key) {
            data[key] = input.value;
        }
    });

    // Handle checkboxes (including those in all sections)
    const checkboxes = document.querySelectorAll('.form-container input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
        const key = checkbox.id || checkbox.name;
        if (key) {
            data[key] = checkbox.checked;
        }
    });

    // Handle radio buttons (including those in all sections)
    const radioGroups = {};
    document.querySelectorAll('.form-container input[type="radio"]').forEach(radio => {
        if (radio.name && radio.checked) {
            radioGroups[radio.name] = radio.value;
        }
    });
    // Merge radio group values into the main data object
    Object.assign(data, radioGroups);

    // Consolidate dependent data from the first dependent section
    for (let i = 1; i <= 8; i++) {
        const depNameInput = document.querySelector(`input[name="dep${i}_nome"]`);
        // Check if the dependent name input exists and has a value
        if (depNameInput && depNameInput.value.trim() !== '') {
             data[`dep${i}_nome`] = depNameInput.value;
             data[`dep${i}_nasc`] = document.querySelector(`input[name="dep${i}_nasc"]`)?.value || '';
             data[`dep${i}_parentesco`] = document.querySelector(`input[name="dep${i}_parentesco"]`)?.value || '';
             data[`dep${i}_plano_amb`] = document.querySelector(`input[name="dep${i}_plano_amb"]`)?.checked || false;
             data[`dep${i}_plano_comp`] = document.querySelector(`input[name="dep${i}_plano_comp"]`)?.checked || false;
             data[`dep${i}_plano_odonto`] = document.querySelector(`input[name="dep${i}_plano_odonto"]`)?.checked || false;
             data[`dep${i}_assist_funeral`] = document.querySelector(`input[name="dep${i}_assist_funeral"]`)?.checked || false;
        } else {
            // If name is empty, clear any residual data for this dependent number
            data[`dep${i}_nome`] = '';
            data[`dep${i}_nasc`] = '';
            data[`dep${i}_parentesco`] = '';
            data[`dep${i}_plano_amb`] = false;
            data[`dep${i}_plano_comp`] = false;
            data[`dep${i}_plano_odonto`] = false;
            data[`dep${i}_assist_funeral`] = false;
        }
    }

    // Ensure proposal number from header is captured if not already by id/name
    if (!data['proposal_number']) { // Check the primary key used
        const proposalInput = document.querySelector('.proposal-number');
        if(proposalInput) {
             data['proposal_number'] = proposalInput.value;
        }
    }

    return data;
}


// --- Helper Function to set form data ---
function setFormData(data) {
    // Updated selector to include the new section inputs by their IDs/names
    const inputs = document.querySelectorAll('.form-container input[type="text"], .form-container input[type="email"], .form-container input[type="tel"], .form-container input[type="date"], .form-container select, .form-container input.proposal-number, #declaration_rs_value, #declaration_local, #declaration_dia, #declaration_mes, #declaration_ano, #agent_name, #agent_registro');
    inputs.forEach(input => {
        const key = input.id || input.name || (input.classList.contains('proposal-number') ? 'proposal_number' : null);

        if (key && data.hasOwnProperty(key)) {
            // For date inputs, ensure the value is in "YYYY-MM-DD" format for the input type="date"
            // Note: declaration date inputs are type="text", no need for yyyy-mm-dd format conversion
            if (input.type === 'date') {
                const dateValue = data[key];
                 if (typeof dateValue === 'string' && dateValue) {
                    let parts;
                    if (dateValue.includes('/')) { // dd/mm/yyyy
                        parts = dateValue.split('/');
                        if (parts.length === 3) {
                             input.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                        } else {
                             input.value = ''; // Invalid format
                        }
                    } else if (dateValue.includes('-')) { // yyyy-mm-dd etc.
                         input.value = dateValue; // Assume it's already in a suitable format
                    } else {
                        input.value = ''; // Unknown format
                    }
                } else if (dateValue instanceof Date && !isNaN(dateValue)) {
                     // Handle potential Date objects from XLSX
                    const year = dateValue.getFullYear();
                    const month = (dateValue.getMonth() + 1).toString().padStart(2, '0');
                    const day = dateValue.getDate().toString().padStart(2, '0');
                    input.value = `${year}-${month}-${day}`;
                }
                 else {
                    input.value = ''; // Empty or non-date-like value
                }
            } else {
                 input.value = data[key] || ''; // Set value for other input types
            }

             // Trigger blur for validation on CPF if it exists
            if (input.id === 'cpf') {
                const blurEvent = new Event('blur');
                input.dispatchEvent(blurEvent);
            }
        } else {
             // If data is missing for a field, clear it in the form
             input.value = '';
        }
    });

    const checkboxes = document.querySelectorAll('.form-container input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
        const key = checkbox.id || checkbox.name;
        if (key && data.hasOwnProperty(key)) {
             // Treat any non-empty, non-false value from Excel as checked
            const value = data[key];
            checkbox.checked = value === true || String(value).toUpperCase() === 'SIM' || (typeof value === 'string' && value.trim() !== '' && value.toUpperCase() !== 'NÃO');
        } else {
             checkbox.checked = false; // Default to unchecked if data is missing or 'Não'
        }
    });

    // Set dependent data
    for (let i = 1; i <= 8; i++) {
        const depNameInput = document.querySelector(`input[name="dep${i}_nome"]`);
        const depNascInput = document.querySelector(`input[name="dep${i}_nasc"]`);
        const depParentescoInput = document.querySelector(`input[name="dep${i}_parentesco"]`);
        const depPlanoAmbInput = document.querySelector(`input[name="dep${i}_plano_amb"]`);
        const depPlanoCompInput = document.querySelector(`input[name="dep${i}_plano_comp"]`);
        const depPlanoOdontoInput = document.querySelector(`input[name="dep${i}_plano_odonto"]`);
        const depAssistFuneralInput = document.querySelector(`input[name="dep${i}_assist_funeral"]`);

        if (depNameInput) depNameInput.value = data[`dep${i}_nome`] || '';

        // Set date value for dependent nascimento (type="date")
        if (depNascInput) {
            const dateValue = data[`dep${i}_nasc`];
             if (typeof dateValue === 'string' && dateValue) {
                 let parts;
                 if (dateValue.includes('/')) { // dd/mm/yyyy
                     parts = dateValue.split('/');
                     if (parts.length === 3) {
                          depNascInput.value = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                     } else {
                          depNascInput.value = '';
                     }
                 } else if (dateValue.includes('-')) { // yyyy-mm-dd etc.
                      depNascInput.value = dateValue;
                 } else {
                     depNascInput.value = '';
                 }
             } else if (dateValue instanceof Date && !isNaN(dateValue)) {
                const year = dateValue.getFullYear();
                const month = (dateValue.getMonth() + 1).toString().padStart(2, '0');
                const day = dateValue.getDate().toString().padStart(2, '0');
                depNascInput.value = `${year}-${month}-${day}`;
             }
             else {
                 depNascInput.value = '';
             }
        }

        if (depParentescoInput) depParentescoInput.value = data[`dep${i}_parentesco`] || '';

        // Set dependent checkboxes
        if (depPlanoAmbInput) depPlanoAmbInput.checked = data[`dep${i}_plano_amb`] === true || String(data[`dep${i}_plano_amb`]).toUpperCase() === 'SIM' || (typeof data[`dep${i}_plano_amb`] === 'string' && data[`dep${i}_plano_amb`].trim() !== '' && data[`dep${i}_plano_amb`].toUpperCase() !== 'NÃO');
        if (depPlanoCompInput) depPlanoCompInput.checked = data[`dep${i}_plano_comp`] === true || String(data[`dep${i}_plano_comp`]).toUpperCase() === 'SIM' || (typeof data[`dep${i}_plano_comp`] === 'string' && data[`dep${i}_plano_comp`].trim() !== '' && data[`dep${i}_plano_comp`].toUpperCase() !== 'NÃO');
        if (depPlanoOdontoInput) depPlanoOdontoInput.checked = data[`dep${i}_plano_odonto`] === true || String(data[`dep${i}_plano_odonto`]).toUpperCase() === 'SIM' || (typeof data[`dep${i}_plano_odonto`] === 'string' && data[`dep${i}_plano_odonto`].trim() !== '' && data[`dep${i}_plano_odonto`].toUpperCase() !== 'NÃO');
        if (depAssistFuneralInput) depAssistFuneralInput.checked = data[`dep${i}_assist_funeral`] === true || String(data[`dep${i}_assist_funeral`]).toUpperCase() === 'SIM' || (typeof data[`dep${i}_assist_funeral`] === 'string' && data[`dep${i}_assist_funeral`].trim() !== '' && data[`dep${i}_assist_funeral`].toUpperCase() !== 'NÃO');
    }
}


// --- CPF Validation Function ---
function validateCPF(cpf) {
    cpf = String(cpf).replace(/[^\d]/g, ''); // Remove non-numeric characters

    if (cpf.length !== 11) return false; // CPF must have 11 digits
    if (/^(\d)\1+$/.test(cpf)) return false; // Invalid sequence like 111.111.111-11

    let sum = 0;
    let remainder;

    // Validate first digit
    for (let i = 1; i <= 9; i++) sum = sum + parseInt(cpf.substring(i - 1, i)) * (11 - i);
    remainder = (sum * 10) % 11;
    if ((remainder === 10) || (remainder === 11)) remainder = 0;
    if (remainder !== parseInt(cpf.substring(9, 10))) return false;

    sum = 0;
    // Validate second digit
    for (let i = 1; i <= 10; i++) sum = sum + parseInt(cpf.substring(i - 1, i)) * (12 - i);
    remainder = (sum * 10) % 11;
    if ((remainder === 10) || (remainder === 11)) remainder = 0;
    if (remainder !== parseInt(cpf.substring(10, 11))) return false;

    return true; // CPF is valid
}

// --- Event Listener for "Gerar Arquivo" ---
generateBtn.addEventListener('click', () => {
    try {
        if(nomeInput.value == "")
        {
            alert("Por Favor, informe o nome antes de gerar o arquivo.")
            nomeInput.focus();
            nomeInput.classList.add('invalid');
            return;
        }
        // Get CPF value and validate it
        const cpfValue = cpfInput.value;
        if (!validateCPF(cpfValue)) {
            alert("Por favor, insira um CPF válido antes de gerar o arquivo.");
            cpfInput.focus(); // Focus the CPF field
            cpfInput.classList.add('invalid'); // Ensure it's marked invalid
            return; // Stop the process
        }

        const nomeFile = nomeInput.value;

        // Generate filename from valid CPF
        const cleanedCpf = cpfValue.replace(/[^\d]/g, ''); // Remove non-numeric chars
        const defaultFileName = `${nomeFile}.xlsx`;

        const formData = getFormData();

        // Prepare data for worksheet
        // Convert boolean checkboxes to "Sim"/"Não" strings for clarity in Excel
        const excelData = {};
        for (const key in formData) {
             if (formData.hasOwnProperty(key)) {
                 if (typeof formData[key] === 'boolean') {
                     excelData[key] = formData[key] ? 'Sim' : 'Não';
                 } else {
                     excelData[key] = formData[key];
                 }
             }
        }

        const ws = XLSX.utils.json_to_sheet([excelData], {
            // Explicitly set cell types for dates if needed, although defaults often work
            // cellDates: true // This might cause issues if dates aren't perfectly formatted
        });

        // Optional: Adjust column widths (example)
        // const cols = [ { wch: 20 }, { wch: 15 }, { wch: 10 } ]; // Adjust based on your data
        // ws['!cols'] = cols;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "CadastroSocio"); // Sheet name

        // Generate workbook binary data
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

        // Create a Blob
        const blob = new Blob([wbout], { type: "application/octet-stream" });

        // Create a temporary anchor element
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", defaultFileName); // Suggest the filename
        link.style.visibility = 'hidden';
        document.body.appendChild(link);

        // Simulate a click to trigger the download/save dialog
        link.click();

        // Clean up
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
        return; // No file selected
    }

    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            // Assume the first sheet is the one we want
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert sheet to JSON (array of objects)
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (jsonData.length > 0) {
                // Convert "Sim"/"Não" back to boolean before setting
                const loadedData = jsonData[0];
                const processedData = {};
                 for(const key in loadedData) {
                     if (loadedData.hasOwnProperty(key)) {
                         const value = loadedData[key];
                         if (typeof value === 'string') {
                              if (value.toUpperCase() === 'SIM') {
                                  processedData[key] = true;
                              } else if (value.toUpperCase() === 'NÃO') {
                                  processedData[key] = false;
                              } else {
                                  processedData[key] = value;
                              }
                         } else {
                             processedData[key] = value;
                         }
                     }
                 }
                setFormData(processedData); // Load the data from the first row
                alert("Dados carregados com sucesso!");
            } else {
                alert("A planilha selecionada está vazia.");
            }
        } catch (error) {
            console.error("Erro ao ler o arquivo:", error);
            alert("Erro ao ler o arquivo. Verifique se é um arquivo Excel válido (.xlsx ou .xls).");
        } finally {
            fileInput.value = null; // Reset file input to allow selecting the same file again
        }
    };

    reader.onerror = (e) => {
        console.error("Erro ao ler o arquivo:", e);
        alert("Ocorreu um erro ao tentar ler o arquivo.");
        fileInput.value = null; // Reset file input
    };

    reader.readAsArrayBuffer(file); // Read the file as ArrayBuffer
});




function collectFormData() {
    const data = [];
    const inputs = document.querySelectorAll(
        '#proposal_number_uniao, #orgao, #entidade, #proponente, ' +
        '#nome, #nasc, #sexo, #est_civil, #rg, #exp, #cpf, #mae, #email, ' +
        '#banco, #agencia, #conta_corrente, #conjuge, #nasc_conjuge, #sexo_conjuge, ' +
        '#end, #num, #compl, #bairro, #cep, #cidade, #est, #tel, #celular, ' +
        '#orgao_funcional, #mat_funcional, #funcao, #upag, #unidade, #setor, ' +
        '#contratacaoResponsavelSexoMasculinoRadio, #contratacaoResponsavelSexoFemininoRadio'+
        '#mensalidade_social, #plano_saude_amb, #plano_saude_comp, #orientacao_juridica, #plano_odonto, #seguro_vida, input[name="seguro_vida_sim"], input[name="seguro_vida_nao"], #auxilio_natalidade, #assistencia_funeral, #convenios, ' +
        '#dep1_nome, #dep1_nasc, #dep1_parentesco, #dep1_plano_amb, #dep1_plano_comp, #dep1_plano_odonto, #dep1_assist_funeral, ' +
        '#dep2_nome, #dep2_nasc, #dep2_parentesco, #dep2_plano_amb, #dep2_plano_comp, #dep2_plano_odonto, #dep2_assist_funeral, ' +
        '#dep3_nome, #dep3_nasc, #dep3_parentesco, #dep3_plano_amb, #dep3_plano_comp, #dep3_plano_odonto, #dep3_assist_funeral, ' +
        '#dep4_nome, #dep4_nasc, #dep4_parentesco, #dep4_plano_amb, #dep4_plano_comp, #dep4_plano_odonto, #dep4_assist_funeral, ' +
        '#dep5_nome, #dep5_nasc, #dep5_parentesco, #dep5_plano_amb, #dep5_plano_comp, #dep5_plano_odonto, #dep5_assist_funeral, ' +
        '#dep6_nome, #dep6_nasc, #dep6_parentesco, #dep6_plano_amb, #dep6_plano_comp, #dep6_plano_odonto, #dep6_assist_funeral, ' +
        '#dep7_nome, #dep7_nasc, #dep7_parentesco, #dep7_plano_amb, #dep7_plano_comp, #dep7_plano_odonto, #dep7_assist_funeral, ' +
        '#dep8_nome, #dep8_nasc, #dep8_parentesco, #dep8_plano_amb, #dep8_plano_comp, #dep8_plano_odonto, #dep8_assist_funeral, ' +
        '#declaration_rs_value, #declaration_local, #declaration_dia, #declaration_mes, #declaration_ano, ' +
        '#agent_name, #agent_registro, ' +
        '#contratacao_proposal_number, #contratacao_vendedor_cpf, #contratacao_vendedor_nome, #contratacao_vendedor_telefone, #contratacao_cnpj_corretora, #contratacao_nome_corretora, ' +
        '#contratacao_contratante_responsavel, #contratacao_contratante_cpf, #contratacao_contratante_nome, #contratacao_contratante_sexo, #contratacao_contratante_rg, #contratacao_contratante_data_emissao_rg, #contratacao_contratante_orgao_emissor, #contratacao_contratante_estado_civil, #contratacao_contratante_data_nascimento, #contratacao_contratante_telefone_fixo, #contratacao_contratante_celular, #contratacao_contratante_whatsapp, #contratacao_contratante_cartao_nacional_saude, #contratacao_contratante_email, #contratacao_contratante_nome_mae, input[name="contratacao_contratante_plano_anterior"], #contratacao_contratante_qual_plano, #contratacao_contratante_data_inicio, #contratacao_contratante_data_ultimo_pagamento, #contratacao_contratante_cep, #contratacao_contratante_logradouro, #contratacao_contratante_numero, #contratacao_contratante_complemento, #contratacao_contratante_bairro, #contratacao_contratante_cidade, #contratacao_contratante_estado' +
        '#dadosNomeTitularTermo, #dadosDataNascTitularTermo, #dadosEnderecoTitularTermo, #dadosNumberTitulatTermo, #dadosComplementoTitularTermo, #dadosBairroTitularTermo, #dadosCepTitularTermo, #dadosCidadeTitularTermo, #dadosEstadoTitularTermo #dadosEmailTitularTermo, #dadosTelefoneTitularTermo, #dadosTipoDocumentoTitularTermoOdonto, #dadosNumeroDocumentoTitularTermoOdonto, #dadosNacionalidadeTitularTermoOdonto, #dadosPaisEmissaoTitularTermo, #dadosResidenciaFiscalTitularTermo, #dadosLocasNascimentoTitularTermo, #dadosProfissaoTitularTermoOdonto, #dadosDetalheOcupacaoTitularTermo, #dadosRendaMediaMensalTitularTermo' +
        '#dadosNomeTitularProposta, #dadosdataNascTitularProposta, #dadosEstadoCivilTitularProposta, #dadosEndTitularProposta, #dadosBairroTitularProposta, #dadosCidadeTitularProposta, #dadosCepTitularProposta, #dadosTelTitularProposta, #dadosNacionalidadeTitularProposta, #dadosManutencaoTitularProposta, #dadosUopTitularProposta, #dadosCiaTitularProposta, #dadosSucursalTitularProposta, #dadosRamoTitularProposta, #dadosApoliceTitularProposta, #dadosNumeroCertificadoTitularProposta, #dadosGrupoTitularProposta, #dadosPlanoTitularProposta, #dadosProLaboreTitularProposta, #dadosEstipulanteTitularProposta, #dadosEstruturaVendaTitularProposta, #dadosPesoTitularProposta, #dadosAlturaTitularProposta, #dadosCargoTitularProposta, #dadosRendaMensalTitularProposta, #dadosDataAdmissaoTitularProposta, #dadosInicioVigenciaTitularProposta, #dadosTerminoVigenciaTitularProposta, #dadosNomeResponsavelTitularProposta, #dadosCpfResponsavelTitularProposta, #dadosCusteioTitularProposta, #dadosEmpresaTitularProposta, #dadosFuncionarioTitularProposta'
    );
    inputs.forEach(input => {
        let value = '';
        if (input.type === 'checkbox') {
            value = input.checked ? 'X' : '';
        } else if (input.type === 'radio') {
            if (input.checked) {
                value = input.value;
            } else {
                return; 
            }
        }
        else if (input.tagName === 'SELECT') {
            value = input.value;
        } else {
            value = input.value;
        }
        data.push(value); 
    });
    return data;
}

function populateFormData(data) {
    const inputs = document.querySelectorAll(
        '#proposal_number_uniao, #orgao, #entidade, #proponente, ' +
        '#nome, #nasc, #sexo, #est_civil, #rg, #exp, #cpf, #mae, #email, ' +
        '#banco, #agencia, #conta_corrente, #conjuge, #nasc_conjuge, #sexo_conjuge, ' +
        '#end, #num, #compl, #bairro, #cep, #cidade, #est, #tel, #celular, ' +
        '#orgao_funcional, #mat_funcional, #funcao, #upag, #unidade, #setor, ' +
        '#contratacaoResponsavelSexoMasculinoRadio, #contratacaoResponsavelSexoFemininoRadio'+
        '#mensalidade_social, #plano_saude_amb, #plano_saude_comp, #orientacao_juridica, #plano_odonto, #seguro_vida, input[name="seguro_vida_sim"], input[name="seguro_vida_nao"], #auxilio_natalidade, #assistencia_funeral, #convenios, ' +
        '#dep1_nome, #dep1_nasc, #dep1_parentesco, #dep1_plano_amb, #dep1_plano_comp, #dep1_plano_odonto, #dep1_assist_funeral, ' +
        '#dep2_nome, #dep2_nasc, #dep2_parentesco, #dep2_plano_amb, #dep2_plano_comp, #dep2_plano_odonto, #dep2_assist_funeral, ' +
        '#dep3_nome, #dep3_nasc, #dep3_parentesco, #dep3_plano_amb, #dep3_plano_comp, #dep3_plano_odonto, #dep3_assist_funeral, ' +
        '#dep4_nome, #dep4_nasc, #dep4_parentesco, #dep4_plano_amb, #dep4_plano_comp, #dep4_plano_odonto, #dep4_assist_funeral, ' +
        '#dep5_nome, #dep5_nasc, #dep5_parentesco, #dep5_plano_amb, #dep5_plano_comp, #dep5_plano_odonto, #dep5_assist_funeral, ' +
        '#dep6_nome, #dep6_nasc, #dep6_parentesco, #dep6_plano_amb, #dep6_plano_comp, #dep6_plano_odonto, #dep6_assist_funeral, ' +
        '#dep7_nome, #dep7_nasc, #dep7_parentesco, #dep7_plano_amb, #dep7_plano_comp, #dep7_plano_odonto, #dep7_assist_funeral, ' +
        '#dep8_nome, #dep8_nasc, #dep8_parentesco, #dep8_plano_amb, #dep8_plano_comp, #dep8_plano_odonto, #dep8_assist_funeral, ' +
        '#declaration_rs_value, #declaration_local, #declaration_dia, #declaration_mes, #declaration_ano, ' +
        '#agent_name, #agent_registro, ' +
        '#contratacao_proposal_number, #contratacao_vendedor_cpf, #contratacao_vendedor_nome, #contratacao_vendedor_telefone, #contratacao_cnpj_corretora, #contratacao_nome_corretora, ' +
        '#contratacao_contratante_responsavel, #contratacao_contratante_cpf, #contratacao_contratante_nome, #contratacao_contratante_sexo, #contratacao_contratante_rg, #contratacao_contratante_data_emissao_rg, #contratacao_contratante_orgao_emissor, #contratacao_contratante_estado_civil, #contratacao_contratante_data_nascimento, #contratacao_contratante_telefone_fixo, #contratacao_contratante_celular, #contratacao_contratante_whatsapp, #contratacao_contratante_cartao_nacional_saude, #contratacao_contratante_email, #contratacao_contratante_nome_mae, input[name="contratacao_contratante_plano_anterior"], #contratacao_contratante_qual_plano, #contratacao_contratante_data_inicio, #contratacao_contratante_data_ultimo_pagamento, #contratacao_contratante_cep, #contratacao_contratante_logradouro, #contratacao_contratante_numero, #contratacao_contratante_complemento, #contratacao_contratante_bairro, #contratacao_contratante_cidade, #contratacao_contratante_estado' + 
        '#dadosNomeTitularTermo, #dadosDataNascTitularTermo, #dadosEnderecoTitularTermo, #dadosNumberTitulatTermo, #dadosComplementoTitularTermo, #dadosBairroTitularTermo, #dadosCepTitularTermo, #dadosCidadeTitularTermo, #dadosEstadoTitularTermo #dadosEmailTitularTermo, #dadosTelefoneTitularTermo, #dadosTipoDocumentoTitularTermoOdonto, #dadosNumeroDocumentoTitularTermoOdonto, #dadosNacionalidadeTitularTermoOdonto, #dadosPaisEmissaoTitularTermo, #dadosResidenciaFiscalTitularTermo, #dadosLocasNascimentoTitularTermo, #dadosProfissaoTitularTermoOdonto, #dadosDetalheOcupacaoTitularTermo, #dadosRendaMediaMensalTitularTermo' +
        '#dadosNomeTitularProposta, #dadosdataNascTitularProposta, #dadosEstadoCivilTitularProposta, #dadosEndTitularProposta, #dadosBairroTitularProposta, #dadosCidadeTitularProposta, #dadosCepTitularProposta, #dadosTelTitularProposta, #dadosNacionalidadeTitularProposta, #dadosManutencaoTitularProposta, #dadosUopTitularProposta, #dadosCiaTitularProposta, #dadosSucursalTitularProposta, #dadosRamoTitularProposta, #dadosApoliceTitularProposta, #dadosNumeroCertificadoTitularProposta, #dadosGrupoTitularProposta, #dadosPlanoTitularProposta, #dadosProLaboreTitularProposta, #dadosEstipulanteTitularProposta, #dadosEstruturaVendaTitularProposta, #dadosPesoTitularProposta, #dadosAlturaTitularProposta, #dadosCargoTitularProposta, #dadosRendaMensalTitularProposta, #dadosDataAdmissaoTitularProposta, #dadosInicioVigenciaTitularProposta, #dadosTerminoVigenciaTitularProposta, #dadosNomeResponsavelTitularProposta, #dadosCpfResponsavelTitularProposta, #dadosCusteioTitularProposta, #dadosEmpresaTitularProposta, #dadosFuncionarioTitularProposta'
    );
    let dataIndex = 0;
    inputs.forEach(input => {
        if (dataIndex < data.length) {
            const cellValue = data[dataIndex] !== undefined && data[dataIndex] !== null ? String(data[dataIndex]) : '';

            if (input.type === 'checkbox') {
                input.checked = (cellValue.toUpperCase() === 'X' || cellValue.toLowerCase() === 'true' || cellValue === '1');
            } else if (input.type === 'radio') {
                if (input.value === cellValue) {
                    input.checked = true;
                } else {
                    input.checked = false; 
                }
            }
            else if (input.tagName === 'SELECT') {
                const optionExists = Array.from(input.options).some(option => option.value === cellValue);
                input.value = optionExists ? cellValue : '';
            } else {
                input.value = cellValue;
            }
            dataIndex++;
        } else {
            if (input.type === 'checkbox') {
                input.checked = false;
            } else if (input.type === 'radio') {
                input.checked = false;
            }
            else if (input.tagName === 'SELECT') {
                input.value = '';
            } else {
                input.value = '';
            }
        }
    });
    updateOdontoTitularFields();
}

sendEmailBtn.addEventListener('click', () => {
    console.log("Send Email button clicked. Server-side implementation needed.");
    alert("Funcionalidade de envio de e-mail não implementada neste ambiente.");
});