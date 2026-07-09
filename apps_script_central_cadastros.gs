/**
 * CENTRAL DE CADASTROS - GOOGLE APPS SCRIPT
 * Planilha conectada: Controle cadastro
 * ID: 1yOi0nIdYD701qygVNuc0FPzTFXvordAVlDPX26WqrM4
 *
 * Como publicar:
 * 1. Abra script.google.com
 * 2. Novo projeto
 * 3. Cole este código em Code.gs
 * 4. Implantar > Nova implantação > Aplicativo da Web
 * 5. Executar como: você
 * 6. Quem tem acesso: qualquer pessoa com o link
 * 7. Copie a URL e cole no HTML em APPS_SCRIPT_URL
 */

const SPREADSHEET_ID = '1yOi0nIdYD701qygVNuc0FPzTFXvordAVlDPX26WqrM4';
const EMAIL_RESPONSAVEL = ''; // opcional: preencha com o e-mail do responsável por cadastros

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const protocolo = gerarProtocolo_();
    const agora = new Date();
    const tipo = data.tipoCadastro || 'Parceiro';

    if (tipo === 'Parceiro') gravarParceiro_(ss, data, protocolo, agora);
    else if (tipo === 'Produto') gravarProduto_(ss, data, protocolo, agora);
    else if (tipo === 'Serviço') gravarServico_(ss, data, protocolo, agora);
    else throw new Error('Tipo de cadastro inválido.');

    gravarLog_(ss, data, protocolo, tipo, agora, 'Recebido');
    enviarEmail_(data, protocolo, tipo);

    return json_({ ok: true, protocolo, message: 'Solicitação registrada com sucesso.' });
  } catch (err) {
    return json_({ ok: false, message: err.message });
  }
}

function doGet() {
  return json_({ ok: true, message: 'API Central de Cadastros ativa.' });
}

function gravarParceiro_(ss, d, protocolo, agora) {
  ss.getSheetByName('Parceiros').appendRow([
    agora, protocolo, d.solicitante, d.emailSolicitante, d.departamento, d.prioridade,
    d.tipoPessoa, d.razaoSocial, d.nomeFantasia, d.cnpjCpf, d.inscricaoEstadual,
    d.contribuinteIcms, d.cep, d.endereco, d.numero, '', d.bairro, d.cidade, d.uf,
    d.telefone, d.emailParceiro, d.banco, d.agencia, d.conta, d.pix,
    d.condicaoPagamento, d.observacoes, 'Pendente', '', '', '', d.linkAnexo || ''
  ]);
}

function gravarProduto_(ss, d, protocolo, agora) {
  ss.getSheetByName('Produtos').appendRow([
    agora, protocolo, d.solicitante, d.emailSolicitante, d.departamento, d.prioridade,
    d.descricaoProduto, d.grupoProduto, d.unidadePadrao, d.marca, d.referenciaFornecedor,
    d.ncm, d.ipi, d.controlaEstoque, d.estoqueMinimo, d.estoqueMaximo, d.custoEstimado,
    d.fornecedorProduto, d.observacoes, 'Pendente', '', '', '', d.linkAnexo || ''
  ]);
}

function gravarServico_(ss, d, protocolo, agora) {
  ss.getSheetByName('Serviços').appendRow([
    agora, protocolo, d.solicitante, d.emailSolicitante, d.departamento, d.prioridade,
    d.descricaoServico, d.grupoServico, d.unidadeServico, d.ncmNbs, d.codigoServicoMunicipio,
    d.iss, d.retemImpostos, d.fornecedorServico, d.centroResultado, d.natureza,
    d.observacoes, 'Pendente', '', '', '', '', d.linkAnexo || '', 'Pendente validação'
  ]);
}

function gravarLog_(ss, d, protocolo, tipo, agora, resultado) {
  ss.getSheetByName('Log_Envios').appendRow([agora, protocolo, tipo, JSON.stringify(d), resultado]);
}

function gerarProtocolo_() {
  const ano = new Date().getFullYear();
  const props = PropertiesService.getScriptProperties();
  const key = 'SEQ_' + ano;
  const seq = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(seq));
  return 'CAD-' + ano + '-' + String(seq).padStart(6, '0');
}

function enviarEmail_(d, protocolo, tipo) {
  if (!EMAIL_RESPONSAVEL) return;
  const assunto = `[Central de Cadastros] ${protocolo} - ${tipo}`;
  const corpo = `Nova solicitação recebida.\n\nProtocolo: ${protocolo}\nTipo: ${tipo}\nSolicitante: ${d.solicitante}\nE-mail: ${d.emailSolicitante}\nDepartamento: ${d.departamento}\nPrioridade: ${d.prioridade}\n\nAcesse a planilha Controle cadastro para tratar a solicitação.`;
  MailApp.sendEmail(EMAIL_RESPONSAVEL, assunto, corpo);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
