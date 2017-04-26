#DEFINE SEMAPHORE_ALL_ACCESS     	2031619  && 0x1F0003 
#DEFINE STANDARD_RIGHTS_REQUIRED  	983040  && 0xF0000 
#DEFINE SS_SEMAPHORE_NAME 			"Star Soft Applications CS"
#DEFINE MESSENGER_SEMAPHORE_NAME 	"Star Soft Messenger"
#DEFINE START_INFO_SIZE  68

#DEFINE adFldIsNullable 6

*- Win32 API Constants
#DEFINE ERROR_SUCCESS			0

*- Registry Value types
#DEFINE REG_NONE				0	&& Undefined Type (default)
#DEFINE REG_SZ					1	&& Regular Null Terminated string
#DEFINE REG_BINARY				3	&& ??? (unimplemented)
#DEFINE REG_DWORD				4	&& Long integer value
#DEFINE MULTI_SZ				7	&& Multiple Null Term strings (not implemented)

#DEFINE maxDword				4294967295	&& 0xffffffff

*- Registry roots
#DEFINE HKEY_CLASSES_ROOT		-2147483648	&& (( HKEY ) 0x80000000 )
#DEFINE HKEY_CURRENT_USER		-2147483647	&& (( HKEY ) 0x80000001 )
#DEFINE HKEY_LOCAL_MACHINE		-2147483646	&& (( HKEY ) 0x80000002 )
#DEFINE HKEY_USERS				-2147483645	&& (( HKEY ) 0x80000003 )

#DEFINE RECORDING			1
#DEFINE MAX_OPENFILE		225
#DEFINE CRLF				CHR(13) + CHR(10)
#DEFINE READING				2
#DEFINE FORM_PERSONALIZADO	'2'
#DEFINE FORM_GENERIC		'1'
#DEFINE FORM_GROUP			'2'
#DEFINE FORM_USER			'3'
#DEFINE CANCEL_BUTTON		8
#DEFINE DELETE_BUTTON		9
#DEFINE EDIT_BUTTON			6
#DEFINE IS_BOF				-2
#DEFINE IS_EOF				-1
#DEFINE MOVE_BOTTOM			4
#DEFINE MOVE_DOWN			3
#DEFINE MOVE_TO				11
#DEFINE MOVE_TOP			1
#DEFINE MOVE_UP  			2
#DEFINE NEW_BUTTON   		5
#DEFINE OUT_BUTTON   		10
#DEFINE SAVE_BUTTON  		7
#DEFINE FILE_ALL     		0
#DEFINE FILE_FORM    		4
#DEFINE FILE_PICTURE 		2
#DEFINE FILE_PROGRAM 		1
#DEFINE FILE_SOUND   		3
#DEFINE FRM_COMPLEMENT   	4
#DEFINE FRM_NORMAL   		0
#DEFINE FRM_FILTER   		2
#DEFINE FRM_MODAL			1
#DEFINE MBOX_CANCEL  		2
#DEFINE MBOX_NO  			7
#DEFINE MBOX_YES	 		6

#DEFINE SV_FatalErr         -1
#DEFINE SV_NoFatalErr        1

#DEFINE MSG_NaoExisteInsere  1
#define _TRUE_				.t.
#define _FALSE_				.f.
#define _ALLBUTTONS_		.t.
#define _ALLTOOLBARS_		.t.

*- Constantes para o Debug Mode
#define Debug_SqlErr			2   && fol_sqlexec
#define Debug_MovAccumulator	4	&& moveaccumulator

*- Cores do Sistema
#define MAX_SYSTEMCOLOR		30
#define AFIELDS2NDCOL		18

*- Atributos de arquivo
#DEFINE FILE_ATTRIBUTE_READONLY       1
#DEFINE FILE_ATTRIBUTE_HIDDEN         2
#DEFINE FILE_ATTRIBUTE_SYSTEM         4
#DEFINE FILE_ATTRIBUTE_DIRECTORY     16
#DEFINE FILE_ATTRIBUTE_ARCHIVE       32
#DEFINE FILE_ATTRIBUTE_NORMAL       128
#DEFINE FILE_ATTRIBUTE_TEMPORARY    512
#DEFINE FILE_ATTRIBUTE_COMPRESSED  2048


DEFINE CLASS CGO_LibMain01 AS custom
  hidden VOA_AccessCfg[4, 4]						 && Array para strings de acesso
  VOA_AccessCfg = ""
  HIDDEN VON_TypeVersion                             && Indica o tipo de Versão: Comercial, Apresentação, Beta 
  HIDDEN VOA_Licence[20]                             && Armazena as informações necessárias para impressão de autorização de licenças

  HIDDEN voc_userManType  							 && Tipo de Manutecao que este usuario paga, Nenhuma(0), Silver(1), Gold(2) ou Platinum(3)
  HIDDEN voc_userTrustedConn						 && Indica que a validacao do usuario se dara pela existencia dele no Dominio (ignora o Password)
  HIDDEN voc_NeverGetThisProperties					 && Lista de propriedades tipo HIDDEN que nao podem ser selecionadas con as funcoes FOU_Get e FOU_Set

  PROTECTED VOC_SecurityCodeSP                       && Utilizado na geração do Servicepack 
  VOC_SecurityCodeSP = 'StarSoft Security Code (Service Pack)' 
 

  DIMENSION VOA_SteLog[1,3]                         && Starsoft Translator Engine  (Log do Sistema)
  DIMENSION voa_007_numbers[1]                      && Números e Símbolos para formaçao do Extenso [ a própria função faz a inicialização ]
  DIMENSION voa_007_signs[15]                       && Números e Símbolos para formaçao do Extenso [ a própria função faz a inicialização ]
  DIMENSION voa_aparc[1]                            && Guarda o valor das parcelas de uma condição de pagamento
  DIMENSION voa_auxarred[29]                        && Guarda os valores para arredondamento de moedas
  DIMENSION voa_avenc[1]                            && Guarda os vencimentos reais de uma condição de pagamento
  DIMENSION voa_avencn[1]                           && Guarda os vencimentos nominais de uma condição de pagamento
  DIMENSION voa_b_currency[3]                       && Armazena os três sistemas monetários utilizados pelo cliente na contabilidade
  DIMENSION voa_currencynames[15]                   && Guarda os nomes de moedas
  DIMENSION voa_currencyvalues[15]                  && Guarda os valores para conversão das moedas com índice > 1 para a com 1
  DIMENSION voa_filter[10]                          && Filtros para clausulas IN no SQL.
  DIMENSION voa_formparam[1]                        && Guarda os parâmetros para os forms
  DIMENSION voa_iconpointer[10,3]                   && Aponta para os icons dos  modulos. Utilizado para selecionar os mesmos atravez do menu de Sistema. 1=Objeto, 2=ID, 3=Prompt.
  DIMENSION voa_index[10]                           && Array de indexes para o fou_editcursor.
  DIMENSION voa_instances[1,4]                      && Contém nomes de forms, referências aos objetos, o número de instâncias ativas, e o número da próxima disponível.
  DIMENSION voa_parameters[1]                       && Armazena parâmetros a serem passados para um form.
  DIMENSION voa_picts[1]                            && Indica os picts do sistema
  DIMENSION voa_rep_currency[5]                     && Armazena 5 S.M. para utilização dos relatórios
  DIMENSION voa_seeksql[10]                         && Guarda os VGC_Seek's
  DIMENSION voa_stackalias[300,5]                   && Armazena pilha para o pushselect e popselect. Posicoes: 1- nome do alias, 2 - Tag corrente, 3 - recno[].
  DIMENSION voa_ukeyofparameterscrd[5,2]            && [1,1] - CRD I Contábil;  [1,2] - CRD II Contábil;  [2,1] - CRD I Estoque;  [2,2] - CRD II Estoque;  [3,1] - CRD I Folha;  [3,2] - CRD II Folha;  [4,1] - CRD I Finanças;  [4,2] - CRD II Finanças;  [5,1
  DIMENSION voa_where[30]                           && Indica os where para a criação dos cursores.
  DIMENSION voo_agentchar[10]                       && Array para armazenar os personagens.
  DIMENSION voa_ciacodes[1, 2]                      && Empresas do usuário
  DIMENSION voa_sets[1, 2]							&& Parâmetros e respostas
  DIMENSION voa_LastSql[1, 2]                       && Armazena os Selects "resolvidos"  dos ultimos cursores Executados
  DIMENSION voa_LastSqlOptions[1, 2]                && Armazena as opções dos Selects "resolvidos"
  DIMENSION voa_LastA54[1, 4]						&& Armazena os dados das últimas pesquisas de alteração de moeda(A57) para otimização da FON_VM()
  DIMENSION voa_LastA37[1, 5]						&& Armazena os dados das últimas pesquisas de conversões de sistemas monetários(A37) para otimização da FON_VM()
  DIMENSION voa_LastA70[1, 5]						&& Armazena os dados das últimas pesquisas de conversões de sistemas monetários projetado (A70) para otimização da FON_VMP()
  dimension VOA_ReportFilter[1] 					&& Expressão utilizada para filtros do tipo Hidden em Relatórios
  dimension VOA_TriggerReturn[1]					&& Retorno de triggers
  dimension VOA_ValidTableStruct[1, 2]			   	&& Armazena as estruturas de campos validos para uma tabela para a FOL_CreateSQLString
  dimension VOA_SystemColors[MAX_SYSTEMCOLOR + 1]	&& Armazena as cores do windows (RGB)
  dimension VOA_ReportResults[1]					&& Armazena os arquivos necessários para apresentar o resultado do relatório
  dimension VOA_formreturn[5]                       && Guarda MULTIPLOS retornos de um form
  dimension VOA_SystemTbl[1]						&& Array com os arquivos de sistema
  dimension VOA_TableIdent[1]						&& Matriz com as tabelas e seus respectivos identificadores
  Declare VOA_TableStdFilter[1]					&& Armazena os filtros padrões dos arquivos

  Name = "cgo_libmain01"

  voc_datacontrol = "components\controls\data\"
  voc_cepstardirectory = ""                          && Armazena o Diretorio de instalacao do PlugIn CEPSTAR
  voc_cnxdirectory = "components\controls\connections"	&& Armazena o Diretório onde o arquivo de conexões será colocado
  voc_cnxFile      = "connect.cnx"						&& Armazena o arquivo de conexões
  
  voc_bannerdirectory = .F.                          && Armazena o diretório onde os Banners para Web serão sugeridos.Vide parâmetro this.FOU_ChkSets("STARPAR-C-00093     ")
  voc_catalogname = ""                               && Nome padrão do catalog
  voc_clipartdirectory = .F.                         && Armazena o diretório onde os Cliparts para Web serão sugeridos.Vide parâmetro this.FOU_ChkSets("STARPAR-C-00092     ")
  voc_companyaddress = ""                            && Guarda o endereço da companhia corrente
  voc_companycgc = ""                                && Guarda o cgc da companhia corrente
  voc_companycity = (space(20))                      && Guarda o código cidade da companhia corrente
  voc_companycityn = ""                              && Guarda o nome da cidade da companhia corrente
  voc_companycode = (space(5))                       && Guarda o código da companhia corrente
  voc_companycountry = (space(20))                   && Guarda o código país da companhia corrente
  voc_companycountryn = ""                           && Guarda o nome país da companhia corrente
  voc_companydddphone = .F.                          && Prefixo DDD do telefone da empresa.
  voc_companyie = (space(20))                        && Armazena a incrição estadual da empresa.
  voc_companyjc = ""                                 && Guarda a junta comercial da companhia corrente
  voc_companylogo = ""                               && Guarda o arquivo com o Logotipo da empresa para impressão em relatórios.
  voc_companyname = ""                               && Guarda o nome da companhia corrente
  voc_companyneighborhood = ""                       && Guarda o bairro da companhia corrente
  voc_companynumber = .F.                            && Numero da empresa no endereço indicado em VOC_CompanyAdress.
  voc_companyphone = .F.                             && Telefone padrão da empresa.
  voc_companystate = (space(20))                     && Guarda o código estado(UF) da companhia corrente
  voc_companystaten = ""                             && Guarda o nome estado(UF) da companhia corrente
  voc_companyuf = ""                                 && Guarda a UF da companhia corrente
  voc_companyzipcode = ""                            && Guarda o cep da companhia corrente
  voc_companycomplement = ""						 && Guarda o complemento do endereço da companhia corrente 
  voc_currency = (space(5))                          && Guarda a moeda padrão do sistema
  voc_currencyplural = "Reais"                       && Guarda o nome plural da moeda
  voc_currencysingular = "Real"                      && Guarda o nome singular da moeda
  voc_decimalplural = "Centavos"                     && Guarda o nome plural da moeda
  voc_decimalsingular = "Centavo"                    && Guarda o nome singular da moeda
  voc_dirhome =(upper(alltrim(sys(5)+sys(2003))))+"\"&& Guarda o diretório padrão do sistema
  voc_formgroup = ""                                 && Guarda o último grupo de arquivos aberto
  voc_groupcode = (space(5))                         && Guarda o código do grupo do usuário corrente
  voc_groupname = ""                                 && Guarda o nome do grupo do usuário corrente
  voc_home = ""                                      && Guarda o diretório padrão do sistema
  voc_mainmenu = ""                                  && Nome do menu principal (arquivo .MPR).
  voc_mainwindcaption = "Star Soft Applications"     && O nome principal da aplicação.
  voc_menuname = ""                                  && Guarda o nome do menu a ser executado após o login
  voc_oldwindcaption = ""                            && Nome da antiga aplicação.
  voc_prefixline = ""                                && Armazena o prefixo a ser discado quando deseja discar via modem.
  voc_presentmodule = "E"                            && Guarda o módulo atual do sistema
  voc_reportname = ""                                && Armazena o nome do report a ser executado.
  voc_returnpageselect = ""                          && Retorno da procura via Select
  voc_screenname = ""                                && Guarda o nome do programa na tela
  voc_StartPar = ""                                  && Guarda o parametro inicial do Sistema
  voc_texttoedit = ""                                && Texto para ser editado do rich text.
  voc_triggersfile = ""								 && Arquivo onde são armazenados os triggers.
  vol_validtriggersfile = .f.                        && Indica se o arquivo de triggers é válido
  voc_typeline = ""                                  && Armazena o tipo de linha (T- tone, P - pulse).
  voc_ukeyfromcombo = ""                             && Guarda o Ukey do combo para refresh do ToolBar informativo (Status)
  voc_ukeyfromform = ""                              && Guarda o Ukey do form para refresh do ToolBar informativo (Status)
  voc_user_language = ""                        	 && Guarda qual o idioma do usuário corrente
  voc_userborn = {}                                  && Guarda a data de nascimento do usuário corrente
  voc_usercivil = ""                                 && Guarda o estado civil do usuário corrente
  voc_usercode = ""                             	 && Guarda o código do usuário corrente
  voc_userdirectory = ""                             && Diretório padrão para documentos do usuário.
  voc_username = ""                                  && Guarda o nome do usuário corrente
  voc_userpassword = ""                              && Guarda o password do usuário corrente
  voc_usersex = ""                                   && Guarda o sexo do usuário corrente
  voc_userManType = 0								 && Tipo de Manutecao que este usuario paga, Nenhuma(0), Silver(1), Gold(2) ou Platinum(3)
  voc_userTrustedConn = 0							 && Indica que a validacao do usuario se dara pela existencia dele no Dominio (ignora o Password)
  voc_version = ""                                   && Guarda a versão do programa.
  voc_spversion = ""                                 && Guarda a versão service pack.
  voc_webserver = ""                                 && Nome do Web Server para ser trocado pelo caminho da rede da tela InternetOption. Vide parâmetro this.FOU_ChkSets("STARPAR-C-00091     ")
  voc_wwwhelpstarsoft = ""                           && Guardo o endereço eletronico do Suporte da Star
  voc_wwwhomestarsoft = ""                           && Guardo o endereço eletronico da Star
  voc_usercia = ""								     && String com as ukeys das empresas as quais o usuário logado tem acesso
  vod_emptydate = {}                                 && Data Vazia
  vod_maxvaliddate = {^2079/01/01}                   && Variáveis para valores máximos e mínimos para data
  vod_minvaliddate = {^1900/01/01}                   && Variáveis para valores máximos e mínimos para data
  vol_currencyprojection = .F.	                     && Variavel para verificar se utiliza projeção
  vol_changeforecolor = .F.                          && Cor das letras desabilitadas igual a das letras habilitadas.
  vol_closeconnection = .T.                          && Indica se sempre fecha a conexão com o Banco de Dados no término da transação
  vol_crd1b = .F.                                    && Indica se o módulo de contabilidade utiliza Centro de Receitas e Despesas I
  vol_crd1d = .F.                                    && Indica se o módulo de estoque utiliza Centro de Receitas e Despesas I
  vol_crd1f = .F.                                    && Indica se o módulo financeiro utiliza Centro de Receitas e Despesas I
  vol_crd1k = .F.                                    && Indica se o módulo de PCP utiliza Centro de Receitas e Despesas I
  vol_crd1m = .F.                                    && Indica se o módulo de folha de pagamento utiliza Centro de Receitas e Despesas I
  vol_crd2b = .F.                                    && Indica se o módulo de contabilidade utiliza Centro de Receitas e Despesas II
  vol_crd2d = .F.                                    && Indica se o módulo de estoque utiliza Centro de Receitas e Despesas II
  vol_crd2f = .F.                                    && Indica se o módulo de financeiro utiliza Centro de Receitas e Despesas II
  vol_crd2k = .F.                                    && Indica se o módulo de PCP utiliza Centro de Receitas e Despesas II
  vol_crd2m = .F.                                    && Indica se o módulo de folha de pagamento utiliza Centro de Receitas e Despesas II
  vol_CreateReportLayOut = .F.					     && Faz com que o conteudo dos relatorios sejam voltados para impressao do LayOut dos FRX (XXX XXXX XXXXXXXXXX)
  vol_currencyhavecents = .T.                        && Indica se a moeda possui centavos
  vol_automatedlogin = .f.						     && Indica se a automatização de login está habilitada permitindo que não seja enviada a senha do usuário
  vol_discountinorderutilization = .T.               && Define se o sistema permite o uso de desconto para itens utilizados de Pedidos em Nota Fiscal de Faturamento.
  vol_exchangeerror = .F.							 && Indica que não foi encontrada conversão para as moedas passadas para a função FON_VM()
  vol_false = .F.                                    && Variáveis para os selects
  vol_haseditmenu = .T.                              && Indica se possui o menu de edição.
  vol_isclean = .F.                                  && Indica se o ambiente está limpo.
  vol_justresponsible = .F.                          && Indica que apenas o usuário que fizer parte de um grupo de responsáveis no módulo de projetos poderá inserir, deletar, editar documentos. E o usuário só conseguirá ver os documentos que pertencem ao r
  vol_limitcare = .T.                                && Cuida dos limites de informação trazidas pelas pesquisas.
  vol_lastSeek  = .f.                                && Armazena o resultado de um seek (quando necesssario) para ser utilizada em cahamdas anteriores (ex:FON_msg e FOC_Caption)
  vol_resizeformdesktop = .T.                        && Esta propriedade é utilizada para indicar ao form (parece um toolbar) de ferramentas de RESIZE se deve ser um form do tipo DESKTOP ou não.
  vol_return = .F.                                   && Guarda o retorno do objeto
  vol_showuniquekey = .T.                            && Se mostra a chave indicando que o campo não pode ser repetido.
  vol_smartspinner = .T.                             && Indica se o spinner será tipo calculadora
  vol_stopcalc = .F.                                 && Variável utilizada para interromper o cálculo Longos (Ativo, Contabil, PCP).
  vol_testingenv = .F.                               && Indica se está em modo de teste.
  vol_testmode = .F.                                 && Variáveis auxiliares para testes automáticos através de Macros (FKY)
  vol_true = .T.                                     && Variáveis para os selects
  vol_workwithgrade = .F.                            && Indica se a empresa trabalha com Grade.
  vol_workwithprojects = .F.                         && Indica se a empresa trabalha com o controle de apropriação de despesas do módulo de projetos.
  vol_workwithraster = .F.                           && Indica se a empresa trabalha com rastreabilidade.
  vol_submenus = .F.								 && Indica se subdivisão de áreas está ativa ou não.
  VOL_LoggedOn = .F.								 && Indica se existe um usuário corretamente logado e validado no sistema
  von_accurracy = 0                                  && Precisão, em casas decimais, parâmetro STARPAR-N-00089.
  von_activechar = 0                                 && Indica o personagem ativo.
  von_calcmode = 1                                   && Indica o modo da calculadora - qual calculador os Spinners vão usar 1-Fox, 2-Windows, 3-Star
  von_connreadonly = .F.                             && Conexão read/only.
  von_connupdate = .F.                               && Conexão para update.
  von_icons = 0                                      && Guarda o número de ícones inicializados no sistema.
  von_LastSql = 0									 && Aramazena os Selects "resolvidos"  dos ultimos cursores Executados
  von_LastSqlOptions = 0							 && Aramazena as opções dos Selects "resolvidos"
  von_maxeditmemo = 250                              && Guarda o máximo de edição do campo Memo
  von_maxrecordsattime = 4000                        && Indica o número máximo de registros por execução de select.
  von_maxrecordsforwarning = 2000                    && Indica o número máximo de registros para dar warning.
  von_modemport = -1                                 && Armazena o port que contém o modem.
  von_nchkcgc = 0                                    && Conta número do Contador para verificação do CGC/CPF/PIS
  von_parameters = .F.                               && Armazena o número de parâmetros.
  von_reprocess = 5                                  && Numero de reprocessos para funções.
  von_roundtype = 0                                  && Tipo de arredondamento, parâmetro STARPAR-N-00088.
  von_screenobject = 0                               && Número e ponteiro ao objeto da tela
  von_stackalias = 0                                 && Armazena o topo da pilha para o pushselect e popselect.
  von_thiscalltype = (FRM_NORMAL)                    && Guarda o tipo de tela que está sendo chamada
  von_tiposenha = 0                                  && Guarda o tipo de senha
  von_whatversion = 0                                && Armazena a versão do foxpro.
  voo_agent = .NULL.                                 && Cria a classe de agentes.
  voo_barcode = .NULL.                               && Guarda o objeto de conversão Código Barras.
  voo_lastscreenobject = .NULL.                      && Aponta para o ultimo icone pressionado, e necessario para controlar o efeito dos desenhos sobre os mesmos
  voo_screenobject = .NULL.                          && Associado  ao VGN_ScreenObject, VGO_ScreenObject - Número e ponteiro ao objeto da tela VOO_ScreenObject
  vou_formreturn = ""                                && Guarda o retorno da última tela
  vol_triggerreturn = .t.                            && Armazena o retorno de um trigger
  VOD_BClosingDate = {}                              && Data de fechamento do módulo Contábil
  VOD_DClosingDate = {}                              && Data de fechamento do módulo de Estoque
  VOD_EClosingDate = {}                              && Data de fechamento do módulo de Compras
  VOD_FClosingDate = {}                              && Data de fechamento do módulo Financeiro
  VOD_GClosingDate = {}                              && Data de fechamento do módulo Ativo
  VOD_HClosingDate = {}                              && Data de fechamento do módulo de Livros Fiscais
  VOD_IClosingDate = {}								 && Data de fechamento do módulo de Assistência Técnica
  VOD_JClosingDate = {}                              && Data de fechamento do módulo de Vendas
  VOD_KClosingDate = {}                              && Data de fechamento do módulo de PCP
  VOD_MClosingDate = {}                              && Data de fechamento do módulo de Folha de Pagamento
  VOD_NClosingDate = {}                              && Data de fechamento do módulo de Ponto Eletrônico
  VOD_OClosingDate = {}                              && Data de fechamento do módulo de Recursos Humanos
  VOD_RClosingDate = {}                              && Data de fechamento do módulo de Projetos
  VOD_VClosingDate = {}								 && Data de fechamento do módulo de Custos
  VOL_RateioContabilObrigatorio = .f.				 && Obriga o rateio contábil
  VOL_RateioFinanceiroObrigatorio = .f.				 && Obriga o rateio financeiro
  VOL_RateioEstoqueObrigatorio = .f.			     && Obriga o rateio do estoque 
  VOL_RateioFolhaObrigatorio = .f.				     && Obriga o rateio do Folha
  VON_EditDocument = 2								 && 1-Permite editar um documento gerado pela integração, 2-Não permite e 3-Pergunta se realmente que editar.
  voc_MachineId   = SYS(0)						     && Identificacao da maquina#usuario
  voc_NetMachine  = ""							     && Identificacao da maquina
  VOC_NetPathResource  = ""							 && Identificação do Resource Temporário do usuário
  voc_NetUserName = "" 				    		     && Identificacao do usuario
  voc_DirSettings = "" 			    			     && Diretorio para gravacao do Resource do usuario
  voc_FileSettings = "" 			    			 && Arquivo Resource do usuario
  voc_CutCopyFile = "" 			    			     && Arquivo Utilizado para armazenar as informacoes de CUT/COPY dos registros
  voc_UserTmp = ""								     && Diretorio para criacao de arquivos temporarios do usuario 
  VOC_ClientRepoDir = "my-apps\"					 && Diretório onde são armazenados as configurações específicas de relatório
  VOC_ClientSearchDir = "my-apps\"				     && Diretório onde são armazenados as configurações específicas de pesquisas
  VOC_ClientSSAADir = "my-apps\"					 && Diretório onde são armazenados as configurações específicas de Análises
  VOD_LastSaveDateTime = {}					 	     && Último datetime utilizado no save de um registro, é utilizado para principalmente para controlar travamento de registros no Servidor
  VOL_ServiceMode = .F.								 && Indica se o sistema foi inicializado no modo serviço (sem interface)
  VOC_LastMessage = ""								 && Guarda a Ultima message utilizada em FON_Msg()
  VON_ReportFilterCount = 0							 && Número de Filtros setados para o relatório
  VON_TriggerReturnCount = 0						 && Retorno para triggers
  VON_Array445 = 1									 && Indica se a conta analítica deverá ser gerada parar (Cliente, fornecedor ou ambos).
  VOC_A03B11UkeySynthetic = ""						 && Ukey da conta sintética do cliente.
  VOC_A08B11UkeySynthetic = ""						 && Ukey da conta sintética do fornecedor.
  
  VON_TableStructCount = 0							 && Contador de tabelas já armazenadas em VOA_ValidTableStruct
  VON_RepoOutflowRecord = 0							 && Indica se os relatórios de registro de saida utiliza a data da emissao ou do registro
  VON_RepoInflowRecord = 0							 && Indica se os relatórios de registro de saida utiliza a data da emissao ou do registro
  VOL_AccountingCover = .F.							 && Indica se deve ser utilizado o valor da capa

  *-- Added by Denis na Revisão das conexões
  VOL_ValidCnxFile = .F.			&& Indica se o arquivo de conexão é válido
  VOC_ConnectionString = ""			&& String de conexão com o banco de dados
  VON_DBType = 0					&& Banco de Dados que será acessado 1-MS SQL-Server, 2-Oracle, 3-DB2
  VOC_TableIdentifier = ""			&& Identificador (database, owner, etc) dos arquivos de configuração
  VON_CnxHandle = 0					&& Handle para a conexão aberta, 0 indica que não há conexões abertas
  VON_CnxStatus = -1				&& Status da conexão: -1: Fechada, 1: Aberta
  VON_TranLevel = 0					&& Número de transações abertas
  VON_TableIdentCount = 0			&& Número de Identificadores de tabelas
  VOL_FillCurrency = .f.			&& Indica se o usuário é responsável pela manutenção de moedas
  VON_Error = 0						&& Número do erro Star Soft
  VOC_ErrorPar = ""					&& String com parâmetros de erro, para tratamento de mensagens
  VON_SqlError = 0					&& Indica o erro ocorrido no SQL.
  VOC_SqlError = ""					&& Armazena a mensagem de erro do sql.
  VOL_ShowDevKit = .f.				&& Indica se informacoes para desenvolvimento serao exisbidas para o usuario (Login.scx)
*---

  VOL_SolveVarsInSQLCmd = .f.		&& Indica se as variaveis do comando SQL (VPA_Seek[n]) devem ser trocadas pelo seu conteúdo
  VOC_DateFormat = ""				&& Formato de data da máquina do usuário
  VON_ParCountEC = 0				&& Contador de parâmetros (VPA_Seek´s) do ExecuteCursor
  VON_ParCountSearch = 0			&& Contador de parâmetros (VPA_Filter's) da pesquisa. Seu valor é atualizado na tela de pesquisa Y34 no metodo FOL_SelectConstruct do formset. 
  VOC_VersionCfgFile = ""			&& Armazena o arquivo version.txt no carregamento do sistema
  VOC_DebugSQLCmdFile = ""			&& Arquivo para armazenar as selects resolvidas com seu tempo de execução, para detectar erros e demoras
  VON_DebugType = 0					&& Armazena o tipo de debug do sistema
  VON_TypeVersion = 0               && Indica o tipo de Versão: Comercial, Apresentação, Beta 
  VOL_SpecialPrint = .t.
  VOC_UsrMsgAccount = ""			&& Conta de Mensagem padrão do usuário
  VOC_ReplyAddress = ""				&& Conta para retorno de mensagens enviadas
  VOC_UsrT77Ukey = .null.			&& Ukey da Conta de Mensagem padrão do usuário

  *- Campos que devem ser excluídos do CreateSQLString
  VOC_SQLExcludeFields = " MYCONTROL SQLCMD TIMESTAMP STATUS ROWGUID CHECKCOL "

  *- Lista de propriedades tipo HIDDEN que nao podem ser selecionadas con as funcoes FOU_Get e FOU_Set
  VOC_NeverGetThisProperties = "voc_NeverGetThisProperties,VOA_AccessCfg,voc_EditString,voc_NewString,voc_DelString,von_TypeVersion,"
  VOO_Security = null

  VOO_JobProc = null				&& Objeto que efetua o processamento de tarefas agendadas

  VON_ReportResultsCount = 0		&& Quantidade de elementos do Array VOA_ReportResults
  VOL_StopAllRules = .F.			&& Verifica se no momento de retorno .t. no workflow para de processar as demais regras
  VOC_SystemMsgAccount = ""			&& Conta administrativa de mensagens inicializada pelo parâmetro STARPAR-C-00276
  VOC_SystemT77Ukey = .null.		&& Ukey da conta administrativa de mensagens inicializada pelo parâmetro STARPAR-C-00276
  VOO_WinSock = .null.				&& Objeto para manipulação de winsock
  VON_SemaphoreHandle = 0			&& Handle ao semaforo de identificação
  VON_MessengerProcessHnd = 0		&& Handle ao processo do messenger, necessário para finalizar o processo
  VON_MessengerMode = 1				&& Modo de tratamento do fechamento do messenger (1 - Usuário, 0 - Sistema)
  VOC_OpenedTrans = ""				&& Nome da primeira transação aberta no banco de dados
  VON_TableStdFilterCount = 0		&& Quantidade de configurações para filtros padrões de tabelas
  VON_DataSessionId = 0				&& DataSession corrente
  VOC_AccIntFile = ""				&& Biblioteca de integração contábil
  VOC_CostIntFile = ""				&& Biblioteca de integração de custos
  VOO_LoadingForm = null			&& Form em processo de carregamento, para documentação de cursores
  VOL_EnableDevArea = .f.			&& Indica se o ambiente de desenvolvimento deve ser habilitado
  VOC_InitialSetPath = ""			&& set path inicial que é enviado pelo startmain
  VOC_LastCaptionProperties = ""	&& Properties do último caption traduzido


*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_MakeAccessString
*/ Parametros       : PLC_User  : Usuário
*/ Descrição        : Monta a String de acesso do sistema
*/ Retorno          : String de Acesso											  
*/----------------------------------------------------------------------------------------
procedure FOL_MakeAccessString

lparameters PLC_User

dimension this.VOA_AccessCfg[4, 4]
store "" to this.VOA_AccessCfg



if !PLC_User == "STAR_" and !empty(PLC_User)

	*- Verifica o acesso por usuario
	*- 0 - Respeita o grupo
	*- 1 - Acesso liberado
	*- 2 - Acesso negado
	*- Caso o usuário faça parte de mais de um grupo, a liberação de acesso a um grupo tem prioridade sobre a negação
	*-  para outro grupo, sendo sobreposta apenas pela liberação direta no usuário

	this.FOL_SetParameter(1, PLC_User)
	this.FOL_ExecuteCursor("Y14_ACCESS", "Y14", 1)
	select y14
	index on y14_001_c tag y14_001_c

	if seek("SSACS_ACCESS_", "y14", "y14_001_c")

		set procedure to tests\b.prg,  foxcodeplus.prg 
	
		
		
		
		*- Novo acesso, baseado em strings já montadas
		scan
			VLN_1stDim = 0
			VLN_2ndDim = 0
			VLC_Option = alltrim(y14.y14_001_c)
			VLN_Plus2ndDim = 0
			
			if empty(nvl(y14.usr_ukey, ""))
				VLN_Plus2ndDim = 2
			endif
			
			do case
				case VLC_Option = "SSACS_ACCESS_GRANT_VIEW"
					VLN_1stDim = 1
					VLN_2ndDim = 1

				case VLC_Option = "SSACS_ACCESS_DENY_VIEW"
					VLN_1stDim = 1
					VLN_2ndDim = 2

				case VLC_Option = "SSACS_ACCESS_GRANT_EDIT"
					VLN_1stDim = 2
					VLN_2ndDim = 1

				case VLC_Option = "SSACS_ACCESS_DENY_EDIT"
					VLN_1stDim = 2
					VLN_2ndDim = 2

				case VLC_Option = "SSACS_ACCESS_GRANT_DELETE"
					VLN_1stDim = 3
					VLN_2ndDim = 1

				case VLC_Option = "SSACS_ACCESS_DENY_DELETE"
					VLN_1stDim = 3
					VLN_2ndDim = 2

				case VLC_Option = "SSACS_ACCESS_GRANT_ADD"
					VLN_1stDim = 4
					VLN_2ndDim = 1

				case VLC_Option = "SSACS_ACCESS_DENY_ADD"
					VLN_1stDim = 4
					VLN_2ndDim = 2

			endcase

			if VLN_1stDim > 0 and VLN_2ndDim > 0
				VLN_2ndDim = VLN_2ndDim + VLN_Plus2ndDim
				this.VOA_AccessCfg[VLN_1stDim, VLN_2ndDim] = this.VOA_AccessCfg[VLN_1stDim, VLN_2ndDim] + y14.y14_006_m
			endif
		endscan
		use in y14

	else
		*- Acesso da maneira antiga, mantenho por compatibilidade
		create cursor tmp_access (y14_001_c c(240), usr_view c(1), usr_edit c(1), usr_del c(1), usr_new c(1),;
			 grp_view c(1), grp_edit c(1), grp_del c(1), grp_new c(1))
		select tmp_access
		index on y14_001_c tag y14_001_c

		store ";" to this.VOA_AccessCfg



		select y14
		scan
			if !seek(padr(y14.y14_001_c, 240), "tmp_access", "y14_001_c")
				append blank in tmp_access
				replace tmp_access.y14_001_c	with y14.y14_001_c in tmp_access
			endif
			
			if !empty(nvl(y14.usr_ukey, ""))
				replace tmp_access.usr_view		with y14.y14_002_c,;
						tmp_access.usr_edit		with y14.y14_003_c,;
						tmp_access.usr_del		with y14.y14_004_c,;
						tmp_access.usr_new		with y14.y14_005_c;
				in tmp_access
			else
				replace tmp_access.grp_view		with max(y14.y14_002_c, tmp_access.grp_view),;
						tmp_access.grp_edit		with max(y14.y14_003_c, tmp_access.grp_edit),;
						tmp_access.grp_del		with max(y14.y14_004_c, tmp_access.grp_del),;
						tmp_access.grp_new		with max(y14.y14_005_c, tmp_access.grp_new);
				in tmp_access
			endif
		endscan			
		use in y14

		select tmp_access
		scan
			VLC_Option = alltrim(tmp_access.y14_001_c) + ";"
			if tmp_access.usr_view == "1" or ((empty(tmp_access.usr_view) or tmp_access.usr_view == "0") and tmp_access.grp_view == "1")
				this.VOA_AccessCfg[1, 1] = this.VOA_AccessCfg[1, 1] + VLC_Option
			endif

			if tmp_access.usr_edit == "1" or ((empty(tmp_access.usr_edit) or tmp_access.usr_edit == "0") and tmp_access.grp_edit == "1")
				this.VOA_AccessCfg[2, 1] = this.VOA_AccessCfg[2, 1] + VLC_Option
			endif

			if tmp_access.usr_del == "1" or ((empty(tmp_access.usr_del) or tmp_access.usr_del == "0") and tmp_access.grp_del == "1")
				this.VOA_AccessCfg[3, 1] = this.VOA_AccessCfg[3, 1] + VLC_Option
			endif

			if tmp_access.usr_new == "1" or ((empty(tmp_access.usr_new) or tmp_access.usr_new == "0") and tmp_access.grp_new == "1")
				this.VOA_AccessCfg[4, 1] = this.VOA_AccessCfg[4, 1] + VLC_Option
			endif
		endscan

		use in tmp_access
	endif
endif	

return .t.


VOC_Between = ""


*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_InitArray                                                		  
*/ Parâmetros  		: 																	  
*/ Descrição   		: Coloca na Memoria todos os Arrays Utilizados pelo sistema. 
*/						Conforme o Idioma do Usuario ativo				  		
*/ Última alteração : 24/12/2001                                                
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
PROCEDURE FOL_InitArray

	this.FOL_UseInPackage("y01", "y01")

	select y01
	set order to standard
	=seek("voa_")
	scan while y01.y01_006_c = "voa_" 

		do case
			case this.VOC_User_Language = "055"	&&	Português
				VLC_Array = y01.y01_012_m
			case this.VOC_User_Language = "001"	&&	Inglês
				VLC_Array = y01.y01_007_m
			case this.VOC_User_Language = "034"	&&	Espanhol
				VLC_Array = y01.y01_008_m
			Other
				VLC_Array = y01.y01_012_m		&&	Português
		endcase
		
		PLC_ArrayName = alltrim(y01.y01_006_c)
		VLN_Elements  = memline( VLC_Array )
		VLC_Property  = PLC_ArrayName + '[' + str(max(1,VLN_Elements) ) + ']'
		PLC_ArrayName = 'this.' + PLC_ArrayName

		if !pemstatus(this, VLC_Property, 5 )
			this.addproperty(VLC_Property)
		endif

		if VLN_Elements > 0
			*	inicia a matriz
			dimension(PLC_ArrayName + '[' + str(VLN_Elements) + ']')
			for VLN_i = 1 to VLN_Elements
				store mline(VLC_Array , VLN_i ) to (PLC_ArrayName + '[' + str(VLN_i) + ']')
			endfor
		else
			store 'Unknown Array ' + PLC_ArrayName  to (PLC_ArrayName+ '[1]')
			store "" to PLC_Store2, PLC_Store2, PLC_Store3, PLC_Store1
		endif
		
		
		
		
		
		
	endscan

ENDPROC 

*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_BaseCaption : Caption em português para procura         		  
*/					: PLC_Group : Grupo para procurar 1-array, 3-captions, 4-menus
*/ Descrição		: Procura no arquivo captions por captions para tradução
*/ Retorno			: A caption traduzida  
*/----------------------------------------------------------------------------------------
procedure FOC_Caption(PLC_Rodrigo, PLC_Stephanie)
lparameters PLC_BaseCaption, PLL_ForceStd



	LOCAL VLC_Return, VLC_Exact, VLC_BaseCaption

	this.VOC_LastCaptionProperties = ""

	if empty(PLL_ForceStd) and !empty(VGO_Custom.VOL_CustomCaption)
		*- Faço a verificação de tradução específica
		if !isnull(VGO_Custom.VOO_CustomCaptionObj)
			this.VOL_LastSeek = .t.
			VLC_Return = VGO_Custom.VOO_CustomCaptionObj.FOC_Caption(PLC_BaseCaption)
			this.VOC_LastCaptionProperties = VGO_Custom.VOC_LastCaptionProperties
			return VLC_Return
		endif
	endif

	*- Rotina de tradução padrão
	VLC_BaseCaption = lower(PLC_BaseCaption)
	
	VLC_Return = ""

	if !used("y01")
		VLC_CurAlias = alias()
		this.FOL_UseInPackage("y01", "y01")
		if !empty(VLC_CurAlias)
			select (VLC_CurAlias)
		endif
	endif

	VLC_Exact = set("EXACT")
	
	set exact on
	
	this.VOL_LastSeek = seek(VLC_BaseCaption, "y01" , "standard")
	
	set exact &VLC_Exact
	
	if this.VOL_LastSeek
		do case
			case this.VOC_User_Language = "055"	&&	Português
				VLC_Return = y01.y01_012_m
			case this.VOC_User_Language = "001"	&&	Inglês
				VLC_Return = y01.y01_007_m
			case this.VOC_User_Language = "034"	&&	Espanhol
				VLC_Return = y01.y01_008_m
			Other
				VLC_Return = y01.y01_012_m		&&	Português
		endcase
	endif

	if !this.VOL_LastSeek or empty(VLC_Return)

		if version(2) <> 0
			VLC_Return = this.VOC_User_Language+ "!" +VLC_BaseCaption
		else
			VLC_Return = PLC_BaseCaption
		endif

	endif

	this.VOC_LastCaptionProperties = y01.y01_999_m

	return VLC_Return
endproc



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias: Indica o alias que quer verificar  
*/                    PLC_Tag  : Indica a tag de qual se quer o key 
*/ Descrição        : Verifica o key de uma tag específica  
*/ Retorno          : key da tag PLC_Tag 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_AliasKey
	LParameters PLC_Alias, PLC_Tag

	LOCAL VLN_Select, VLC_Ret

	this.FOL_PushSelect()

	*- Indica o alias que quer verificar
	select (PLC_Alias)
	VLN_Select = select()

	VLC_Ret = key(tagno(PLC_Tag))

	this.FOL_PopSelect()

	return VLC_Ret
ENDPROC




*/----------------------------------------------------------------------------------------
*/ Parametros       : Nenhum 
*/ Descrição        : Retorna um ukey para chave de arquivo 
*/ Retorno          : O ukey (EMPRESA + USUARIO + sys(2015))  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_GetUkey
	return this.VOC_CompanyCode + this.VOC_UserCode + sys(2015)
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias: Indica o alias que quer verificar  
*/ Descrição        : Retorna o nome real do alias  
*/ Retorno          : nome real do alias 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_RealNameAlias
	LParameters PLC_Alias

	LOCAL VLC_Ret

	VLC_Ret = alltrim(upper(PLC_Alias))
	do while VLC_Ret = "RDO_"
		VLC_Ret = substr(VLC_Ret, 5)
	enddo

	return VLC_Ret
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Alias     : O nome do alias a ser aberta a transação  
*/ Descrição   		: Executa um begin transaction  
*/ Retorno 			: .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_BeginTransaction
	lparameters PLC_Alias

	VLL_Ret = .F.

	if this.VON_TranLevel <= 0 && Não há transações abertas

		do case
			case this.VON_DBType = 1 && Sql Server
				VLL_Ret = this.FOL_SqlExec("BEGIN TRANSACTION")
			case this.VON_DBType = 2 && Oracle
				VLL_Ret = this.FOL_SqlExec("SET TRANSACTION READ WRITE")
			case this.VON_DBType = 3 && DB2
				VLL_Ret = this.FOL_SqlExec("commit")
		endcase

		if VLL_Ret
			this.VON_TranLevel = this.VON_TranLevel + 1
			this.VOC_OpenedTrans = "TRAN_" + PLC_Alias
		endif
	else
		VLL_Ret = .T.
		this.VON_TranLevel = this.VON_TranLevel + 1
	endif

	return VLL_Ret
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Alias     : O nome do alias a ser fechada a transação 
*/ Descrição   		: Executa um end transaction 
*/ Retorno 			: .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CommitTrans
	LParameters PLC_Alias

	VLL_Ret = .F.

	VLC_TransName = "TRAN_" + PLC_Alias
	if this.VON_TranLevel > 0 && Há transações abertas
		*- Não trabalhamos com transações encadeadas, logo verifico se a transação é a primeira que foi aberta
		if this.VOC_OpenedTrans == VLC_TransName

			do case
				case this.VON_DBType = 1 && Sql Server
					VLL_Ret = this.FOL_SqlExec("COMMIT TRANSACTION")
				case this.VON_DBType = 2 && Oracle
					VLL_Ret = this.FOL_SqlExec("COMMIT")
				case this.VON_DBType = 3 && DB2
					VLL_Ret = this.FOL_SqlExec("COMMIT")
			endcase

			if VLL_Ret
				this.VON_TranLevel = 0
				this.VOC_OpenedTrans = ""
			endif
		else
			VLL_Ret = .T. && Não trabalhamos com commit de transações intermediárias
		endif
	endif

	return VLL_Ret
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Alias     : O nome do alias a ser dado rollback 
*/                    PLL_All	    : Todas as Transações					  
*/ Descrição   		: Executa um rollback  
*/ Retorno 			: .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_RollBack
	LParameters PLC_Alias, PLL_All

	VLL_Ret = .F.

	VLC_TransName = "TRAN_" + PLC_Alias
	if this.VON_TranLevel > 0 && Há transações abertas
		*- Não trabalhamos com transações encadeadas, logo verifico se a transação é a primeira que foi aberta
		if this.VOC_OpenedTrans == VLC_TransName or PLL_All

			do case
				case this.VON_DBType = 1 && Sql Server
					VLL_Ret = this.FOL_SqlExec("ROLLBACK TRANSACTION")
				case this.VON_DBType = 2 && Oracle
					VLL_Ret = this.FOL_SqlExec("ROLLBACK")
				case this.VON_DBType = 3 && DB2
					VLL_Ret = this.FOL_SqlExec("ROLLBACK")
			endcase

			if VLL_Ret
				this.VON_TranLevel = 0
				this.VOC_OpenedTrans = ""
			endif
		else
			VLL_Ret = .T. && Não trabalhamos com rollback de transações intermediárias
		endif
	endif

	return VLL_Ret
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_TableName : O nome do table a ser fechado 
*/ Descrição        : Fecha o table PLC_TableName, e todos relacionados a ele. Se não for 
*/                    passado o parâmetro fecha todos.  
*/ Retorno          : .T. : Fechou com sucesso, .F. : c.c 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CloseAllTable
	LParameters PLC_TableName, PLL_AllBut

	LOCAL VLN_i, VLC_File

	if empty(PLC_TableName)
		close data all
		return .T.
	else
		return this.FOL_CloseTable(PLC_TableName)
	endif
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_TableName : O nome do table a ser fechado 
*/ Descrição        : Fecha o table PLC_TableName 
*/ Retorno          : .T. : Fechou com sucesso, .F. : c.c 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CloseTable
	LParameters PLC_TableName

	LOCAL  PLN_NoSelect

	if used( PLC_TableName )
	*--	fecha o table PLC_TableName
		use in (PLC_TableName)
	endif
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_TableName : O nome do table a ser criado  
*/              	  PLC_Fields    : String que contêm os fields 
*/             	  	  PLC_Type      : Indica se o table é um cursor ou um table 
*/               	  PLC_IndexOn   : String de indexação, no caso de mais de um índice 
*/               	  				 separar com ;      								  
*/                    PLC_Tag       : Nome do Tag, separados por ;  
*/                    PLC_TagType   : Tipo do tag, separado por ; 
*/                    PRA_Array     : Indica uma referência ao array que criará a tabela  
*/ Descrição   		: Cria o table PLC_TableName 
*/ Retorno     		: .T. : Abriu com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CreateTable
	LParameters PLC_TableName, PLC_Fields, PLC_Type, PLC_IndexOn, PLC_Tag, PLC_TagType, PRA_Array

	LOCAL VLC_Create, VLC_Index, VLN_Select, VLN_TagNumber, VLN_Counter, VLC_Tag1, VLC_Tag2
	LOCAL VLC_Tag3, VLN_PosIndexOn, VLN_PosTagName, VLN_PosTagType, VLN_LastFound1, VLN_LastFound2
	LOCAL VLN_LastFound3

	*- Cria a string de criação
	this.FOL_MakeDir(PLC_TableName, .T.)
	if !empty(PLC_Fields)
		VLC_Create = "create "+PLC_Type+" "+PLC_TableName+ " ("+PLC_Fields+")"
	else
		VLC_Create = "create "+PLC_Type+" "+PLC_TableName+ " from array PRA_Array"
	endif

	*- Cria o cursor
	select 0
	&VLC_Create

	if empty(PLC_Tag)
		PLC_Tag = []
	endif

	if empty(PLC_TagType)
		PLC_TagType = []
	endif

	if !empty(PLC_IndexOn)
		PLC_IndexOn 	= PLC_IndexOn + [;]
		PLC_Tag 		= PLC_Tag + [;]
		PLC_TagType		= PLC_TagType + [;]

		VLN_TagNumber 	= occurs([;], PLC_IndexOn)
		VLN_PosIndexOn 	= 0
		VLN_PosTagName 	= 0
		VLN_PosTagType 	= 0
		VLN_LastFound1	= 0
		VLN_LastFound2	= 0
		VLN_LastFound3	= 0

		for VLN_Counter = 1 to VLN_TagNumber
	*--		Carrega a matriz da string de indexação
			VLN_LastFound1 = at([;], PLC_IndexOn, VLN_Counter)
			VLC_Tag1 = subs(PLC_IndexOn, max(1,VLN_PosIndexOn), VLN_LastFound1-max(1,VLN_PosIndexOn))
			VLN_PosIndexOn = VLN_LastFound1+1
	*--		Carrega a matriz do tag name
			VLN_LastFound2 = at([;], PLC_Tag, VLN_Counter)
			VLC_Tag2 = subs(PLC_Tag, max(1,VLN_PosTagName), VLN_LastFound2-max(1,VLN_PosTagName))
			VLN_PosTagName = VLN_LastFound2+1
	*--		Carrega a matriz do tag type
			VLN_LastFound3 = at([;], PLC_TagType, VLN_Counter)
			VLC_Tag3 = subs(PLC_TagType, max(1,VLN_PosTagType), VLN_LastFound3-max(1,VLN_PosTagType))
			VLN_PosTagType = VLN_LastFound3+1
	*-- 	Indexa o cursor

			VLC_Index = "index on "+VLC_Tag1+" tag "+VLC_Tag2+" "+VLC_Tag3
			&VLC_Index
		endfor

	endif

	this.FOL_CursorBuffering(3)

	return .T.
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLN_Type      : tipo de buffering a ser setado  
*/              	  PLN_Select    : Select a ser setado 
*/ Descrição   		: Altera o buffering para o cursor  
*/ Retorno     		: .T. : executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CursorBuffering
	LParameters PLN_Type, PLN_Select

	set multilocks on
	if empty(PLN_Select)
		PLN_Select = select()
	endif
	return cursorsetprop("buffering", PLN_Type, PLN_Select)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Form      : Indica em qual form é executada o cancelamento  
*/                    PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/                    PLC_Which     : Guarda um campo para o cancel update  
*/                    PLL_Close     : Fecha o arquivo Base no Final 
*/                    PLL_OnlyLoop  : Indica se irá executar apenas o loop  
*/                    PLL_ToIsTable : Indica que o destino é uma tabela física  
*/ Descrição        : Deleta a informação do cursor editavel 'From' no arquivo Pai 'To' 
*/ Retorno          : .T. : Sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_DeleteCursor
	LParameters PLO_Form, PLC_FromFile, PLC_ToFile, PLC_FieldUkey, PLC_Which, PLL_Close, ;
				PLL_OnlyLoop, PLL_ToIsTable
				
	LOCAL VLL_Found, VLL_Ret, VLL_HasMycontrol, VLL_Continue, VLC_Status

	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	this.FOL_PushSelect()

	*- Muda o Status do DELETED para garantir uma Correção adequada
	VLC_OldDELETED = set("DELETED")
	set deleted off

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualização
	select (PLC_FromFile)
	this.FOL_PushSelect()

	*- Verifica se existe o campo MYCONTROL no arquivo destino
	VLL_HasMycontrol = (type(PLC_ToFile + ".MYCONTROL") = "C")

	VLL_Ret = .T.

	*- executa um scan no arquivo fonte
	scan

*---	Verifica se existe no arquivo Original
		VLU_Ukey   = eval(PLC_FieldUkey)
		if PLL_ToIsTable
			VLC_Status = eval("status")
		else
			VLC_Status = "W"
		endif

*---	se ainda não abriu ADO destino ou o registro já estava gravado, pesquisa o ADO
		if VLC_Status = "W"
			VLL_Found = seek(VLU_Ukey, PLC_ToFile, "UKEY")

*---		se não existe no arquivo original ou são distintos retorna erro
			if eof(PLC_ToFile) or eval("timestamp") <> eval(PLC_ToFile + ".timestamp")
				this.VON_Error = SV_FatalErr
				this.VON_SQLError = -1 && Timestamp diferente
				VLL_Ret = .F.
				exit
			endif

*----		Executa o cancelamento do form, se encontrou e não está
*----		  deletado volta ao valor anterior, c.c. deleta.
			if !isnull(PLO_Form)
				VLL_Continue = PLO_Form.FOL_CancelUpdates(PLC_Which)
			else
				VLL_Continue = .T.
			endif

			if VLL_Continue
				if !PLL_OnlyLoop
*------				deleta o registro no ADO
					replace mycontrol with "1" in (PLC_ToFile)
					delete in (PLC_ToFile)
				
*------				executa um update do ADO
					if !this.FOL_TableUpdate(PLC_ToFile, .F.)
						this.VON_Error = SV_NoFatalErr
						VLL_Ret = .F.
						exit
					endif
				endif
			else
				VLL_Ret = .F.
				exit
			endif

		endif


		select (PLC_FromFile)
	endscan

	go top in (PLC_FromFile)

	this.FOL_PopSelect()

	*- Fecha o arquivo origem se necessário
	if PLL_Close		&& Fecha o arquivo base no Final
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina antes de fechar
		this.FOL_CloseTable(PLC_FromFile)
	else
	*--	Executa um pack e volta todos os mycontrol com '0'
		select (PLC_FromFile)
		this.FOL_PackCursor(PLC_FromFile)
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina
	endif

	if VLC_OldDELETED = "ON"
		set deleted on
	endif
	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias : O arquivo que deve ser deletado o registro  
*/ Descrição        : Deleta o registro no arquivo PLC_Alias  
*/ Retorno          : .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_DeleteRecord
	LParameters PLC_Alias

	delete in (PLC_Alias)
	this.FOL_TableUpdate(PLC_Alias, .F.)

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Cursor   : Cursor de origem 
*/                    PLC_FromFile : Arquivo de Origem  
*/                    PLC_ToFile   : Cursor gravavel de destino, se o nome iniciar com "-"
*/                                   não executa replace do MYCONTROL com "0" 
*/                    PLC_What     : Z - zap 
*/                    PLN_IndexN   : número de indices a serem criados  
*/                    PLN_WhereN   : número de filtros a serem usados 
*/ Descrição        : Transforma qualquer arquivo ou Cursor em um Cursor EDITAVEL,  
*/                    acrescentado um campo de controle (CONTROL) 
*/ Retorno          : .T. - sucesso, .F. - c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_EditCursor
	LParameters PLC_Cursor, PLC_FromFile, PLC_ToFile, PLC_What, PLN_IndexN, PLN_WhereN

	LOCAL VLL_Ret

	this.FOL_PushSelect()

	VLL_Ret = .T.
	
	if this.FOL_ExecuteCursor(PLC_Cursor, PLC_FromFile, PLN_WhereN)
		if at("Z", PLC_What) > 0 and used(PLC_ToFile)
			zap in (PLC_ToFile)
		endif
		this.FOU_EditCursor(PLC_FromFile, PLC_ToFile, PLN_IndexN, PLC_Cursor)

		*	Armazena no form os cursores utilizados.  A variavel voo_loadingForm e ativada no LOAD dos forms
		*
		if !isnull(this.voo_loadingForm)
			
			xxx = createobject("form")
			
			with xxx
				.von_cursorlist = .von_cursorlist + 1
				dimension .voa_cursorlist(.von_cursorlist, 3)
				
				.voa_cursorlist(.von_cursorlist,1) = PLC_Cursor
				.voa_cursorlist(.von_cursorlist,2) = PLC_ToFile
				.voa_cursorlist(.von_cursorlist,3) = ""
			
			endwith
		endif


		use in (PLC_FromFile)
	else
		VLL_Ret = .F.
	endif				

	this.FOL_PopSelect()

	return VLL_Ret
ENDPROC




*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_CursorCatalog : Cursor a ser executado  
*/                    PLC_Cursor        : Cursor resultado  
*/                    PLN_WhereN        : Número de where's a serem usados na função  
*/ Descrição        : Executa um cursor definido no objeto  
*/ Retorno          : .T. : Se executou com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ExecuteCursor
lparameters PLC_CursorCatalog, PLC_Cursor, PLN_WhereN

LOCAL VLC_Name, VLC_Macro, VLL_Ret, VLC_Subst, VLN_ParCounter, VLN_Select, VLN_Connection, VLL_IsLocal, VLC_AliasCursor,;
		VLL_UseAdo, VLN_Options, VLN_LastSqlOptions
	
private VPA_Seek
dimension VPA_Seek[1]

this.FOL_PushSelect()
VLC_Name = upper(alltrim(PLC_CursorCatalog))
VLN_Options = 0

VLN_LastSql = ascan(this.VOA_LastSql, "<scs." + VLC_Name + ">", -1, -1, 1, 1+2+4+8)
*- Verifica se encontra a select na matriz de histórico

if VLN_LastSql = 0
	*- Não encontrou no histórico ou é select local

	VLL_Found = this.FOL_LookInCursor(VLC_Name)
	if VLL_Found
		VLL_IsLocal = y02.y02_005_l
		VLC_Macro = this.FOC_SelectField()
		VLN_Options = y02.y02_003_n
	else
		this.FOL_Triggers(.null., .null., "FOL_EXECUTECURSOR." + VLC_Name)
		if this.VON_TriggerReturnCount >= 2
			VLL_IsLocal = this.VOA_TriggerReturn[1]
			VLC_Macro = this.VOA_TriggerReturn[2]

			if alen(this.VOA_TriggerReturn) > 2
				if !empty(this.VOA_TriggerReturn[3])
					VLN_Options = this.VOA_TriggerReturn[3]
				endif
			endif

			VLL_Found = .t.
		endif
	endif

	VLC_Macro = this.FOC_AdjustSelect(VLC_Macro)

	if VLL_IsLocal
		*-	Retiro os nolock's
		VLC_Macro = strtran(VLC_Macro, "(NOLOCK)", "")
		*-	Retiro os STAR_DATA@
		VLC_Macro = strtran(VLC_Macro, "STAR_DATA@", "")
	
		if !empty(PLC_Cursor)
			VLC_Macro = VLC_Macro + " into cursor " + PLC_Cursor + " readwrite"
			
			
		endif
	else
		this.VON_LastSql = this.VON_LastSql + 1
		dimension this.VOA_LastSql(this.VON_LastSql, 2)
		this.VOA_LastSql(this.VON_LastSql, 1) = "<scs." + VLC_Name + ">"
		this.VOA_LastSql(this.VON_LastSql, 2) = ""
		
		if !empty(VLN_Options)
			this.VON_LastSqlOptions = this.VON_LastSqlOptions + 1
			dimension this.VOA_LastSqlOptions(this.VON_LastSqlOptions, 2)
			this.VOA_LastSqlOptions(this.VON_LastSqlOptions, 1) = "<scs." + VLC_Name + ">"
			this.VOA_LastSqlOptions(this.VON_LastSqlOptions, 2) = VLN_Options
		endif
	endif
else
	VLN_LastSqlOptions = ascan(this.VOA_LastSqlOptions, "<scs." + VLC_Name + ">", -1, -1, 1, 1+2+4+8)
	
	if VLN_LastSqlOptions > 0
		VLN_Options = this.VOA_LastSqlOptions(VLN_LastSqlOptions, 2)
	endif
	
	VLC_Macro = this.VOA_LastSql(VLN_LastSql, 2)
	VLL_Found = .t.
endif

if !empty(VLN_Options)
	*- 1 - Executa o comando utilizando ADO se estiver em MS-SQL Server
	*- 2 - Executa o comando utilizando ADO se estiver em Oracle
	*- 4 - Executa o comando utilizando ADO se estiver em IBM DB2
	VLL_UseAdo = (bittest(VLN_Options, 0) and this.VON_DBType = 1) or;
				(bittest(VLN_Options, 1) and this.VON_DBType = 2) or;
				(bittest(VLN_Options, 2) and this.VON_DBType = 3)
endif

if VLL_Found
	if PLN_WhereN > 0
		dimension VPA_Seek[PLN_WhereN]
		*- Inicializa as variáveis para o select
		for VLN_ParCounter = 1 to PLN_WhereN
			VPA_Seek[VLN_ParCounter] = this.VOA_Where[VLN_ParCounter]
		endfor
	endif

	*- Verifica se o select é local
	if VLL_IsLocal
		*- Executa a macro
		this.FOL_SelectLocal(VLC_Macro)
		VLL_Ret = used(PLC_Cursor)
	else	
		this.VON_ParCountEC = PLN_WhereN
		if !VLL_UseAdo
			VLL_Ret = this.FOL_SqlExec(VLC_Macro, PLC_Cursor, (VLN_LastSql > 0))
		else
			VLL_Ret = this.FOL_AdoSqlExec(VLC_Macro, PLC_Cursor, (VLN_LastSql > 0))
		endif
	endif

	if used(PLC_Cursor)
		go top in (PLC_Cursor)
	endif

endif

this.FOL_PopSelect()

return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Id    : Id a ser procurado  
*/ Descrição        : Procura uma expressão no catalog.sys  
*/ Retorno          : .T. : Achou, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_LookinCursor
	LParameters PLC_Id

	LOCAL VLL_Ret

	if empty(PLC_Id)
		PLC_Id = ""
	endif

	*- Y02 é o arquivo padrão de cursores
	*- se não estiver aberto, abre o catalog.sys, e procura PLC_TableName nele

	this.FOL_UseInPackage("y02", "y02")

	VLL_Ret = seek(PLC_Id, "y02", "id")

	return VLL_Ret
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias : O arquivo que deve ser acrescentado um novo registro  
*/                    PLC_Ukey  : Indica o ukey do arquivo  
*/ Descrição        : Insere um novo registro no arquivo PLC_Alias colocando o Ukey 
*/ Retorno          : .T. : Se sucesso, .F. - c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_NewRecordtoForm
	LParameters PLC_Alias, PLC_Ukey

	LOCAL VLN_Fields
	LOCAL array VLA_Fields[1]

	select (PLC_ALias)
	*- Adiciono um novo registro
	append blank in (PLC_ALias)
	if empty(PLC_Ukey)
*---	se o alias tiver um UKEY com menos posições pego as últimas posições
		replace UKEY with right(this.FOC_GetUkey(), len(UKEY)) in (PLC_Alias)
	else
		replace UKEY with PLC_UKey in (PLC_Alias)
	endif

	replace TIMESTAMP with datetime(), CIA_UKEY with this.VOC_CompanyCode in (PLC_Alias)

	return .t.
ENDPROC


*-----------------------------------------------------------------------------------------------------------------------
*- Criptografa e Descriptografa uma string baseado em uma chave de codificação
*- Parâmetros: 	- PLC_String      : String que será decodificada
*- Retorno   : String decodificada
*- Denis Carrizo em 10/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_Encrypt
lparameters PLC_String, PLC_EncryptKey
	return this.VOO_Security.FOC_Encrypt(PLC_String, PLC_EncryptKey)
endproc


*-----------------------------------------------------------------------------------------------------------------------
*- Decodifica uma string baseado em um algoritmo padrão
*- Parâmetros: 	- PLC_String      : String que será decodificada
*- Retorno   : String decodificada
*- Denis Carrizo em 10/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_Decode
lparameters PLC_String
	return this.VOO_Security.FOC_Decode(PLC_String)
endproc


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Propertie : Propriedade para ser procurada na linha de parametros
*/                    
*/ Descrição        : Procura uma determinada propriedade na linha de parametros inicial
*/ Retorno          : a propriedade : Sucesso,   "" : c.c.  
*/ Denis 11/09/2001
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_StartPar
lparameters PLC_Propertie

local VLC_String
PLC_Propertie = upper(PLC_Propertie)

if !empty(this.voc_StartPar) and PLC_Propertie $ this.voc_StartPar
	
	VLC_String = substr(this.voc_StartPar, at(PLC_Propertie,this.voc_StartPar )+len(PLC_Propertie) +1) + ";"
	VLC_String = substr(VLC_String, 1, at(";", VLC_String)-1 )
	
	return VLC_String
	
else
	return ""
endif

endproc


*-----------------------------------------------------------------------------------------------------------------------
*- Carrega uma array com os identificadores das tabelas do sistema para a substituição do STAR_DATA@
*- Parâmetros - PLC_CiaUkey : Ukey da empresa que será carregada
*- Retorno   : .T. se conseguir carregar, .F. c.c.
*- Denis Carrizo em 11/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOL_LoadTblId
lparameters PLC_CiaUkey
	return this.VOO_Security.FOL_LoadTblId(PLC_CiaUkey)
endproc


*-----------------------------------------------------------------------------------------------------------------------
*- Retorna a ukey de uma tabela no arquivo sys07
*- Parâmetros: 	- PLC_CiaUkey  : Ukey da empresa
*-				- PLC_TblId    : Id padrão do Arquivo
*- Retorno   : Ukey do arquivo
*- Denis Carrizo em 11/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_TableUkey
lparameters PLC_CiaUkey, PLC_TblId

VLC_Alias = this.FOC_TableIdent("SYS10") + "SYS10" && Identificador da tabela

VLC_Select = "select sys07_ukey from " + VLC_Alias + " where tab_id = '" + PLC_TblId +;
			 "' and cia_ukey = '" + PLC_CiaUkey + "'"

VLL_SqlRet = this.FOL_SqlExec(VLC_Select, "tmpcursor")

VLC_TblUkey = ""

if VLL_SqlRet
	select tmpcursor
	if recc("tmpcursor") > 0
		VLC_TblUkey = tmpcursor.sys07_ukey
	endif
	use in tmpcursor
endif

return VLC_TblUkey
endproc


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_DataBaseName : O nome do database a ser aberto  
*/ Descrição   		: Abre o database PLC_DataBaseName  
*/ Retorno     		: .T. : Abriu com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_OpenDatabase
	LParameters PLC_DataBaseName

	*- Abre o Database PLC_DataBaseName.Se não existe cria
	if !file(alltrim(PLC_DataBaseName)+".DBC")
		this.FOL_MakeDir(PLC_DataBaseName, .T.)
		create database (PLC_DataBaseName)
	endif
	open database (PLC_DataBaseName)

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Cursor : Arquivo de Origem  
*/ Descrição        : Executa um pack de um cursor  
*/ Retorno          : .T. : Ok, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_PackCursor
	LParameters PLC_Cursor

	LOCAL VLN_Fieldcount, VLC_Macro, VLA_FieldStru[1], VLC_Deleted

	VLC_Deleted = set("DELETED")

	this.FOL_PushSelect()

	*- Acessa o alias Solicitado (deve ser aberto previamente)
	select (PLC_Cursor)

	*- Lê sua estrutura
	VLN_Fieldcount = afields(VLA_FieldStru)

	*- Cria fisicamente o Cursor
	create cursor __PACK from array VLA_FieldStru



	*- Mantêm deleted on
	set deleted on

	*- Copia os registros para __PACK e depois copia de volta
	select __PACK
	append from dbf(PLC_Cursor)
	select (PLC_Cursor)
	zap
	append from dbf("__PACK")

	*- Fecha o arquivo __PACK
	use in __PACK

	*- Volta o deleted
	set deleted &VLC_Deleted

	this.FOL_PopSelect()

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLL_LeaveRecNo : Deixa o arquivo posicionado no registro atual  
*/ Descrição        : Retira da pilha a area e a indexação  
*/ Retorno          : .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_PopSelect
	LParameters PLL_LeaveRecNo

	*- verifica se a sessão de dados é igual à que foi empilhada
	if set("datasession") <> this.VOA_StackAlias[this.VON_StackAlias, 4]
		this.FOL_SetDataSession(this.VOA_StackAlias[this.VON_StackAlias, 4])
	endif

	*- Se a pilha não estiver vazia, seleciona a área do topo, ordem, e registro
	if !empty(this.VOA_StackAlias[this.VON_StackAlias, 1])
		select (this.VOA_StackAlias[this.VON_StackAlias, 1])

	*--	se não for SQL não posso reposicionar já que não sei qual cursor está ativo
		if !empty(alias())
			if !empty(this.VOA_StackAlias[this.VON_StackAlias, 2])
				if !empty(this.VOA_StackAlias[this.VON_StackAlias, 5])
					set order to (this.VOA_StackAlias[this.VON_StackAlias, 2]) descending
				else
					set order to (this.VOA_StackAlias[this.VON_StackAlias, 2]) ascending
				endif
			endif

			if !PLL_LeaveRecNo
				if recno() <> this.VOA_StackAlias[this.VON_StackAlias, 3]
					if !empty(this.VOA_StackAlias[this.VON_StackAlias, 3]) and between(this.VOA_StackAlias[this.VON_StackAlias, 3], 1, reccount())
						goto (this.VOA_StackAlias[this.VON_StackAlias, 3])
					else
						go bottom
						if !eof()
							skip
						endif
					endif
				endif
			endif
		endif
	endif

	*- Decrementa o topo
	this.VON_StackAlias = this.VON_StackAlias - 1

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : Nenhum 
*/ Descrição        : Guarda na pilha a area corrente 
*/ Retorno          : .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_PushSelect

	local VLN_ArraySize

	*- Incrementa o topo e guarda a área, tag, e número do registro corrente
	this.VON_StackAlias = this.VON_StackAlias + 1
	VLN_ArraySize = alen(this.VOA_StackAlias, 1)

	if VLN_ArraySize < this.VON_StackAlias
		VLN_ArraySize = VLN_ArraySize + 300
		dimension this.VOA_StackAlias[VLN_ArraySize, 5]
	endif
	
	this.VOA_StackAlias[this.VON_StackAlias, 1] = alias()
	if !empty(alias())
		this.VOA_StackAlias[this.VON_StackAlias, 2] = tagno()
		this.VOA_StackAlias[this.VON_StackAlias, 5] = descending()
		=reccount()					&& Gambiarra para atualizar o recno()
		this.VOA_StackAlias[this.VON_StackAlias, 3] = iif( eof(), 0, recno() )
	endif
	this.VOA_StackAlias[this.VON_StackAlias, 4] = set("datasession")

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Field : Campo a ser dado o replace  
*/                    PLC_What  : Indica o que deve ser dado como replace 
*/ Descrição        : Executa um replace do valor PLC_What no campo PLC_Field 
*/ Retorno          : .T. : Se sucesso, .F. - c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ReplaceValue
	LParameters PLC_Field, PLC_What

	LOCAL  VLN_i, VLC_AuxAlias

	*- se existe determinação do alias então executa replace "in"
	VLN_i = at(".", PLC_Field)
	if VLN_i > 0
		VLC_AuxAlias = subs(PLC_Field, 1, VLN_i - 1)
		replace (PLC_Field) with PLC_What in (VLC_AuxAlias)
	else
		replace (PLC_Field) with PLC_What
	endif

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Form      : Indica em qual form é executada o cancelamento  
*/                    PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/                    PLC_Which     : Guarda um campo para o cancel update  
*/                    PLL_Close     : Fecha o arquivo Base no Final 
*/                    PLL_OnlyLoop  : Indica se irá executar apenas o loop  
*/                    PLL_ToIsTable : Indica que o destino é uma tabela física  
*/ Descrição        : Atualiza a informacao do cursor editavel 'From' no arquivo Pai 'To' 
*/ Retorno          : Nenhum 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SaveCursor
	LParameters PLO_Form, PLC_FromFile, PLC_ToFile, PLC_FieldUkey, PLC_Which, PLL_Close, ;
				PLL_OnlyLoop, PLL_ToIsTable

	LOCAL VLL_Found, VLL_Ret, VLU_Ukey, VLC_Status

	*- se não está aberto o arquivo fonte retorna erro

	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	*this.FOL_UpdatingErrMsg(PLC_ToFile)

	this.FOL_PushSelect()

	*- Muda o Status do DELETED para garantir uma Correcao adequada
	VLC_OldDELETED = set("DELETED")
	set deleted off

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
	select (PLC_FromFile)
	this.FOL_PushSelect()

	set order to MYCONTROL

	VLL_Ret = .T.

*--	executa um scan com MYCONTROL = "1"
	=seek("1")
	scan while mycontrol = "1"

*---	Verifica se existe no arquivo Original
		VLU_Ukey   = eval(PLC_FieldUkey)
		if PLL_ToIsTable
			VLC_Status = eval("status")
		else
			VLC_Status = " "
		endif

*---	Verifica se existe no arquivo Original
		VLL_Found = seek(VLU_Ukey, PLC_ToFile, "UKEY")

*---	verifica se está deletado e estava gravado
		if deleted(PLC_FromFile)
			if VLC_Status = "W" or !PLL_ToIsTable
				if !eof(PLC_ToFile)
					if !isnull(PLO_Form)
						if !PLO_Form.FOL_CancelUpdates(PLC_Which)
							VLL_Ret = .F.
							exit
						endif
					endif

*-----				move o conteúdo do fonte para o destino
					select (PLC_FromFile)
					scatter memvar memo

					select (PLC_ToFile)
					gather memvar memo
					replace mycontrol with "1" in (PLC_ToFile)
					delete in (PLC_ToFile)
				endif
			endif
*---	se não está deletado
		else
*----		verifica se não foi adicionado
			if PLL_ToIsTable and VLC_Status = "W"
*-----			e não está no destino
				if eof(PLC_ToFile)
					this.VON_Error = SV_FatalErr
					VLL_Ret = .F.
					exit
				endif
*-----			e os timestamps são distintos
				if eval(PLC_ToFile + ".timestamp") <> eval("timestamp")
					this.VON_Error = SV_FatalErr
					this.VON_SQLError = -1 && Timestamp diferente
					VLL_Ret = .F.
					exit
				endif
			endif

*----		Se o Form não for .NULL. executa o método FOL_SaveUpDates
			if !isnull(PLO_Form)
				if !PLO_Form.FOL_SaveUpDates(PLC_Which, VLL_Found)
					VLL_Ret = .F.
					exit
				endif
			endif

			if !PLL_OnlyLoop

*-----			se está em EOF ou foi adicionado
				if eof(PLC_ToFile)
					append blank in (PLC_ToFile)
				endif

*-----			move o conteúdo do fonte para o destino
				select (PLC_FromFile)
				scatter memvar memo

				select (PLC_ToFile)
				gather memvar memo
				if PLL_ToIsTable
					replace timestamp with datetime()
					replace status with "W" + subs(status,2)
				endif
			endif

		endif

*---	atualiza
		if !PLL_OnlyLoop
			if !this.FOL_TableUpdate(PLC_ToFile, .F.)
				this.VON_Error = SV_NoFatalErr
				VLL_Ret = .F.
				exit
			endif
		endif
	endscan

	go top in (PLC_FromFile)


	this.FOL_PopSelect()

	*- Fecha o arquivo origem se necessário
	if PLL_Close		&& Fecha o arquivo base no Final
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina antes de fechar
		this.FOL_CloseTable(PLC_FromFile)
	else
	*--	Executa um pack e volta todos os mycontrol com '0'
		select (PLC_FromFile)
		this.FOL_PackCursor(PLC_FromFile)
	*	replace all MYCONTROL with "0" in (PLC_FromFile)
		go top in (PLC_FromFile)
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina
	endif

	if VLC_OldDELETED = "ON"
		set deleted on
	endif

	return VLL_Ret
ENDPROC


*-----------------------------------------------------------------------------------------------------------------------
*- Prepara uma string para ser executada no servidor de banco de dados
*- Parâmetros:	- PLC_SQLCmd		: Comando sql que será enviado
*-				- PLL_Treated		: Indica se o comando SQL já foi rodado alguma vez (resolvido), matriz VOA_LastSql
*- Retorno   : string pronta para ser executada
*- Denis Carrizo em 19/07/2004
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_SqlStrToExec
lparameters PLC_SQLCmd, PLL_Treated, PLL_SolveVarsInSqlCmd

	local VLC_String, VLL_DebugMode, VLC_VarName

	if empty(PLL_Treated)
		VLC_String = this.FOC_ChangeString(PLC_SQLCmd)
		if this.VON_LastSql >= 1 and empty(this.VOA_LastSql(this.VON_LastSql, 2))
			*- Não encontrou no histórico, pego o último
			this.VOA_LastSql(this.VON_LastSql, 2) = VLC_String
		endif
	else
		VLC_String = PLC_SQLCmd
	endif

	*- Propriedade que indica se as variáveis devem ser trocadas
	if !empty(PLL_SolveVarsInSqlCmd)
		*- A FOC_ChangeString garante o case como maiusculo

		if this.VON_ParCountEC > 0
			for VLN_VarCounter = 1 to this.VON_ParCountEC
				VLC_VarName = "VPA_SEEK["+ alltrim(str(VLN_VarCounter)) + "]"
				if VLC_VarName $ VLC_String
					VLU_VarContent = VPA_Seek[VLN_VarCounter]

					VLC_Type = type(VLC_VarName)
					VLC_Trans = ""

					do case
						case isnull(VLU_VarContent)
							VLC_Trans = "NULL"
						case VLC_Type = "C"
						 	if left(ltrim(VLU_VarContent), 1) = "'" and right(rtrim(VLU_VarContent), 1) = "'"
	 						 	VLC_Trans = VLU_VarContent
	 						 else
							 	VLC_Trans = "'" + VLU_VarContent + "'"
							endif
						case VLC_Type = "D"
						 	VLC_Trans = "'" + dtoc(VLU_VarContent) + "'"

						case VLC_Type = "T"
						 	VLC_Trans = "'" + ttoc(VLU_VarContent) + "'"

						case VLC_Type = "N"
						 	VLC_Trans = strtran(transform(VLU_VarContent), set("point"), ".")
					endcase

			   		VLC_String = strtran(VLC_String, "?" + VLC_VarName, VLC_Trans)
			   	endif
		   	endfor
		   	this.VON_ParCountEC = 0
		endif
		
		*- Tratamento do VPA_Filter
		if this.VON_ParCountSearch > 0
			for VLN_VarCounter = 1 to this.VON_ParCountSearch
				VLC_VarName = "VPA_FILTER["+ alltrim(str(VLN_VarCounter)) + "]"
				if VLC_VarName $ VLC_String
					VLU_VarContent = VPA_Filter[VLN_VarCounter]

					VLC_Type = type(VLC_VarName)
					VLC_Trans = ""

					do case
						case isnull(VLU_VarContent)
							VLC_Trans = "NULL"
						case VLC_Type = "C"
						 	if left(ltrim(VLU_VarContent), 1) = "'" and right(rtrim(VLU_VarContent), 1) = "'"
	 						 	VLC_Trans = VLU_VarContent
	 						 else
							 	VLC_Trans = "'" + VLU_VarContent + "'"
							endif
						case VLC_Type = "D"
						 	VLC_Trans = "'" + dtoc(VLU_VarContent) + "'"

						case VLC_Type = "T"
						 	VLC_Trans = "'" + ttoc(VLU_VarContent) + "'"

						case VLC_Type = "N"
						 	VLC_Trans = strtran(transform(VLU_VarContent), set("point"), ".")
					endcase

			   		VLC_String = strtran(VLC_String, "?" + VLC_VarName, VLC_Trans)
			   	endif
		   	endfor
		   	
			this.VON_ParCountSearch = 0 
		endif
	endif

	*- Tratamento para substituiçaão de expressões em comandos SQL
	if at("<<", VLC_String) > 0 and at(">>", VLC_String) > 0
		VLC_String = textmerge(VLC_String, .f., "<<", ">>")
	endif

	return VLC_String

endproc

*-----------------------------------------------------------------------------------------------------------------------
*- Executa um comando sql na conexão padrão
*- Parâmetros:	- PLC_SQLCmd		: Comando sql que será enviado
*-				- PLC_ReturnCursor	: Nome do cursor caso haja retorno
*-				- PLL_Treated		: Indica se o comando SQL já foi rodado alguma vez (resolvido), matriz VOA_LastSql
*- Retorno   : .T. se executado com sucesso, .F. c.c.
*- Denis Carrizo em 19/07/2004
*-----------------------------------------------------------------------------------------------------------------------
procedure FOL_SqlExec
lparameters PLC_SQLCmd, PLC_ReturnCursor, PLL_Treated

local VLN_Ret, VLC_SqlCmd, VLC_String, VLL_DebugMode

if !(this.VON_CnxHandle > 0 and this.VON_CnxStatus > 0)
	return .f.
endif

if empty(PLC_ReturnCursor)
	PLC_ReturnCursor = ""
endif

VLN_Ret = 0
VLC_SqlCmd = ""
VLC_String = ""

VLL_DebugMode = !empty(this.VOC_DebugSQLCmdFile) and (bitand(Debug_SqlErr, this.VON_DebugType) = Debug_SqlErr)

if VLL_DebugMode
	if empty(justext(this.VOC_DebugSQLCmdFile))
		this.VOC_DebugSQLCmdFile = alltrim(this.VOC_DebugSQLCmdFile) + ".dbf"
	endif

	if !file(this.VOC_DebugSQLCmdFile)
		create table (this.VOC_DebugSQLCmdFile) (start t, elap_time b(18), sql_cmd m, status c(4), debug_type n)
		this.FOL_CloseTable(juststem(this.VOC_DebugSQLCmdFile))
	endif
	VLT_Start = datetime()
	VLN_StartSec = seconds()
endif

VLC_String = this.FOC_SqlStrToExec(PLC_SQLCmd, PLL_Treated, this.VOL_SolveVarsInSQLCmd or VLL_DebugMode)

do while VLN_Ret = 0
	VLN_Ret = sqlexec(this.VON_CnxHandle, VLC_String, PLC_ReturnCursor)
	if VLN_Ret = 0
		=inkey(1, "H") && Para não ter tantas chamadas na função
	endif
enddo

VLL_Return = VLN_Ret > 0

*- Código para debugar velocidade de execução
if VLL_DebugMode
	VLN_ElapsedSec = seconds() - VLN_StartSec
	VLC_Status = iif(VLL_Return, " OK ", "ERRO")

	use (this.VOC_DebugSQLCmdFile) alias DEBUGFILE in 0 again
	append blank in debugfile
	replace debugfile.start		 with VLT_Start		,;
			debugfile.elap_time	 with VLN_ElapsedSec,;
			debugfile.sql_cmd	 with VLC_String	,;
			debugfile.status	 with VLC_Status	,;
			debugfile.debug_type with Debug_SqlErr	in DEBUGFILE
	use in debugfile
endif

if !VLL_Return
	this.FOL_SetError()
else
	if this.VON_DBType <> 1 and this.VON_TranLevel <= 0 && Tratamento de transações do DB2
		sqlexec(this.VON_CnxHandle, "commit")
	endif
endif

return VLL_Return
endproc


*----------------------------------------------------------------------------------------
* Função Objeto	: FOL_SetError
* Descrição 	: Seta o erro do comando sql.
* Retorno     	: sempre .t.
* Denis Carrizo 11/09/2001
*----------------------------------------------------------------------------------------
procedure FOL_SetError

LOCAL VLA_Error[1], VLN_i, VLC_Return, VLC_Lower, VLC_StringError, VLN_ErrorNo

aerror(VLA_Error)
VLN_ErrorNo 		= VLA_Error[5]
this.VOC_SqlError 	= iif(this.VON_DBType = 3, VLA_Error[3], VLA_Error[2])

do case
	*- Violação de FK, passo o nome da FK violada em .VOC_ErrorPar
	case (this.VON_DBType = 1 and VLN_ErrorNo = 547) or; && MS-SQL Server
		 (this.VON_DBType = 2 and VLN_ErrorNo = 2292) or; && Oracle
		 (this.VON_DBType = 3 and VLN_ErrorNo = -532) && Oracle

		this.VON_SQLError = -9
		this.VON_Error = -1
		this.VOC_ErrorPar = ""

		VLN_i = at("FK_", this.VOC_SqlError) && As nossas FK´s sempre começam com esse prefixo
		if VLN_i > 0
			VLC_Par = subs(this.VOC_SqlError, VLN_i)
			do case
				case this.VON_DBType = 1 && No SQL, o nome da FK fica entre aspas simples
					VLC_EndChar = "'"
				case this.VON_DBType = 2 && No Oracle, o nome da FK fica entre parênteses
					VLC_EndChar = ")"
				case this.VON_DBType = 3 && No DB2, o nome da FK fica entre aspas duplas
					VLC_EndChar = '"'
			endcase
			this.VOC_ErrorPar = upper(subs(VLC_Par, 1, at(VLC_EndChar, VLC_Par) - 1))
		endif

	*- Chave única violada, passo o nome do índice único violado em .VOC_ErrorPar
	case (this.VON_DBType = 1 and (VLN_ErrorNo = 2601 or VLN_ErrorNo = 2627)) or; && MS-SQL Server
		 (this.VON_DBType = 2 and VLN_ErrorNo = 1) or; && Oracle
		 (this.VON_DBType = 3 and VLN_ErrorNo = -803) && DB2

		this.VON_SQLError = -7
		this.VON_Error = -1

		do case
			case this.VON_DBType = 1 && No SQL, o nome do índice fica entre aspas simples
				VLN_i = at("I_", this.VOC_SqlError) && Os nossos índices sempre começam com esse prefixo
				if VLN_i > 0
					VLC_Par = subs(this.VOC_SqlError, VLN_i)
					VLC_EndChar = "'"
					this.VOC_ErrorPar = upper(subs(VLC_Par, 1, at(VLC_EndChar, VLC_Par) - 1))
				endif

			case this.VON_DBType = 2 && No Oracle, o nome do índice fica entre parênteses
				VLN_i = at("I_", this.VOC_SqlError) && Os nossos índices sempre começam com esse prefixo
				if VLN_i > 0
					VLC_Par = subs(this.VOC_SqlError, VLN_i)
					VLC_EndChar = ")"
					this.VOC_ErrorPar = upper(subs(VLC_Par, 1, at(VLC_EndChar, VLC_Par) - 1))
				endif

			case this.VON_DBType = 3 && No DB2, o nome do índice não é referenciado na mensagem, apenas o número e a tabela
				VLC_IndexId = strextract(this.VOC_SqlError, ["], ["], 1)
				VLC_FullTableName = strextract(this.VOC_SqlError, ["], ["], 3)

				if !empty(VLC_IndexId) and !empty(VLC_FullTableName)
					VLN_PointPos = at(".", VLC_FullTableName)
					VLC_TableName = substr(VLC_FullTableName, VLN_PointPos + 1)
					VLC_SchemaName = substr(VLC_FullTableName, 1, VLN_PointPos - 1)
					VLC_SQLCmd = "select indname from syscat.indexes where tabschema = '" + VLC_SchemaName + "' and " +;
								"tabname = '" + VLC_TableName + "' and iid = " + VLC_IndexId + " for fetch only"
					if this.FOL_SqlExec(VLC_SQLCmd, "tmp_idx")
						if !eof("tmp_idx")
							this.VOC_ErrorPar = alltrim(tmp_idx.indname)
						endif
					endif
					
					if used("tmp_idx")
						use in tmp_idx
					endif
				endif
		endcase		


	*- DeadLock gerado
	case (this.VON_DBType = 1 and VLN_ErrorNo = 1205) or; && MS-SQL Server
		 (this.VON_DBType = 2 and VLN_ErrorNo = 60)  && Oracle
		this.VON_SQLError 	= -8
		this.VON_Error 		= -1
		this.VOC_ErrorPar 	= ""
		*- Denis 04/06/2002
		*- O DeadLock fecha as transações abertas, logo não posso efetuar o rollback
		this.VON_TranLevel	= 0
		*-----

	*- Espaço no Log ou database
	case this.VON_DBType = 1 and inlist(VLN_ErrorNo, 1703, 9002) && MS-SQL Server
		this.VON_SQLError 	= -10
		this.VON_Error 		= -1
		this.VOC_ErrorPar 	= ""
		*- Denis 11/12/2002
		*- A falta de espaço fecha as transações abertas, logo não posso efetuar o rollback
		this.VON_TranLevel	= 0
		*-----

	*- Tempo de timeout excedido
	case this.VON_DBType = 1 and VLN_ErrorNo = 1222
		this.VON_SQLError 	= -6
		this.VON_Error 		= 1
		this.VOC_ErrorPar 	= ""

*------------------------------------------------------------------------------------------------------------------------
	otherwise
		this.VON_SqlError 	= VLN_ErrorNo
		this.VON_Error 		= -1
		this.VOC_ErrorPar	= ""
*------------------------------------------------------------------------------------------------------------------------
endcase

return .T.
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias : O arquivo que deve ser revertido  
*/                    PLL_All   : Se deve ser todo revertido  
*/ Descrição        : Reverte o arquivo PLC_Alias 
*/ Retorno          : .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_TableRevert
	LParameters PLC_Alias, PLL_All

	LOCAL  VLL_Ret

	this.FOL_PushSelect()

	select (PLC_Alias)
	VLL_Ret = tablerevert(.T.)

	this.FOL_PopSelect()
	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_TableSource - O nome da Tabela Original 
*/              	: PLC_TableTarget - O nome do Cursor a ser gerado 
*/					: PLL_NotClose    - Se fecha a Tabela a Principal 
*/               	: PLC_IndexOn     - String de indexação, no caso de mais de um índice 
*/               	:  separar com ;      								  
*/                  : PLC_Tag        - Nome do Tag, separados por ; 
*/                  : PLC_TagType    - Tipo do tag, separado por ;  
*/ Descrição   		: Cria um cursor temporário com as características da tabela principal
*/ Retorno     		: .T. : Criou com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_TabletoCursor
	LParameters PLC_TableSource, PLC_TableTarget, PLL_NotClose, PLC_IndexOn, PLC_Tag, PLC_TagType

	this.FOL_PushSelect()

	LOCAL VLN_Position, VLC_IndexOn, VLC_Tag, VLC_TagType

	*--	Cursor temporário para Ficha de Atendimento
	select (PLC_TableSource)
	afields(VLA_Parent)
	if ascan(VLA_Parent, "MYCONTROL") = 0
		VLN_New = alen(VLA_Parent,1) + 1
		dimension VLA_Parent(VLN_New, AFIELDS2NDCOL)
		VLA_Parent(VLN_New, 1) = "MYCONTROL"
		VLA_Parent(VLN_New, 2) = "C"
		VLA_Parent(VLN_New, 3) = 1
		VLA_Parent(VLN_New, 4) = 0
		store "" to VLA_Parent(VLN_New, 7), VLA_Parent(VLN_New, 8),  VLA_Parent(VLN_New, 9), ;
					VLA_Parent(VLN_New, 10), VLA_Parent(VLN_New, 11), VLA_Parent(VLN_New, 12), ;
					VLA_Parent(VLN_New, 13), VLA_Parent(VLN_New, 14), VLA_Parent(VLN_New, 15), ;
					VLA_Parent(VLN_New, 16)
	endif


	if !PLL_NotClose
		this.FOL_CloseTable(PLC_TableSource)
	endif

	VLC_IndexOn = "UKEY;MYCONTROL" + iif(!empty(PLC_IndexOn), ";"+PLC_IndexOn, "")
	VLC_Tag		= "UKEY;MYCONTROL" + iif(!empty(PLC_Tag), ";"+PLC_Tag, "")
	VLC_TagType = ";" + iif(!empty(PLC_TagType), ";"+PLC_TagType, "")

	this.FOL_PopSelect()

	return this.FOL_CreateTable(PLC_TableTarget, "", "CURSOR", VLC_IndexOn, VLC_Tag, VLC_TagType, @VLA_Parent)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias : O arquivo que deve ser atualizado 
*/                    PLL_All   : Se deve ser todo atualizado 
*/ Descrição        : Atualiza o arquivo PLC_Alias  
*/ Retorno          : .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_TableUpdate
	LParameters PLC_Alias, PLL_All

	LOCAL  VLL_Ret, VLN_Connection, VLN_Select

	this.FOL_PushSelect()

	VLL_Ret = .F.
	select (PLC_Alias)
	VLN_Select = select()
	VLL_Ret = tableupdate(0, .F.)

	this.FOL_PopSelect()
	return VLL_Ret
ENDPROC


*----------------------------------------------------------------------------------------
* Função Objeto	: FON_Tagarray
* Parâmetros   	: PLC_Alias			- Tabela a ser avaliada
*                 PRA_TagArray		- Matriz referência que irá armazenar os indices da tabela avaliada
* Descrição 	: Preenche a matriz referência com os índices únicos e FK´s da tabela indicada
* Retorno     	: Número de índices que a tabela possui
*----------------------------------------------------------------------------------------
PROCEDURE FON_TagArray
lparameters PLC_Alias, PRA_TagArray

local VLC_TableId, VLN_IdxCount, VLC_Sys07, VLC_Sys08, VLC_Sys10

VLC_TableId = upper(padr(PLC_Alias, 20))
VLC_CiaUkey = this.VOC_CompanyCode

VLC_Sys07 = this.FOC_TableIdent("SYS07") + "SYS07" && Identificador da tabela SYS07
VLC_Sys08 = this.FOC_TableIdent("SYS08") + "SYS08" && Identificador da tabela SYS08
VLC_Sys10 = this.FOC_TableIdent("SYS10") + "SYS10" && Identificador da tabela SYS10

*- Índices da tabela solicitada, incluindo suas FK´s
VLC_Select = "select sys07.tab_id, sys07.tab_loc, sys08.* from " + VLC_Sys08 + " sys08, " + VLC_Sys07 + " sys07, " +;
		VLC_Sys10 + " sys10" + " where sys08.sys07_ukey = sys07.ukey and sys10.sys07_ukey = sys07.ukey and sys10.cia_ukey = '" +;
		VLC_CiaUkey + "' and sys07.tab_id = '" + VLC_TableId + "' and (sys08.idx_type = 1 or sys08.idx_type = 3)"

VLL_SqlRet = this.FOL_SqlExec(VLC_Select, "srv_index")

if !VLL_SqlRet
	return -1
endif

VLN_IdxCount = 0

select srv_index
scan
	if srv_index.idx_type = 9 && FK
		VLC_RefTable 	= this.FOC_TagContent(srv_index.idx_exp, "REFTABLE", 1)
		VLC_FKey 		= this.FOC_TagContent(srv_index.idx_exp, "FKEY", 1)
		VLC_RefExpr		= this.FOC_TagContent(srv_index.idx_exp, "REFEXPR", 1)
	else
		VLC_RefTable 	= ""
		VLC_FKey 		= ""
		VLC_RefExpr		= ""
	endif

	VLN_IdxCount = VLN_IdxCount + 1
	dimension PRA_TagArray[VLN_IdxCount, 7]
	PRA_TagArray[VLN_IdxCount, 1] = alltrim(srv_index.idx_name)
	PRA_TagArray[VLN_IdxCount, 2] = alltrim(srv_index.idx_exp)
	PRA_TagArray[VLN_IdxCount, 3] = srv_index.idx_type
	PRA_TagArray[VLN_IdxCount, 4] = VLC_RefTable
	PRA_TagArray[VLN_IdxCount, 5] = VLC_FKey
	PRA_TagArray[VLN_IdxCount, 6] = VLC_RefExpr
	PRA_TagArray[VLN_IdxCount, 7] = alltrim(srv_index.tab_id)
endscan
use in srv_index

return VLN_IdxCount
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_TableName : O nome do table a ser aberto  
*/              	  PLC_Alias     : Indica o alias que deseja abrir 
*/                    PLC_Tag       : Indexação a ser usada 
*/                    PLL_Again     : Se abrirá com again 
*/                    PLL_NoUpdate  : Se abrirá com noupdate  
*/                    PLL_In0       : Se abrirá na área 0 
*/                    PLL_Exclusive : Se abrirá o arquivo de modo exclusivo 
*/ Descrição   		: Abre o table PLC_TableName. 
*/ Retorno     		: .T. : Abriu com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_UseTable
	LParameters PLC_TableName, PLC_Alias, PLC_Tag, PLL_Again, PLL_NoUpdate, PLL_In0, PLL_Exclusive

	LOCAL VLC_Macro, VLL_Ret

	VLC_Macro = "use '" + alltrim(PLC_TableName) + "'" + iif(!empty(PLC_Alias), " alias "+PLC_Alias, "")+ ;
				iif(PLL_Again, " again ", "")+;
				iif(!empty(PLC_Tag), " order "+ PLC_Tag + " ", "")+ ;
				iif(PLL_NoUpdate, " noupdate ", "")+iif(PLL_In0, " in 0 ", "")+ ;
				iif(PLL_Exclusive, " exclusive ", " shared ")

	VLL_Ret = .T.
	on error VLL_Ret = .F.
	&VLC_Macro
	on error
	with this
		if VLL_Ret
			this.FOL_CursorBuffering(3, select(PLC_Alias))
		endif
	endwith

	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FromFile : Arquivo de Origem  
*/                    PLC_ToFile   : Cursor gravavel de destino, se o nome iniciar com "-"
*/                                   não executa replace do MYCONTROL com "0" 
*/                    PLN_IndexN   : número de índices a serem criados  
*/                    PLC_Cursor   : Nome do cursor do Y02
*/ Descrição        : Transforma qualquer arquivo ou Cursor em um Cursor EDITAVEL,  
*/                    acrescentado um campo de controle (CONTROL) 
*/ Retorno          : RECC() do arquivo Criado 
*/----------------------------------------------------------------------------------------
PROCEDURE FOU_EditCursor
	LParameters PLC_FromFile, PLC_ToFile, PLN_IndexN, PLC_Cursor

	LOCAL VLC_Macro, VLN_Fieldcount, VLA_FieldStru[1], VLN_Select, VLL_Replace, VLN_i, VLC_Aux, ;
		  VLC_PreviousDeleted, VLA_Array, VLN_Cont, VLN_Tot, VLL_HasBeenCreated

	this.FOL_PushSelect()

	*- Acessa o alias Solicitado (deve ser aberto previamente)
	select (PLC_FromFile)

	VLL_Replace = .T.
	if PLC_ToFile = "-"
		PLC_ToFile = subs(PLC_ToFile, 2)
		VLL_Replace = .F.
	endif


	if this.VON_DBType > 1 and !empty(PLC_Cursor) and this.FOL_LookInCursor(PLC_Cursor)
		VLC_String = upper(y02.y02_010_m)
		if !y02.y02_005_l and "STAR_DATA@" $ VLC_String and "UNION" $ VLC_String and !("COUNT" $ VLC_String)
			dimension VLA_Array[1]
			VLN_Tot = afields(VLA_Array, PLC_FromFile)
			for VLN_Cont = 1 to VLN_Tot
				if !inlist(VLA_Array[VLN_Cont, 2], "D", "M") and (!("UKEY" $ VLA_Array[VLN_Cont, 1]) or ;
					(("UKEY" $ VLA_Array[VLN_Cont, 1]) and ("UKEYP" $ VLA_Array[VLN_Cont, 1]) or ;
						"CIA_UKEY" $ VLA_Array[VLN_Cont, 1]))
					if inlist(VLA_Array[VLN_Cont, 2], "B", "C", "N")
						VLA_Array[VLN_Cont, 5] = .F.
						if inlist(VLA_Array[VLN_Cont, 2], "B", "N")
							replace all (VLA_Array[VLN_Cont, 1]) with 0 for isnull(evaluate(VLA_Array[VLN_Cont, 1])) in (PLC_FromFile)
						else
							replace all (VLA_Array[VLN_Cont, 1]) with ""  for  isnull(evaluate(VLA_Array[VLN_Cont, 1])) in (PLC_FromFile)
						endif
					endif
				endif
			endfor
		endif
	endif


	if !used(PLC_ToFile)
		*- Tratamento para os cursores em Oracle que as selects possuem union e não são selects de TextAlias
		if this.VON_DBType > 1 and !empty(PLC_Cursor) and this.FOL_LookInCursor(PLC_Cursor)
			VLC_String = upper(y02.y02_010_m)
			if !y02.y02_005_l and "STAR_DATA@" $ VLC_String and "UNION" $ VLC_String and !("COUNT" $ VLC_String)
				create cursor (PLC_ToFile) from array VLA_Array
				
				
				
				select (PLC_ToFile)
				append from dbf(PLC_FromFile)
				VLL_HasBeenCreated = .T.
			endif
		endif

		if !VLL_HasBeenCreated
		*-- abre o cursor a partir do PLC_FromFile
			VLC_Aux = dbf(PLC_FromFile)
			select 0
			use (VLC_Aux) again alias (PLC_ToFile)
			select (PLC_ToFile)
		endif

	*--	Indexa o cursor PLC_ToFile
		index on UKEY tag UKEY
		index on MYCONTROL tag MYCONTROL
		if PLN_IndexN >= 1
			VLC_Macro = "index on " + iif(this.VOA_Index[1] = "-", ;
			            substr(this.VOA_Index[1], 2) + " DESCENDING", this.VOA_Index[1]) + ;
			            " tag CODE"
			&VLC_Macro
		endif	
		for VLN_i = 2 to PLN_IndexN
			VLC_Macro = "index on " + iif(this.VOA_Index[VLN_i] = "-", ;
			            substr(this.VOA_Index[VLN_i], 2) + " DESCENDING", this.VOA_Index[VLN_i]) + ;
			            " tag TAG" + alltrim(str(VLN_i - 1))
			&VLC_Macro
		endfor
		replace all MYCONTROL with "0"
		this.FOL_CursorBuffering(3, VLN_Select)
	else
	*-- Lê o arquivo original para importar seu conteúdo
		select (PLC_ToFile)
		set order to UKEY
		select (PLC_FromFile)

		VLC_PreviousDeleted = set("deleted")

		set deleted off
		scan
			if !deleted()
				if !seek(ukey, PLC_ToFile)
					scatter memvar memo
					select (PLC_ToFile)
					append blank
					gather memvar memo
					if VLL_Replace
						replace MYCONTROL with "0"
					endif
					select (PLC_FromFile)
				endif
			endif
		endscan

		set deleted &VLC_PreviousDeleted

	endif

	select (PLC_ToFile)
	if !empty(tagno("CODE"))
		set order to code
	endif
	go top

	this.FOL_PopSelect()

	return reccount(PLC_ToFile)		&& Retorna o numero de registros criados.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FromFile : Arquivo de Origem  
*/                    PLC_ToFile : Cursor gravável de destino 
*/                    PLL_CloseFromFile : Fecha o arquivo Base no Final 
*/                    PLC_FieldUkey : Campo Ukey 
*/ Descrição        : Atualiza a informacao do cursor editavel 'From' no arquivo Pai 'To' 
*/ Retorno          : Nenhum 
*/----------------------------------------------------------------------------------------
PROCEDURE FOU_SaveCursor
	LParameters PLC_FromFile, PLC_ToFile, PLL_CloseFromFile, PLC_FieldUkey

	LOCAL VLL_Found, VLL_Ret

	if !used(PLC_FromFile)
		return .T.
	endif

	this.FOL_PushSelect()

	*- Muda o Status do DELETED para garantir uma Correcao adequada
	VLC_OldDELETED = set("DELETED")
	set deleted off

	*- abre o arquivo destino se este não estiver aberto
	select (PLC_ToFile)

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
	select (PLC_FromFile)
	this.FOL_PushSelect()

	set order to MYCONTROL

	VLL_Ret = .T.

	=seek("1")
	do while MYCONTROL = "1"
		scatter memvar memo

*---	Se não existir no arquivo Pai Cria
		VLL_Found = seek(evaluate(PLC_FieldUkey), PLC_ToFile, "UKEY")

		select (PLC_ToFile)

*---	Verifica o Status de Deletado
		if deleted(PLC_FromFile)
			if VLL_Found
				gather memvar memo
				delete
			endif
		else
			if VLL_Found
				if deleted()
					recall
				endif
			else
				append blank
			endif
			gather memvar memo
		endif

*---	se não conseguiu atualizar executa um rollback
		if !this.FOL_TableUpdate(PLC_ToFile, .F.)
			VLL_Ret = .F.
			exit
		endif

		select (PLC_FromFile)
		skip
	enddo
	go top

	this.FOL_PopSelect()

	*- Fecha o arquivo origem se necessário
	if PLL_CLoseFromFile		&& Fecha o arquivo base no Final
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina antes de fechar
		this.FOL_CloseTable(PLC_FromFile)
	else
	*--	Executa um pack e volta todos os mycontrol com '0'
		select (PLC_FromFile)
		this.FOL_PackCursor(PLC_FromFile)
		replace all MYCONTROL with "0"
		go top
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina
	endif

	if VLC_OldDELETED = "ON"
		set deleted on
	endif
	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_Num : número do parâmetro 
*/                    PLU_Val : valor do parâmetro  
*/ Descrição        : Seta um parâmetro para execução do select 
*/ Retorno          : nenhum 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SetParameter
	lparameters PLN_Num, PLU_Val, PLN_Type, PLN_IO, PLN_Len, PLC_Name

	*- Tratamento de data mínima no oracle
	if this.VON_DBType <> 1 && Oracle e DB2
		if empty(PLU_Val)
			if inlist(vartype(PLU_Val), "D", "T")
				PLU_Val = {^1900/01/01 00:00:00}
			endif
		endif
	endif

	this.VOA_Where[PLN_Num] = PLU_Val

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias   : Indica o nome do alias  
*/                    PLC_Field01 : Indica o nome dos fields a serem verificados  
*/                    PLC_Field02 : Indica o nome dos fields a serem verificados  
*/                    PLC_Field03 : Indica o nome dos fields a serem verificados  
*/                    PLC_Field04 : Indica o nome dos fields a serem verificados  
*/                    PLC_Field05 : Indica o nome dos fields a serem verificados  
*/ Descrição        : Indica se campos do registro atual do PLC_Alias foi alterado  
*/ Retorno          : .T. : Se foi alterado, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ChangeField
	LParameters PLC_Alias, PLC_Field01, PLC_Field02, PLC_Field03, PLC_Field04, PLC_Field05

	LOCAL  VLC_Ret

	if !empty(PLC_Field01)
		if PLC_Field01 = "-1"
			VLC_Ret = getfldstate(-1, PLC_Alias)
			if len(strtran(VLC_Ret,"1","")) > 0
				return .T.
			endif
		else
			if getfldstate(PLC_Field01, PLC_Alias) > 1
				return .T.
			endif
		endif
	else
		return .F.
	endif
	if !empty(PLC_Field02)
		if getfldstate(PLC_Field02, PLC_Alias) > 1
			return .T.
		endif
	else
		return .F.
	endif
	if !empty(PLC_Field03)
		if getfldstate(PLC_Field03, PLC_Alias) > 1
			return .T.
		endif
	else
		return .F.
	endif
	if !empty(PLC_Field04)
		if getfldstate(PLC_Field04, PLC_Alias) > 1
			return .T.
		endif
	else
		return .F.
	endif
	if !empty(PLC_Field05)
		if getfldstate(PLC_Field05, PLC_Alias) > 1
			return .T.
		endif
	else
		return .F.
	endif

	return .F.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_Id		: 1-TextBox de Pesquisa, 2-Form, 3-Chamada Específica,
*/									  4-Chamada Específica (seleção única)				  
*/ 					  PLC_ObjectId	: Nome, identificador único da pesquisa 
*/ 					  PLO_OriginObj	: Objeto responsável pela chamada da pesquisa 
*/					  PLC_ReturnFile: Nome do Cursor de retorno 
*/ Descrição		: Chama o form de pesquisa, permitindo que o usuário defina o escopo  
*/ Retorno			: .T. : Existe arquivo de retorno - .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_MakeSearch
	LParameters PLN_Id, PLC_ObjectId, PLO_OriginObj, PLC_ReturnFile

	LOCAL VLL_Return, VLC_SelectedAlias

	VLC_SelectedAlias = alias()

	this.FOL_DoNormalForm("Y34", PLN_Id, upper(PLC_ObjectId), PLO_OriginObj, PLC_ReturnFile)
	
	if used("TMP_Search")
		use in TMP_Search
	endif

	VLL_Return = used(PLC_ReturnFile)

	if !empty(VLC_SelectedAlias)
		if used(VLC_SelectedAlias)
			select (VLC_SelectedAlias)
		endif
	endif

	return VLL_Return
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/ Descrição        : Atualiza a informacao do cursor editavel 'From' no Pai 'To' 
*/ Retorno          : .T. - sucesso   .F. - erro, retorna no VON_Error o erro 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SaveRecord
	LParameters PLC_FromFile, PLC_ToFile, PLC_FieldUkey

	LOCAL VLL_Found, VLU_Ukey

	*- se não está aberto o arquivo fonte retorna erro
	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
	select (PLC_FromFile)

	VLU_Ukey = eval(PLC_FieldUkey)


	*- Verifica se existe no arquivo Original
	VLL_Found = seek(VLU_Ukey, PLC_ToFile, "UKEY")

	*- se não existe no arquivo original
	if eof(PLC_ToFile)
	*-- se inseriu no arquivo fonte (status[1] = " "), insere no ADO
		if eval("STATUS") = " "
	*---	retorna se não conseguiu inserir o Registro
			append blank in (PLC_ToFile)
		else
			this.VON_Error = SV_FatalErr
			return .F.
		endif
	else
	*--	verifica se os registros não possuem o mesmo timestamp, retorna com erro
		if eval("timestamp") <> eval(PLC_ToFile + ".timestamp")
			this.VON_Error = SV_FatalErr
			this.VON_SQLError = -1 && Timestamp diferente
			return .F.
		endif
	endif

	*- move do DBF para o ADO
	select (PLC_FromFile)
	replace timestamp with datetime()
	replace status with "W" + subs(status,2)
	scatter memvar memo

	select (PLC_ToFile)
	gather memvar memo

	if !this.FOL_TableUpdate(PLC_ToFile, .F.)
		this.VON_Error = SV_NoFatalErr
		return .F.
	endif

	return .T.
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Select : select a ser executada 
*/ Descrição        : Executa uma select substituindo os VPA_Seek[n] por VPC_Seekn  
*/                    e VPA_Filter[n] por VPC_Filtern 
*/ Retorno          : .T. : Se executou com sucesso, .F. : c.c. 
*/----------------------------------------------------------------------------------------
procedure FOL_SelectLocal
	lparameters PLC_Select
	&PLC_Select
endproc


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Form      : Indica em qual form é executada o cancelamento  
*/                    PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/                    PLC_Which     : Guarda um campo para o cancel update  
*/                    PLL_Close     : Fecha o arquivo Base no Final 
*/                    PLL_OnlyLoop  : Indica se irá executar apenas o loop  
*/					  PLN_DelAction : 0 - Deleta registros deletados e Salva os demais registros
*/									  1 - Ignora registros deletados e Salva os demais registros
*/									  2 - Deleta registros deletados e Ignora os demais registros
*/					  PLC_SaveIndex : Nome do índice para ordenar o salvamento dos registros,
*/									  se estiver vazio utiliza o mycontrol
*/ Descrição        : Atualiza a informacao do cursor editavel 'From' no arquivo Pai 'To' 
*/ Retorno          : .T. - sucesso   .F. - erro, retorna no VON_Error o erro 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SaveCursorSQL
	lparameters PLO_Form, PLC_FromFile, PLC_ToFile, PLC_FieldUkey, PLC_Which, PLL_Close, ;
				PLL_OnlyLoop, PLN_DelAction, PLC_SaveIndex

	LOCAL VLL_Ret, VLC_OldDELETED, VLA_Fields, VLA_Values, VLC_Select, VLN_Ret, VLL_UpdatePB
	PRIVATE VPC_Status, VPT_DateTime, VPC_Null

	if empty(PLN_DelAction)
		PLN_DelAction = 0
	endif

	if empty(PLC_SaveIndex)
		PLC_SaveIndex = "mycontrol"
	endif

	VPC_Null = .NULL.
		
	*- se não está aberto o arquivo fonte retorna erro
	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	this.FOL_PushSelect()

	*- Muda o Status do DELETED para garantir uma Correcao adequada
	VLC_OldDELETED = set("DELETED")
	set deleted off

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
	select (PLC_FromFile)
	this.FOL_PushSelect()

	set order to (PLC_SaveIndex)

	if at(".", PLC_FieldUkey) = 0
		PLC_FieldUkey = PLC_FromFile + "." + PLC_FieldUkey
	endif

	VLL_Ret = .T.

	VLL_UpdatePB = !this.VOL_ServiceMode

	if VLL_UpdatePB and !VGO_ToolBar.VOL_ProgBarRunning
		VLN_RecCount = 0
		count for mycontrol = "1" to VLN_RecCount
		VGO_ToolBar.FOL_ProgressBar("reset")
		VGO_ToolBar.FOL_ProgressBar("setmax", VLN_RecCount, this.FOC_Caption("table_" + lower(PLC_ToFile)))
	endif

	if seek("1")
		*- Executa um scan com MYCONTROL = "1"
		scan while mycontrol = "1"
			*- Verifica se está deletado
			if deleted(PLC_FromFile)
				*- Se estava gravado no deleta o registro no SQL
				if status = "W"
					if PLN_DelAction <> 1 && Ignorar registros deletados
						if !isnull(PLO_Form)
							if !PLO_Form.FOL_CancelUpdates(PLC_Which) && Executa um cancelupdates da tela
								VLL_Ret = .F.
								exit
							endif
							select (PLC_FromFile)
						endif

						if !PLL_OnlyLoop
							if !this.FOL_DeleteRecordSql(PLC_FromFile, PLC_ToFile, PLC_FieldUkey)
								VLL_Ret = .f.
								exit
							endif
						endif
					endif			
				endif
			else
				if PLN_DelAction <> 2 && Ignorar registros não deletados
					if !isnull(PLO_Form)
						if !PLO_Form.FOL_SaveUpdates(PLC_Which, .T.)
							VLL_Ret = .F.
							exit
						endif
						select (PLC_FromFile)
					endif

					if !PLL_OnlyLoop
						if !this.FOL_SaveRecordSql(PLC_FromFile, PLC_ToFile, PLC_FieldUkey)
							VLL_Ret = .f.
							exit
						endif
					endif
				endif
			endif

			if VLL_UpdatePB
				VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
			endif

		endscan

		go top in (PLC_FromFile)

	endif

	this.FOL_PopSelect()

	*- Fecha o arquivo origem se necessário
	if PLL_Close		&& Fecha o arquivo base no Final
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina antes de fechar
		this.FOL_CloseTable(PLC_FromFile)
	else
		select (PLC_FromFile)
		go top in (PLC_FromFile)
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina
	endif

	if VLC_OldDELETED = "ON"
		set deleted on
	endif

	if VLL_UpdatePB
		VGO_ToolBar.FOL_ProgressBar("visible", .f.)
		VGO_ToolBar.FOL_ProgressBar("reset")	
	endif

	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/ Descrição        : Atualiza a informacao do cursor editavel 'From' no ADO Pai 'To' 
*/ Retorno          : .T. - sucesso   .F. - erro, retorna no VON_Error o erro 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SaveRecordSql
	LParameters PLC_FromFile, PLC_ToFile, PLC_FieldUkey

	LOCAL VLL_Ret, VLA_Fields, VLA_Values, VLC_Select, VLN_Ret, VLC_SQLCmd

	PRIVATE VPC_Status, VPT_DateTime, VPC_Null

	VPC_Null = .null.
	VLL_Ret = .T.

	*- se não está aberto o arquivo fonte retorna erro
	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		VLL_Ret = .F.
	else
		if eof(PLC_FromFile)
			return
		endif
	endif

	this.FOL_PushSelect()

	if VLL_Ret
		 this.FOC_Message( "gravando" , " : " + lower(PLC_FromFile))
	*--	Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
		select (PLC_FromFile)

*-		Número de linhas do comando SQL
		VLN_LineCount = alines(VLA_SQLCmd, nvl(sqlcmd, ""))
		if VLN_LineCount > 1
			VLC_SQLCmd = VLA_SQLCmd[2]
		else
			VLC_SQLCmd = ""
		endif

		if empty(VLC_SQLCmd)
			this.FOL_PopSelect()
			return .t.
		endif

		VLL_Insert = empty(subs(status, 1, 1))

*---		se deve inserir
		if VLL_Insert
			VPT_DateTime = datetime()
			VPC_Status   = "W" + subs(status,2)
			VLC_Select = VLC_SQLCmd

*----		Added by Denis. Resolvendo o problema do campo memo no salva. 22/02/2000
			scatter memvar memo
*----		Fim

			with this
				VLN_Ret = 0
				VLL_Ret = .FOL_SqlExec(VLC_Select, "")
*----			o comando foi executado, preciso verificar se foi executado corretamente
				if VLL_Ret
					VLC_TableFullPath = this.FOC_TableIdent(PLC_ToFile) + PLC_ToFile && Identificador da tabela
					VLC_VerifySQLCmd = "select count(*) as reccntr from " + VLC_TableFullPath +;
						" where ukey = ?" + PLC_FieldUkey
					VLL_Verify = .FOL_SqlExec(VLC_VerifySQLCmd, "verify")
					if VLL_Verify
						select verify
						if verify.reccntr > 0 && Indica que foi inserido com sucesso
							VLN_Ret = 1
						else && Não houve inserção do registro
							VLN_Ret = 0
						endif
						use in verify
					endif
				endif
			endwith

*----		se inseriu mais de 1 registro
			if VLN_Ret <= 0
				this.VON_Error = SV_FatalErr
				VLL_Ret = .F.
			else
				VLL_Ret = .T.
			endif


*---		se deve dar um update
		else
			VPT_DateTime = datetime()
			VPC_Status   = "W" + subs(status,2)
			VLC_Select   = VLC_SQLCmd + ' where ukey = ?' + PLC_FieldUkey + ' and timestamp = ?timestamp'
								
*----			Added by Denis. Resolvendo o problema do campo memo no salva. 22/02/2000
			scatter memvar memo
*----			Fim

			with this
				VLN_Ret = 0
				VLL_Ret = .FOL_SqlExec(VLC_Select, "")
*----			o comando foi executado, preciso verificar se foi executado corretamente
				if VLL_Ret
					VLC_TableFullPath = this.FOC_TableIdent(PLC_ToFile) + PLC_ToFile && Identificador da tabela
					VLC_VerifySQLCmd = "select count(*) as reccntr from " + VLC_TableFullPath +;
						' where ukey = ?' + PLC_FieldUkey + ' and timestamp = ?VPT_DateTime'
					VLL_Verify = .FOL_SqlExec(VLC_VerifySQLCmd, "verify")
					if VLL_Verify
						select verify
						if verify.reccntr > 0 && Indica que foi atualizado com sucesso
							VLN_Ret = 1
						else && Não houve atualização do registro
							VLN_Ret = 0
*------						não encontrou ou diferem em datetimes
							this.VON_Error = SV_FatalErr
							this.VON_SQLError = -1 && Timestamp diferente
							VLL_Ret = .F.
						endif
						use in verify
					else
						this.VON_Error = SV_NoFatalErr
						VLL_Ret = .F.
					endif
				else
					this.VON_Error = SV_NoFatalErr
					VLL_Ret = .F.
				endif
			endwith
		endif

		set message to ""
	endif
	
	*- Armazeno o último datetime utilizado
	this.VOD_LastSaveDateTime = VPT_DateTime

	this.FOL_PopSelect()

	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Form      : Indica em qual form é executada o cancelamento  
*/                    PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/                    PLC_Which     : Guarda um campo para o cancel update  
*/                    PLL_Close     : Fecha o arquivo Base no Final 
*/                    PLL_OnlyLoop  : Indica se irá executar apenas o loop  
*/					  PLC_DeleteIndex : Nome do índice para ordenar o arquivo que será apagado
*/ Descrição        : Deleta a informação do cursor editavel 'From' no arquivo Pai 'To' 
*/ Retorno          : .T. : Sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_DeleteCursorSQL
	LParameters PLO_Form, PLC_FromFile, PLC_ToFile, PLC_FieldUkey, PLC_Which, PLL_Close, ;
				PLL_OnlyLoop, PLC_DeleteIndex

	LOCAL VLL_Ret, VLC_OldDELETED, VLC_Select

	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	this.FOL_PushSelect()

	*- Muda o Status do DELETED para garantir uma Correção adequada
	VLC_OldDELETED = set("DELETED")
	set deleted off

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualização
	select (PLC_FromFile)
	this.FOL_PushSelect()

	if at(".", PLC_FieldUkey) = 0
		PLC_FieldUkey = PLC_FieldUkey
	endif

	VLL_Ret = .T.

	if !empty(PLC_DeleteIndex)
		set order to (PLC_DeleteIndex) in (PLC_FromFile)
		go top in (PLC_FromFile)
	endif

	*--	executa um scan no arquivo fonte
	scan
		if !isnull(PLO_Form)
			if !PLO_Form.FOL_CancelUpdates(PLC_Which)
				VLL_Ret = .F.
				exit
			endif
			select (PLC_FromFile)
		endif
		if !PLL_OnlyLoop
			if !this.FOL_DeleteRecordSql(PLC_FromFile, PLC_ToFile, PLC_FieldUkey)
				VLL_Ret = .f.
				exit
			endif
		endif
	endscan

	go top in (PLC_FromFile)

	this.FOL_PopSelect()

	*- Fecha o arquivo origem se necessário
	if PLL_Close		&& Fecha o arquivo base no Final
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina antes de fechar
		this.FOL_CloseTable(PLC_FromFile)
	else
	*--	Executa um pack e volta todos os mycontrol com '0'
		select (PLC_FromFile)
		this.FOL_PackCursor(PLC_FromFile)
		this.FOL_PopSelect()			&& Volto o arquivo anterior a execução dessa rotina
	endif

	if VLC_OldDELETED = "ON"
		set deleted on
	endif
	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FromFile  : Arquivo de Origem 
*/                    PLC_ToFile    : Cursor gravável de destino  
*/                    PLC_FieldUkey : Campo Ukey 
*/ Descrição        : Deleta a informacao do cursor editavel 'From' executando comando  
*/                    SQL. 
*/ Retorno          : .T. - sucesso   .F. - erro, retorna no VON_Error o erro 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_DeleteRecordSql
	LParameters PLC_FromFile, PLC_ToFile, PLC_FieldUkey

	LOCAL VLL_Ret, VLC_Select

	*- se não está aberto o arquivo fonte retorna erro
	if !used(PLC_FromFile)
		this.VON_Error = SV_FatalErr
		return .F.
	endif

	this.FOL_PushSelect()

	*- Varre o arquivo Base procurando o Ukey no arquivo Pai para atualizacao
	select (PLC_FromFile)

	VLL_Ret = .F.

	VLC_TableFullPath = this.FOC_TableIdent(PLC_ToFile) + PLC_ToFile && Identificador da tabela
	VLC_Select = "delete from " + VLC_TableFullPath + " where ukey = ?" + PLC_FieldUkey + ;
					 " and timestamp = ?timestamp"

	VLL_Ret = this.FOL_SqlExec(VLC_Select)

*--	Indica que o delete foi executado, vou verificar se a ukey foi deletada
	VLN_Ret = -1
	if VLL_Ret
		VLC_TableFullPath = this.FOC_TableIdent(PLC_ToFile) + PLC_ToFile && Identificador da tabela
		VLC_VerifySQLCmd = "select count(*) as reccntr from " + VLC_TableFullPath + " where ukey = ?" + PLC_FieldUkey
		VLL_Verify = this.FOL_SqlExec(VLC_VerifySQLCmd, "VERIFY")
		if VLL_Verify
			select verify
			if verify.reccntr = 0 && Indica que a ukey ainda existe
				VLN_Ret = 1
			else
				VLN_Ret = 0 && Timestamp diferentes
			endif
			use in verify
		endif
	endif

	do case
		case VLN_Ret = 1 && Registro deletado corretamente
			VLL_Ret = .T.

		case VLN_Ret = 0 && Registro foi alterado
			VLL_Ret = .F.
			this.VON_Error = SV_FatalErr
			this.VON_SQLError = -1 && Timestamp diferente

		case VLN_Ret = -1 && Registro não foi apagado, necessário verificar o erro do SQL
			VLL_Ret = .F.
			this.VON_Error = SV_NoFatalErr
	endcase

	this.FOL_PopSelect()

	return VLL_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Alias  : alias a ser alterado 
*/                    PLC_Principal : alias do principal  
*/                    PLL_Insert : ****** não é mais usada *******
*/                    PLL_All : se insere todos as colunas/campos 
*/                    PLC_ExcludeFields : indica campos que devem ser excluídos, formato: 
*/                                    <campo 1>, <campo 2>, ... 
*/ Descrição        : Cria/Altera a string para sql 
*/ Retorno          : String modificada  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CreateSqlString
	LParameters PLC_Alias, PLC_Principal, PLL_Insert, PLL_All, PLC_ExcludeFields

	local VLA_Fields, VLA_Values, VLA_FieldsType, VLN_i, VLN_j, VLC_Aux, VLC_FieldName, VLN_Fields
	local VLC_NewRecSt, VLC_OldRecSt, VLC_WriteRecSt, VLN_StructId, VLN_FldCounter, VLC_SqlCommand
	local VLN_FldCount, VLC_FldList, VLN_Counter

	VLC_Principal = alltrim(lower(PLC_Principal))

	VLC_TableFullPath = this.FOC_TableIdent(upper(VLC_Principal)) + VLC_Principal && Identificador da tabela

	if this.VON_DBType = 1 && Sql Server
		VLC_RowLock = " with (rowlock)"
		VLC_NoLock = " (nolock) "
	else && Oracle e DB2
		VLC_RowLock = ""
		VLC_NoLock = ""
	endif

	*- Denis 12/06/2002
	*- Tratamento de campos da tabela no SQL
	VLN_StructId = ascan(this.VOA_ValidTableStruct, VLC_Principal, -1, -1, 1, 7)
	if VLN_StructId <= 0
		this.VON_TableStructCount = this.VON_TableStructCount + 1
		VLN_StructId = this.VON_TableStructCount
		dimension this.VOA_ValidTableStruct[this.VON_TableStructCount, 2]
		this.VOA_ValidTableStruct[this.VON_TableStructCount, 1] = VLC_Principal
		*- Preciso da estrutura do sql
		VLC_SqlCommand = "select * from " + VLC_TableFullPath + VLC_NoLock + " where 1 = 0"
		if this.FOL_SqlExec(VLC_SqlCommand, "sql_struct")
			VLN_FldCount = afields(VLA_SqlStruct, "sql_struct")
			VLC_FldList = " "
			for VLN_Counter = 1 to VLN_FldCount
				VLC_FldList = VLC_FldList + VLA_SqlStruct[VLN_Counter, 1] + " "
			endfor
			this.VOA_ValidTableStruct[this.VON_TableStructCount, 2] = VLC_FldList
			use in sql_struct
		else
			return .f.
		endif
	else
		VLN_StructId = asubscript(this.VOA_ValidTableStruct, VLN_StructId, 1)
	endif

	select (PLC_Alias)

*-	Pego o status atual do registro
	VLC_NewRecSt  = substr(getfldstate(-1, PLC_Alias), 2)
	VLC_OldSQLCmd = nvl(sqlcmd, "")

	VLN_FldCounter = afields(VLA_Changed, PLC_Alias)
	VLN_Fields = 0

*-	Tratamento dos campos excluídos do comando
*-	Denis: Retirei esse tratamento através de parâmetro e o automatizei com a matriz VOA_ValidTableStruct
*- 	Gustavo: Inclui o campo rowguid que é o campo criado quando há replicação de dados no SQL
	if !empty(PLC_ExcludeFields)
		PLC_ExcludeFields = " " + strtran(upper(PLC_ExcludeFields), ",", " ") +  " "
		PLC_ExcludeFields = PLC_ExcludeFields + this.VOC_SQLExcludeFields
	else
		PLC_ExcludeFields = this.VOC_SQLExcludeFields
	endif


*-	Verifico se o registro já foi atualizado alguma vez
	if empty(VLC_OldSQLCmd)
		VLC_OldRecSt = VLC_NewRecSt
		VLC_WriteRecSt = VLC_NewRecSt
	else
		VLN_LineCount = alines(VLA_SQLCmd, VLC_OldSQLCmd)
		if VLN_LineCount >= 1
			VLC_OldRecSt = VLA_SQLCmd[1] && Linha com o antigo getfldstate()
		else
			VLC_OldRecSt = ""
		endif
		VLC_WriteRecSt = VLC_OldRecSt
	endif

	if empty(subs(status, 1, 1))
		VLL_Insert = .t.

		VLC_SQLCmd = "insert into " + VLC_TableFullPath + VLC_RowLock + " ("
	else
		VLL_Insert = .f.

		VLC_SQLCmd = "update " + VLC_TableFullPath + VLC_RowLock + " set "
	endif

	VLC_InsFldCmd = "timestamp, status"
	VLC_InsValCmd = "?VPT_Datetime, ?VPC_Status"
	VLC_UpdCmd = "timestamp = ?VPT_Datetime, status = ?VPC_Status"

	for VLN_i = 1 to VLN_FldCounter
		VLL_Modified = inlist(substr(VLC_NewRecSt, VLN_i, 1), "2", "4") or;
						inlist(substr(VLC_OldRecSt, VLN_i, 1), "2", "4")
		
*--		Verifico se o campo foi modificado ou está na lista de obrigatórios e não está na lista de excluídos
		VLC_FieldStr = " " + VLA_Changed[VLN_i, 1] + " "

		if (PLL_All or VLL_Modified) and (VLC_FieldStr $ this.VOA_ValidTableStruct[VLN_StructId, 2] and;
			 !VLC_FieldStr $ PLC_ExcludeFields)

			VLC_WriteRecSt = stuff(VLC_WriteRecSt, VLN_i, 1, "2")

			VLC_FldName = lower(VLA_Changed[VLN_i, 1]) && Nome do campo
			VLL_EmptyFldValue = empty(nvl(eval(VLC_FldName), "")) && Verifico se está vazio

			if VLA_Changed[VLN_i, 5] and VLL_EmptyFldValue
				VLC_NewValue = "?VPC_Null"
			else
				if this.VON_DBType = 2 and VLA_Changed[VLN_i, 2] = "M"
					VLC_NewValue = "?" + alltrim(PLC_Alias) + "." + VLC_FldName
				else
					VLC_NewValue = "?m." + VLC_FldName
				endif
			endif
			
			if VLL_Insert
				VLC_InsFldCmd = VLC_InsFldCmd + ", " + VLC_FldName
				VLC_InsValCmd = VLC_InsValCmd + ", " + VLC_NewValue
			else
				VLC_UpdCmd = VLC_UpdCmd + ", " + VLC_FldName + " = " + VLC_NewValue
			endif

			VLC_Comma = ", "
		endif
	endfor

	if VLL_Insert
		VLC_SQLCmd = VLC_SQLCmd + VLC_InsFldCmd + ") values (" + VLC_InsValCmd + ")"
	else
		VLC_SQLCmd = VLC_SQLCmd + VLC_UpdCmd
	endif

	VLC_WriteRecSt = VLC_WriteRecSt + CRLF + VLC_SQLCmd

	replace sqlcmd with VLC_WriteRecSt in (PLC_Alias)

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_CfgStr	- String com a configuração da pesquisa
*/                    PLC_ToCursor	- Alias da tabela destino, que sobrepõe a configuração da string
*/ Descrição        : Executa uma pesquisa pré-parametrizada anteriormente
*/ Retorno          : Número de registros da pesquisa
*/----------------------------------------------------------------------------------------
procedure FON_RunStrSearch
	lparameters PLC_CfgStr, PLC_ToCursor

	local VLN_Ret, VLO_SearchObj

	VLN_Ret = 0
	VLO_SearchObj = createobject("cgo_selectscollectionbuilder")
	if VLO_SearchObj.FOL_LoadAnswersStr(PLC_CfgStr)
		if !empty(PLC_ToCursor)
			VLO_SearchObj.VOC_CursorName = PLC_ToCursor
		endif
		VLN_Ret = VLO_SearchObj.FON_MakeSearch()
	endif

	VLO_SearchObj = null

	return VLN_Ret
endproc


*/----------------------------------------------------------------------------------------
*/ Descrição		: Procura uma expressão e retorna se encontrou			  
*/ Parametros       : PLC_CursorCatalog : Cursor a ser executado  
*/                    PLC_Cursor        : Cursor resultado  
*/                    PLC_WhatSeek      : o que deseja procurar 
*/                    PLL_Close         : se deseja fechar depois da execução 
*/ 					  PLL_Edit          : Indica se o cursor criado será editável 
*/					  PLC_Index	        : Indices a serem criados separados por ";" 
*/ Retorno			: .T. - encontrou, .F. c.c						  
*/ Última Alteração	: 30/06/1999  												  
*/ Alterado Por		: Marcelo Ris												  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_FindExpression
	LParameters PLC_CursorCatalog, PLC_Cursor, PLC_WhatSeek, PLL_Close, PLL_Edit, PLC_Index

	LOCAL VLL_Ret, VLN_Occurs, VLN_Index, VLN_i, VLN_Pos, VLC_Temp, VLN_Size, VLN_At
	VLN_Index = 0

	this.FOL_SetParameter(1, PLC_WhatSeek)

	*- Se houverem indices
	if !empty(PLC_Index)
		VLN_Occurs = occurs(";", PLC_Index)
		VLN_Index = VLN_Occurs + 1
		VLN_Pos = 1
		for VLN_i = 1 to VLN_Index
			VLN_At = at(";", PLC_Index, VLN_i)
			VLN_Size = iif(VLN_At > 0, VLN_At - VLN_Pos ,;
			len(PLC_Index) - VLN_Pos + 1)
			 this.VOA_Index[VLN_i] = alltrim(substr(PLC_Index, VLN_Pos, VLN_Size))
			VLN_Pos = VLN_At + 1
		endfor
	endif

	VLL_Ret = .F.
	if !PLL_Edit
		VLL_Ret = this.FOL_ExecuteCursor(PLC_CursorCatalog, PLC_Cursor, 1)
	else
		VLC_Temp = "T" + substr(sys(3), 1, 7)
		VLL_Ret = this.FOL_EditCursor(PLC_CursorCatalog, VLC_Temp, PLC_Cursor, "Z", VLN_Index, 1)
	endif

	if VLL_Ret
		go top in (PLC_Cursor)
		VLL_Ret = !eof(PLC_Cursor)
		if PLL_Close
			this.FOL_CloseTable(PLC_Cursor)
		endif
	endif

	return VLL_Ret
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Select: comando select a ser ajustado 
*/ Descrição        : Ajusta um comando select para ser executado no FOL_ExecuteCursor  
*/ Retorno          : comando select ajustado  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_AdjustSelect
	LParameters PLC_Select

	LOCAL VLC_Macro

	*- se for empty inicia com o 1o. parâmetro setado anteriormente
	if empty(PLC_Select)
		VLC_Macro = this.VOA_Where[1]
		VLN_i = 2
	else
		VLC_Macro = PLC_Select
	endif

	*- Retira o CR, LF, TAB e ; do comando
	VLC_Macro = strtran(strtran(strtran(strtran(VLC_Macro, chr(13), " "), ;
						chr(10), " "), chr(9), " "), ";", " ")

	return VLC_Macro
ENDPROC



*/----------------------------------------------------------------------------------------
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SetDatasession
	lparameters PLN_DataSession
	
	this.VON_DataSessionId = PLN_DataSession
	set datasession to (pln_datasession)
	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ActionInToolBar("", PLN_DataSession, "datasession", _ALLTOOLBARS_ )
	endif
	this.VOO_Security.FOL_SetDataSession(PLN_DataSession)
	this.VOO_JobProc.FOL_SetDataSession(PLN_DataSession)

	VGO_DevKit.FOL_SetDataSession(PLN_DataSession)
	
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FON_RgbScheme		  
*/ Parâmetros  		: 																	  
*/ Descrição   		: Retorna no formato RGB() as cores conforme os Scheme de cores solicitado
*/						
*/						Ex : FON_RgbScheme( 4 )	&& Retorna a cor da Borda da Janela Inativa
*/						Ex : FON_RgbScheme( 6 )	&& Retorna a cor da Borda da Janela Ativa
*/						
*/						
*/ Última alteração : 31/10/2002
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
PROCEDURE FON_RgbScheme
LPARAMETERS PLN_Par, PLL_ForeCOlor, PLN_Scheme

	IF EMPTY(PLN_Scheme)	&& WIndows Default
		PLN_Scheme = 4
	endif

	IF empty(Pln_par)	&& Opcao da cor
		Pln_par= 6
	endif
	IF !between(Pln_par,1,10)	&& Opcoes possiveis !
		Pln_par = 6
	endif

	VLC_Color =  UPPER(rgbSCHEME(PLN_Scheme, Pln_par))
	VLC_Color =  CHRTRAN(VLC_Color, "RGB()", "") + ","

	VLN_MaxCOlor = occurs(  ",", VLC_Color  )

	dimension VLA_Color(VLN_MaxCOlor)

	VLN_ini = 0

	for VLN_i =1 to VLN_MaxColor

		VLN_end = at(",", VLC_Color , VLN_i )
		VLC_Cor = substr(VLC_Color, VLN_ini+1, VLN_end-VLN_ini-1)
		VLN_ini = VLN_end
		
		VLA_Color(VLN_i) = val( VLC_Cor )
		
	next

	if !empty(PLL_ForeCOlor)
		return rgb( VLA_Color(1),VLA_Color(2),VLA_Color(3) )	
	else 
		return rgb( VLA_Color(4),VLA_Color(5),VLA_Color(6) )	
	endif 

ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  : PLN_RGBcolor- Codigo da cor no  
*/             : PLN_FatorCOR- Fator de desvio da cor (40 é o padrao) 
*/             : PLL_ReturnChar- Retorna a cor passada no formato RGB() 
*/ Descrição   : Altera a cor passada como parametro para causar um pequeno efeito visual 
*/ Retorno     : Conforme o parametro passado  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_ChangeColor
	LParameters PLN_RGBcolor, PLN_FatorCOR, PLL_ReturnChar

	LOCAL VLN_Red, VLN_Green, VLN_Blue, VLN_Aux

	if empty(PLN_FatorCOR)
		PLN_FatorCOR = 40
	endif
	*- Inicializa as variaveis
	VLN_Red 	= mod(PLN_RGBcolor, 256)
	VLN_Green	= mod(int(PLN_RGBcolor/256), 256)
	VLN_Blue	= mod(int(PLN_RGBcolor/65536),256)

	if empty(PLL_ReturnChar)

		if ((VLN_Red+VLN_Green+VLN_Blue)/3) > 200
			PLN_FatorCOR = -abs(PLN_FatorCOR)
		else
			PLN_FatorCOR = abs(PLN_FatorCOR)
		endif

		VLN_Red 	= VLN_Red  + PLN_FatorCOR
		VLN_Green	= VLN_Green+ PLN_FatorCOR
		VLN_Blue	= VLN_Blue + PLN_FatorCOR

		VLN_Red 	= min( max(0, VLN_Red  ), 255 )
		VLN_Green	= min( max(0, VLN_Green), 255 )
		VLN_Blue	= min( max(0, VLN_Blue ), 255 )

		return rgb( VLN_Red, VLN_Green, VLN_Blue )
	else
		return alltrim(str(VLN_Red)) + " ," +alltrim(str(VLN_Green)) + " ," + alltrim(str(VLN_Blue ))
	endif
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_FileName : Nome do arquivo a ser utilizado              		  
*/ Descrição		: Retorna o Path do arquivo de imagem 
*/ Retorno			: Noe e Path do arquivo desejado  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_Gallery
	LParameters PLC_FileName

	if !('.' $ PLC_FileName )
		PLC_FileName = trim(PLC_FileName) + ".bmp"	&& Assumo uma extensao padrao
	endif

	return PLC_FileName
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_form   : O form que está ativado  
*/                    PLO_object : O objeto atual 
*/					  PLO_RefreshObj : Objeto que se deseja ativar o refresh() a cada modificação no arquivo
*/ Descrição        : Chama um filter form do form atual  
*/ Retorno          : .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CallFilterForm
	LPARAMETERS PLO_Form, PLO_Object

	this.FOL_DoForm(PLO_Object.VOC_SecondForm, PLO_Object.VON_FormType, PLO_Object)

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_form   : O form que está ativado  
*/                    PLO_object : O objeto atual 
*/ Descrição        : Chama um modal form do form atual 
*/ Retorno          : .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_CallModalForm
	LPARAMETERS PLO_Form, PLO_Object

	LOCAL  VLN_RetMsg, VLL_Ret

	*- Verifica se o valor é diferente do que aparece na tela
	if trim(PLO_Object.value) <> trim(PLO_Object.displayvalue)

	*-- Pergunta se quer inserir registro
		VLN_RetMsg = this.FON_Msg(MSG_NaoExisteInsere)

		do case
	*--- 	Se quiser insere
			case VLN_RetMsg = MBOX_YES
				VLC_Alias = FOC_RealNameAlias(PLO_Object.VOC_NameAlias)
				select (VLC_Alias)
				if this.FOL_NewRecordToForm(VLC_Alias)
					PLO_Object.FOL_AddNew()

					PLO_Object.enabled = .F.
					this.FOL_DoForm(PLO_Object.VOC_InsertForm, FRM_MODAL, PLO_Object)
					PLO_Object.enabled = .T.
					select (PLO_Form.VOC_AliasDefault)
					if VGU_FormReturn
						PLO_Object.FOL_RetNew()
						PLO_Object.rowsource = ""
						PLO_Object.FOL_SelectCommand()
						PLO_Object.value = PLO_Object.displayvalue
						PLO_Object.refresh()
						return .T.
					else
						PLO_Object.VOC_OldWhatToSeek = ""
						PLO_Object.rowsource = ""
						PLO_Object.refresh()
						PLO_Object.setfocus()
						return .T.
					endif
				endif

	*--- 	Se não quiser volta e mantém em edição
			case VLN_RetMsg = MBOX_NO
				return .F.

	*---	Se cancela volta o valor anteriror e continua
			case VLN_RetMsg = MBOX_CANCEL
				PLO_Object.VOC_OldWhatToSeek = ""
				PLO_Object.rowsource = ""
				=PLO_Object.refresh()
				return .T.
		endcase

	endif
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : Nenhum 
*/ Descrição        : Verifica se o form é um objeto  
*/ Retorno          : .T. : Se é um objeto, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_FormisObject

	*- Retorna se o form é um objeto e o BaseClass é um Form
	return (type("_screen.activeform") == "O" AND ;
	        upper(_screen.ActiveForm.BaseClass) = "FORM")
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_FormId 			: Identificador do Form 
*/               	  PLN_Action 			: Ação a ser executada (READING, RECORDING) 
*/               	  PLN_Top, PLN_Left	    : Posição de topo e Lado esquerdo 
*/               	  PLL_Docked, PLN_DockPosition 	: Se está Docked e sua Posição  
*/               	  PLN_Height, PLN_Width 		: Altura e Largura do Form  
*/ Descrição   		: Seta ou guarda a posição de um form no arquivo Y09  
*/ Retorno     		: .T. : Executou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SetUserFormPosition
	LPARAMETERS PLC_Id, PLN_Action, PLN_N1, PLN_N2, PLN_L1, PLN_N3, PLN_N4, PLN_N5, PLN_L2, PLC_M1

	LOCAL VLN_Found, VLC_UserCode, VLL_Close

	*- Se não existir o diretorio correto do usuario, sai!
	if empty(this.VOC_DirSettings)
		return .t.
	endif

	*- Verifica se os parâmetros de Dock são empty

	if empty(PLN_N1)
		PLN_N1 = 0
	endif

	if empty(PLN_N2)
		PLN_N2 = 0
	endif

	if empty(PLN_L1)
		PLN_L1 = .F.
	endif

	if empty(PLN_L2)
		PLN_L2 = .F.
	endif

	if empty(PLN_N3)
		PLN_N3 = 0
	endif

	if empty(PLN_N4)
		PLN_N4 = 0
	endif

	if empty(PLN_N5)
		PLN_N5 = 0
	endif

	if empty(PLC_M1)
		PLC_M1 = ""
	endif

	*- Guarda o arquivo corrente
	this.FOL_PushSelect()

	*- Função colocada para verificar se o resource está intacto.
	this.FOL_SetUserResource()

	if !used("USERRESOURCE")
		use (this.VOC_FileSettings) in 0 alias USERRESOURCE
		VLL_Close = .T.
	endif 
	
	select userresource
		
	if !seek(padr(PLC_Id, len(userresource.id)), "USERRESOURCE", "ID")
		append blank
		replace id with PLC_Id in USERRESOURCE
		PLN_Action = 1 && RECORDING
	endif

	if PLN_Action = 1 && RECORDING
		if rlock()
			replace	userresource.N1	with IIF(EMPTY(PLN_N1), 0  , PLN_N1) ,;
					userresource.N2	with IIF(EMPTY(PLN_N2), 0  , PLN_N2) ,;
					userresource.N3 with IIF(EMPTY(PLN_N3), 0  , PLN_N3) ,;
					userresource.N4 with IIF(EMPTY(PLN_N4), 0  , PLN_N4) ,;
					userresource.N5 with IIF(EMPTY(PLN_N5), 0  , PLN_N5) ,;
					userresource.L1	with IIF(EMPTY(PLN_L1), .f., PLN_L1) ,;
					userresource.L2 with IIF(EMPTY(PLN_L2), .f., PLN_L2) ,;
					userresource.M1 with IIF(EMPTY(PLC_M1), "", PLC_M1) in USERRESOURCE
			
			VLN_Found = .T.
			unlock
		else
			VLN_Found = .F.
		endif
	else
		*- Restaura
		PLN_N1 = UserResource.N1
		PLN_N2 = UserResource.N2
		PLN_N3 = UserResource.N3
		PLN_N4 = UserResource.N4
		PLN_N5 = UserResource.N5
		PLN_L1 = UserResource.L1
		PLN_L2 = UserResource.L2
		PLC_M1 = UserResource.M1
		VLN_Found = .T.
	endif

	*- Restaura o arquivo corrente
	this.FOL_PopSelect()

	this.FOL_CloseTable("USERRESOURCE")

	return VLN_Found
endproc

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_FormName : Nome do form a ser referenciado  
*/ Descrição        : Procura o form no _screen, e retorna uma referência a ele 
*/ Retorno          : Ponteiro ao form, .NULL. : Caso não encontrou 
*/----------------------------------------------------------------------------------------
PROCEDURE FOO_ObjectForm
	LPARAMETERS PLC_FormName

	LOCAL VLN_i

	PLC_FormName = upper(PLC_FormName)

	for VLN_i = 1 to _screen.FormCount
		if upper(alltrim(_screen.forms[VLN_i].name)) == PLC_FormName
			return _screen.forms[VLN_i]
		endif
	endfor

	return .NULL.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Object : Ponteiro ao objeto 
*/                    PLO_Source : Ponteiro ao objeto que tem o foco  
*/ Descrição        : Verifica se o objeto possui o foco ou um de seus componentes  
*/ Retorno          : .T. : Possui o foco, .F. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_HasFocus
	LParameters PLO_Object, PLO_Source

	*- Se o source for null retorna string vazia
	if isnull(PLO_Source)
		return .F.
	endif

	do while .T.
		if PLO_Source.name == PLO_Object.name
			if this.FOC_ObjectName(PLO_Source) == this.FOC_ObjectName(PLO_Object)
				return .T.
			endif
		endif
		if upper(alltrim(PLO_Source.baseclass)) == "FORM"
			return .F.
		endif
		PLO_Source = PLO_Source.parent
	enddo
ENDPROC


*-- Fecha todas as janelas, restaura o nome da janela principal, etc.
PROCEDURE FOL_Cleanup
	LParameters PLL_ExitWindows

	*-- Quando desejamos terminar a aplicação, não podemos simplesmente dar um
	*-- release no objeto oApp e esperar que o método Destroy seja executado
	*-- antes de executar CLEAR EVENTS desde que READ EVENTS foi executado no método Do().
	*-- Sendo assim, este método foi criado para limpar o ambiente antes de terminar a
	*-- aplicação. Ele também permite a interrupção condicional de saída por qualquer razão.

	LOCAL PLN_Form, PLN_FormToClose, VLN_i

	*- fecha os forms
	PLN_FormToClose = 1
	for PLN_Form = 1 to _screen.FormCount
		if type("_screen.Forms(PLN_FormToClose)") == "O" .and. ;
	  		upper(alltrim(_screen.Forms(PLN_FormToClose).baseclass)) == [FORM]
			if _screen.Forms(PLN_FormToClose).QueryUnload()
				_screen.Forms(PLN_FormToClose).Release()
			else
				return .F.
	  		endif
		else
			PLN_FormToClose = PLN_FormToClose + 1
		endif
	endfor

	_screen.caption = this.VOC_OldWindCaption

	if !this.VOL_ServiceMode and vartype(VGO_ToolBar) = "O"
		VGO_ToolBar.FOL_ActionInToolBar("", _TRUE_, "destroy", _ALLTOOLBARS_ )
		
		if this.VON_WhatVersion > 0
			set sysmenu to default
			VGO_ToolBar.FOL_FoxToolbars(_TRUE_)
		endif
		VGO_ToolBar = .null.
	endif

	this.VOO_Agent = .NULL.			


	set classlib to
	set procedure to
	
	*- Elimina arquivos temporários
	this.FOL_CleanTmpDir()

	this.VOL_IsClean = .T.

	close databases all
	clear events

	*- fecha o windows
	if PLL_ExitWindows
		*declare ExitWindowsEx in user32 integer, integer
		*ExitWindowsEx(1,0)
	endif

	this.VON_ScreenObject = 0
	this.VOO_ScreenObject = _screen
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_Form   : Nome do Form 			  		               		  
*/           		: PLN_Type   : Tipo de Tela que será chamada             		  	  
*/           		: PLU_1      : Se é form											  
*/ Descrição		: Pega o nome da form como parâmetro, roda a form,					  
*/					  e põe o toolbar se necessário.							  
*/ Retorno			: .T. executou a rotina com sucesso, .F. .c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_DoForm
	LPARAMETERS PLC_Form, PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, ;
						PLU_10, PLU_11, PLU_12, PLU_13, PLU_14

	LOCAL VLU_Ret, VLC_FormToDo, VLN_Parameters, VLC_AuxForm, VLL_Execute

	*- guarda o número de parâmetros passados
	VLN_Parameters = PARAMETERS()

	PLC_Form = alltrim(upper(PLC_Form))

	*- Pega a padrão
	VLC_FormToDo = ALLTRIM(PLC_Form)

	dimension this.VOA_Parameters[15]
	this.VOA_Parameters = .F.
	this.VON_Parameters = VLN_Parameters
	this.VOA_Parameters[1] = PLN_Type
	this.VOA_Parameters[2] = PLU_1
	this.VOA_Parameters[3] = PLU_2
	this.VOA_Parameters[4] = PLU_3
	this.VOA_Parameters[5] = PLU_4
	this.VOA_Parameters[6] = PLU_5
	this.VOA_Parameters[7] = PLU_6
	this.VOA_Parameters[8] = PLU_7
	this.VOA_Parameters[9] = PLU_8
	this.VOA_Parameters[10] = PLU_9
	this.VOA_Parameters[11] = PLU_10
	this.VOA_Parameters[12] = PLU_11
	this.VOA_Parameters[13] = PLU_12
	this.VOA_Parameters[14] = PLU_13
	this.VOA_Parameters[15] = PLU_14

	if VLN_Parameters < 2
		this.VON_ThisCallType = FRM_NORMAL
		DO FORM (VLC_FormToDo)
		VLU_Ret = this.VOU_FormReturn
	else
		do case
			case VLN_Parameters = 2
				do form (VLC_FormToDo) with PLN_Type

			case VLN_Parameters = 3
				do form (VLC_FormToDo) with PLN_Type, PLU_1

			case VLN_Parameters = 4
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2

			case VLN_Parameters = 5
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3

			case VLN_Parameters = 6
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4

			case VLN_Parameters = 7
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5

			case VLN_Parameters = 8
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6

			case VLN_Parameters = 9
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7

			case VLN_Parameters = 10
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8

			case VLN_Parameters = 11
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9

			case VLN_Parameters = 12
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9,;
											PLU_10

			case VLN_Parameters = 13
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9,;
											PLU_10, PLU_11

			case VLN_Parameters = 14
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9,;
											PLU_10,	PLU_11, PLU_12

			case VLN_Parameters = 15
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9,;
											PLU_10,	PLU_11, PLU_12, PLU_13
			case VLN_Parameters = 16
				do form (VLC_FormToDo) with PLN_Type, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9,;
											PLU_10,	PLU_11, PLU_12, PLU_13, PLU_14
		endcase
		VLU_Ret = this.VOU_FormReturn
	endif

	return VLU_Ret
ENDPROC


*-- Chama um normal form.
PROCEDURE FOL_DoNormalForm
	*-- executa telas normais com ou sem parâmetros
	LPARAMETERS PLC_Form, PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, ;
						PLU_10, PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
						PLU_18, PLU_19, PLU_20, PLU_21, PLU_22, PLU_23, PLU_24, PLU_25

	LOCAL VLU_Ret, VLC_FormToDo, VLN_Parameters, VLL_Execute

	*- guarda o número de parâmetros passados
	VLN_Parameters = PARAMETERS()

	VLC_FormToDo = ALLTRIM(PLC_Form)

	this.VON_ThisCallType = FRM_NORMAL

	VLL_Execute = .T.
	*!*	if type("_screen.activeform.cmd_focus") = "O"
	*!*		_screen.activeform.cmd_focus.enabled = .T.
	*!*		_screen.activeform.cmd_focus.setfocus()
	*!*		_screen.activeform.cmd_focus.enabled = .F.
	*!*		VLL_Execute = _screen.activeform.VOL_OutEdit
	*!*	endif

	if VLL_Execute
	*--	inicia o array de parâmetros
		dimension this.VOA_Parameters[25]

		with this
			.VOA_Parameters = .F.
			.VON_Parameters = VLN_Parameters
			.VOA_Parameters[1] = PLU_1
			.VOA_Parameters[2] = PLU_2
			.VOA_Parameters[3] = PLU_3
			.VOA_Parameters[4] = PLU_4
			.VOA_Parameters[5] = PLU_5
			.VOA_Parameters[6] = PLU_6
			.VOA_Parameters[7] = PLU_7
			.VOA_Parameters[8] = PLU_8
			.VOA_Parameters[9] = PLU_9
			.VOA_Parameters[10] = PLU_10
			.VOA_Parameters[11] = PLU_11
			.VOA_Parameters[12] = PLU_12
			.VOA_Parameters[13] = PLU_13
			.VOA_Parameters[14] = PLU_14
			.VOA_Parameters[15] = PLU_15
			.VOA_Parameters[16] = PLU_16
			.VOA_Parameters[17] = PLU_17
			.VOA_Parameters[18] = PLU_18
			.VOA_Parameters[19] = PLU_19
			.VOA_Parameters[20] = PLU_20
			.VOA_Parameters[21] = PLU_21
			.VOA_Parameters[22] = PLU_22
			.VOA_Parameters[23] = PLU_23
			.VOA_Parameters[24] = PLU_24
			.VOA_Parameters[25] = PLU_25
		endwith

		do case

			case VLN_Parameters = 1
				do form (VLC_FormToDo)

			case VLN_Parameters = 2
				do form (VLC_FormToDo) with PLU_1

			case VLN_Parameters = 3
				do form (VLC_FormToDo) with PLU_1, PLU_2

			case VLN_Parameters = 4
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3

			case VLN_Parameters = 5
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4

			case VLN_Parameters = 6
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5

			case VLN_Parameters = 7
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6

			case VLN_Parameters = 8
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7

			case VLN_Parameters = 9
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8

			case VLN_Parameters = 10
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9

			case VLN_Parameters = 11
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10

			case VLN_Parameters = 12
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11

			case VLN_Parameters = 13
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12

			case VLN_Parameters = 14
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13

			case VLN_Parameters = 15
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14

			case VLN_Parameters = 16
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15
			case VLN_Parameters = 17
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16
			case VLN_Parameters = 18
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17
			case VLN_Parameters = 19
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18
			case VLN_Parameters = 20
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17,;
											PLU_18, PLU_19
			case VLN_Parameters = 21
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20
			case VLN_Parameters = 22
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20, PLU_21
			case VLN_Parameters = 23
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20, PLU_21, PLU_22
			case VLN_Parameters = 24
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20, PLU_21, PLU_22, PLU_23
			case VLN_Parameters = 25
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20, PLU_21, PLU_22, PLU_23, PLU_24
			case VLN_Parameters = 26
				do form (VLC_FormToDo) with PLU_1, PLU_2, PLU_3, PLU_4, PLU_5, PLU_6, PLU_7, PLU_8, PLU_9, PLU_10, ;
											PLU_11, PLU_12, PLU_13, PLU_14, PLU_15, PLU_16, PLU_17, ;
											PLU_18, PLU_19, PLU_20, PLU_21, PLU_22, PLU_23, ;
											PLU_24, PLU_25
		endcase
		return this.VOU_FormReturn
	endif

	return .T.
ENDPROC



*-- Executa aplicações não FoxPro.
PROCEDURE FOL_RunExe
	LPARAMETERS PLC_ProgramName, PLL_NoWait, PLC_WindowType

	LOCAL VLC_Macro

	*- Quanto aos valores de PLN_WindowType
	*-	1	Janela Ativa e tamanho normal
	*-	2	Janela Ativa e minimizada
	*-	3	Janela Ativa e maximizada
	*-	4	Janela Inativa e tamanho normal
	*-	7	Janela Inativa e minimizada

	VLC_Macro = "run "

	if PLL_Nowait
		VLC_Macro = VLC_Macro + " /N "
		if !empty(PLC_WindowType) and subs(PLC_WindowType, 1, 1) $ "12347"
			VLC_Macro = VLC_Macro + subs(PLC_WindowType, 1, 1) + " "
		endif
	endif

	VLC_Macro = VLC_Macro + PLC_ProgramName

	&VLC_Macro

	return
ENDPROC



*-- Init do Objeto cgo_application incorporado para o objeto cgo_prgform.
PROCEDURE FOL_InitApplication
	LParameters PLC_Program, PLC_Menu

	PUBLIC VGO_Ctb, VGO_Acc, VGO_Scripts, VGO_ToolBar, VGO_Lib, VGO_Custom, VGO_WebPages, VGO_Mail, VGO_Cost, VGO_devKit

	local VLL_SecuritySP, VLC_UDFPARMS 

	*-	Declaracao das DLLs
	
	*	Somente para Windows 2000  (Transparencia)	
	IF VAL(OS(3)) >= 5	
		DECLARE SetWindowLong In Win32Api AS _Sol_SetWindowLong Integer, Integer, Integer
		DECLARE SetLayeredWindowAttributes In Win32Api AS _Sol_SetLayeredWindowAttributes Integer, String, Integer, Integer
	ENDIF

	*- Abre uma chave no registro
	declare integer RegOpenKey in win32api integer nHKey, string cSubKey, integer @nHandle
	*- Cria uma chave no registro
	declare integer RegCreateKey in win32api integer nHKey, string cSubKey, integer @nHandle
	*- Fecha uma chave do registro
	declare integer RegCloseKey in win32api integer nHKey
	*- Apaga o valor de uma chave no registro
	declare integer RegDeleteValue in win32api integer nHKEY, string cEntry
	*- Altera o valor de uma chave no registro
	declare integer RegSetValueEx in win32api integer nHKey, string lpszEntry, integer dwReserved, integer fdwType, string lpbData, integer cbData
	*- Instancia um aplicativo para visualizar um arquivo, identificando-o pela extensão
	declare integer ShellExecute in shell32.dll integer nWinHandle, string cOperation, string cFileName,;
		string cParameters,	string cDirectory, integer nShowWindow
	*- Lê o valor de uma chave no registro
	declare integer RegQueryValueEx in  win32api ;
		integer nHKey, string lpszValueName, integer dwReserved, integer @lpdwType, string @lpbData, integer @lpcbData
	*- Permite ler as cores do sistema
	declare integer GetSysColor in user32 integer nIndex

	*- Alteração de atributos de arquivo
	declare short SetFileAttributes in kernel32 string lpFileName, integer dwFileAttributes
	*- Retorno de atributos de arquivo
	declare integer GetFileAttributes in kernel32 string lpFileName 

	*-	FIM da Declaracao das DLLs

	************************************************************************************************************************
	* Verificação do formato de datas
	************************************************************************************************************************
	
	set century on

	*- Verifico se o ano está configurado para 4 posições (século ligado)	
	if !"2000" $ dtoc(date(2000, 1, 1))
		this.FON_Msg(1614)
		quit
	endif
	
	************************************************************************************************************************

	this.VOC_DirHome 	= addbs(this.VOC_DirHome)
	this.VOO_Security 	= createobject("CGO_LibMain06")

	VGO_Ctb 			= createobject("CGO_Formulas")
	VGO_Acc 			= createobject("CGO_PrgAccounting")
	VGO_Cost    		= createobject("CGO_PrgCosting")

	if !this.VOL_ServiceMode
		VGO_ToolBar 	= createobject("CGO_LibMain05")
	endif

	VGO_Lib     		= createobject("CGO_LibMain07")
	VGO_Custom			= createobject("CGO_LibMain08")
	VGO_WebPages		= createobject("CGO_LibMain11")
	VGO_Mail			= createobject("CGO_LibMain03")
	this.VOO_WinSock	= createobject("CGO_LibMain09")
	this.VOO_JobProc	= createobject("CGO_LibMain14")

	
	this.VOC_InitialSetPath = set("path")
	
	if !this.VOL_ServiceMode
		this.FOL_SetUserResource()
		VGO_ToolBar.FOL_FoxToolbars(_FALSE_)
	endif

	VGO_DevKit			= createobject("CGO_TecInfo")			&& CGO_LibMain13

*-	Seto o idioma principal para Português
	this.VOC_USER_Language = "055"
	this.FOL_InitArray()

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_BuildToolBar()
		VGO_ToolBar.FOL_ProgressBar("reset")
		VGO_ToolBar.FOL_ProgressBar("setmax", 3, "")

		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
	endif

	this.FOL_DeclarePublicVars()

	if !empty(PLC_Program)
		this.VOC_MainWindCaption = PLC_Program
	else
		this.VOC_MainWindCaption = ""
	endif
	if !empty(PLC_Menu)
		this.VOC_MainMenu = PLC_Menu
	endif
	LOCAL VLA_Version[12]

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
	endif

	*-- Inicializa o ambiente
	this.FOL_InitEnvironment()

	*-- Salva o título corrente e coloca um novo
	this.VOC_OldWindCaption = _screen.Caption

	*- armazena a versão do foxpro
	this.VON_WhatVersion = version(2)
	this.VOC_Version = this.FOC_GetVersion()

	_screen.icon = this.FOC_Gallery("starsoft.ico")
	_screen.caption = this.VOC_MainWindCaption+" - "+this.VOC_Version + ' ' + this.VOC_SPVersion

	*!*	RELEASE LIBRARY FOXTOOLS.FLL
	this.VOC_ScreenName = _screen.caption

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
		VGO_ToolBar.FOL_ProgressBar("visible", .f.)
	endif

	*- Trtamento para tornar os formatos de data do BD iguais ao do VFP
	do case
		case this.VON_DBType = 1 && MS-SQL Server
			this.FOL_SqlExec("SET DATEFORMAT " + this.VOC_DateFormat)

		case this.VON_DBType = 2 && Oracle
			*- O Formato do Oracle de datas é 'YYYY-Mon-DD HH24:MI:SS'
			VLC_OracleFormat = this.VOC_DateFormat
			VLC_OracleFormat = stuff(VLC_OracleFormat, 2, 0, "/") && Coloco o primeiro / (ex. D/MY)
			VLC_OracleFormat = stuff(VLC_OracleFormat, 4, 0, "/") && Coloco o segundo / (ex. D/M/Y)
			VLC_OracleFormat = strtran(VLC_OracleFormat, "Y", "YYYY")
			VLC_OracleFormat = strtran(VLC_OracleFormat, "M", "MM")
			VLC_OracleFormat = strtran(VLC_OracleFormat, "D", "DD")
			VLC_OracleFormat = "'" + VLC_OracleFormat + " HH24:MI:SS'"

			this.FOL_SqlExec("alter session set NLS_DATE_FORMAT=" + VLC_OracleFormat)
			this.FOL_SqlExec("alter session set nls_sort = 'BINARY'")
	endcase

	*- Armazena as cores do windows
	for VLN_SystemColors = 1 to MAX_SYSTEMCOLOR + 1
		this.VOA_SystemColors[VLN_SystemColors] = GetSysColor(VLN_SystemColors - 1)
	endfor

	clear dlls GetSysColor
ENDPROC

*/----------------------------------------------------------------------------------------


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_Message : Menssagem para exibiçao                       		  
*/           		: PLC_Complemento : complemento da menssagem                  		  
*/ Descrição		: Monstra uma messagem no rodape da tela  
*/ Retorno			: A Message traduzida  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_Message
	LParameters PLC_Message, PLC_complemento

	LOCAL VLC_Return

	if empty(PLC_complemento)
		PLC_complemento = ""
	endif

	VLC_Return = this.FOC_Caption(PLC_Message) + this.FOC_Caption(PLC_complemento)

	Set MESSAGE to VLC_Return


	return VLC_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Form      - Nome do Form (Tela) 
*/                    PLC_Propertie - Propriedade a ser consultada  
*/ Descrição   		: Pega o valor da propriedade 
*/ Retorno     		: Valor da propriedade 
*/----------------------------------------------------------------------------------------
PROCEDURE FOU_GetPropertieForm
	LParameters PLC_Form, PLC_Propertie

	LOCAL VLU_Ret, VLO_Form, VLC_Macro, VLU_Ret

	*- Guarda a Rotina de Erros Padrão
	VLC_Error = on("error")

	*- Desvia Rotina de Erros
	VLU_Ret = .T.
	on error VLU_Ret = .F.

	*- Inicializa o Form
	VLO_Form  = this.FOO_ObjectForm(PLC_Form)

	*- Preenche a variável retorno com o valor da propriedade
	VLU_Ret   = VLO_Form.&PLC_Propertie

	*- Retorna a Rotina de Erros Padrão
	on error &VLC_Error

	return VLU_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Form      - Nome do Form (Tela) 
*/                    PLC_Propertie - Propriedade a ser consultada  
*/ Descrição   		: Altera o valor da propriedade 
*/ Retorno     		: .T., se com alterou com sucesso, .F.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOU_SetPropertieForm
	LParameters PLC_Form, PLC_Propertie, PLU_Value

	LOCAL VLL_Ret, VLO_Form, VLC_Macro, VLU_Ret, VLC_Error

	*- Guarda a Rotina de Erros Padrão
	VLC_Error = on("error")

	*- Desvia Rotina de Erros
	VLL_Ret = .T.
	on error VLL_Ret = .F.

	*- Inicializa o Form
	VLO_Form  = this.FOO_ObjectForm(PLC_Form)

	*- Altera o valor da propriedade com PLC_Propertie
	VLO_Form.&PLC_Propertie = PLU_Value

	*- Retorna a Rotina de Erros Padrão
	on error &VLC_Error

	return VLL_Ret
ENDPROC


*-- Inicializa o ambiente.
PROCEDURE FOL_InitEnvironment
	LOCAL VLC_Help

	*-- Seta os SETs e ONs 
	SET SYSFORMATS 	ON
	*- Denis 27/06/2002
	*- Esse set coloca o set date para short ou long, dependendo do windows. Isso causa um problema em textbox
	*-  que armazenam data, pois data anteriores a 01/01/1601 e datas inválidas apresentam o erro 2034 do fox
	*-  e não existe método ou evento que esse erro possa ser tratado, mostrando o erro para o usuário. Por isso
	*-  preciso verificar o formato do windows para setar o correto no VFP.
	
	VLC_TestDate = dtoc(date(2000, 12, 31))
	dimension VLA_DateOrder[3,2]
	VLA_DateOrder[1,1] = at("31", VLC_TestDate)
	VLA_DateOrder[1,2] = "D"
	VLA_DateOrder[2,1] = at("12", VLC_TestDate)
	VLA_DateOrder[2,2] = "M"
	VLA_DateOrder[3,1] = at("2000", VLC_TestDate)
	VLA_DateOrder[3,2] = "Y"
	asort(VLA_DateOrder, -1, -1, 0, 1)
	VLC_SetDate = VLA_DateOrder[1,2] + VLA_DateOrder[2,2] + VLA_DateOrder[3,2]
	set date 		to (VLC_SetDate)
	this.VOC_DateFormat = VLC_SetDate
	
	*-
	SET REPRO		TO 	60
	SET REFRE		TO	5
	SET TALK		OFF
	SET	CARR		OFF
	SET EXAC		OFF
	SET CENTURY     ON
	SET HOURS       TO 24   
	SET FIXED		OFF
	SET SAFETY 		OFF
	SET MEMOWIDTH 	TO 	256
	SET MULTILOCKS 	ON  
	set message to
	SET DELETED 	ON
	SET EXCLUSIVE 	OFF
	SET NOTIFY 		OFF
	SET BELL 		OFF
	SET NEAR 		OFF
	SET EXACT 		OFF
	SET INTENSITY 	OFF
	SET CONFIRM 	ON
  	SET ESCAPE OFF
	SET NULLDISPLAY TO ""
	SET COLLATE TO "MACHINE"
	SET DECIMALS TO 6
	SET STATUS BAR OFF
	SET ENGINEBEHAVIOR 70

	if !this.VOL_ServiceMode
		* verificar função para o shutdown
		ON SHUTDOWN VGO_Gen.FOL_ShutDown()
	endif
ENDPROC



*/-----------------------------------------------------------------------------------------
*/ Parametros  		: 
*/ Descrição   		: Cria a area de trabalho do usuario (RESOURCE)  
*/ Retorno     		: Não há retorno  
*/-----------------------------------------------------------------------------------------
PROCEDURE FOL_SetUserResource

LOCAL VLC_MachineId, VLC_AliasRes, VLL_RSServer
VLC_AliasRes = ""

if used("resource")
	use in resource
endif 

VLL_RSServer = (this.VON_DbType = 2 and !(this.FOC_StartPar("RESOURCE") == 'DBSERVER'))  

with this
	.VOC_MachineId   = sys(0)
	.VOC_NetMachine  = ""
	.VOC_NetUserName = "" 
	.VOC_DirSettings = "" 
	.VOC_UserTmp	 = ""

	VLC_MachineId   = .VOC_MachineId

	if vartype(VLC_MachineId) = "C"

		IF "#" $ .VOC_MachineId
			.VOC_NetMachine  = alltrim(lower(substr( .VOC_MachineId, 1, at("#", .VOC_MachineId) -1 )))
			.VOC_NetUserName = alltrim(lower(substr( .VOC_MachineId,    at("#", .VOC_MachineId) +1 )))

			if empty(.VOC_NetPathResource) and !VLL_RSServer
				.VOC_NetPathResource = sys(3)
			endif

			IF !EMPTY(.VOC_NetUserName)
				.VOC_DirSettings= this.VOC_DirHome + "My-Apps\Controls\Settings\" + .VOC_NetUserName + IIF(!VLL_RSServer, "\", '') + .VOC_NetPathResource
				IF !Directory(.VOC_DirSettings)
					MKDIR (.VOC_DirSettings)
				endif
			else
				.VOC_DirSettings= this.VOC_DirHome + "My-Apps\Controls\Settings\" + "default" + IIF(!VLL_RSServer, "\", '') + .VOC_NetPathResource
				IF !Directory(.VOC_DirSettings)
					MKDIR (.VOC_DirSettings)
				endif
			endif
		endif
		
		if !directory(.VOC_DirSettings)
			.VOC_DirSettings  = ""
			.VOC_FileSettings = ""
			.voc_CutCopyFile  = ""
		else
		
			if right(.VOC_DirSettings, 1) <> "\"
		    	.VOC_FileSettings = .VOC_DirSettings + "\resource.res"
		    	.voc_CutCopyFile  = .VOC_DirSettings + "\resource1.res"
		    	.voc_UserTmp      = .VOC_DirSettings + "\temp"
		   	else
		    	.VOC_FileSettings = .VOC_DirSettings + "resource.res"
		    	.voc_CutCopyFile  = .VOC_DirSettings + "resource1.res"
		    	.voc_UserTmp      = .VOC_DirSettings + "temp"
		   	endif

			*- Diretorio De arquivos temporarios do usuario
			if !directory(.voc_UserTmp)
				MKDIR (.voc_UserTmp)
			endif
			VLC_AliasRes = sys(2015)
			if !file(.VOC_FileSettings) or !this.FOL_usetable(this.VOC_FileSettings, VLC_AliasRes, "",.T.,.F.,.T.,.F.) or empty(tag(1, VLC_AliasRes))
				this.FOL_CreateResourceFile(.t., .f.)
			endif
			this.FOL_CloseTable(VLC_AliasRes)

			VLC_AliasRes = sys(2015)
			if !file(.VOC_CutCopyFile  ) or !this.FOL_usetable(this.VOC_CutCopyFile, VLC_AliasRes, "",.T.,.F.,.T.,.F.) or empty(tag(1, VLC_AliasRes)) or empty(tag(2, VLC_AliasRes))
				this.FOL_CreateResourceFile(.f., .t.)
			endif
			this.FOL_CloseTable(VLC_AliasRes)


		endif
		
		if !file(.VOC_FileSettings)
	  		.VOC_FileSettings = ""
		endif

		if !file(.VOC_CutCopyFile)
	  		.VOC_CutCopyFile  = ""
		endif
	endif

endwith

ENDPROC


*/-----------------------------------------------------------------------------------------
*/ Descrição   		: Declara todas as variáveis globais do sistema  
*/ Retorno     		: Não há retorno  
*/-----------------------------------------------------------------------------------------
PROCEDURE FOL_DeclarePublicVars
	LOCAL VLC_Parameter, VLN_Content, VLC_ModuloVarName, VLC_String

	with this
		.VOO_Barcode = createobject("CGO_Barcode_C")		&&		*-- inicia a classe do Barcode

		dimension .VOA_AuxArred[29]		&&		*- .VOA_Rounding - Guarda os valores para arredondamento de moedas
		dimension .VOA_007_Numbers[1]	&&		*- Inicialização das informações para Escrita do Extenso
		dimension .VOA_007_Signs[15]	&& 		*- Números e Símbolos para formaçao do Extenso ( a própria função faz a inicialização )

		.VOC_ReturnPageSelect = ""		&&		*- .VOC_ReturnPageSelect - Retorno da procura via Select
		.VOC_UkeyFromForm = ""			&&		*- .VOC_UkeyFromForm - Guarda o Ukey do form para refresh do ToolBar informativo (Status)
		.VOC_UkeyFromCombo = ""			&&		*- .VOC_UkeyFromCombo - Guarda o Ukey do combo para refresh do ToolBar informativo (Status)
		.VON_maxEditMemo = 250			&&		*- .VON_MaxEditMemo - Guarda o máximo de edição do campo Memo
		.VOC_CompanyCode= space(5)		&&		*- .VOC_CompanyCode - Guarda o código da companhia corrente
		.VOC_CompanyName = ""			&&		*- .VOC_CompanyName - Guarda o nome da companhia corrente
		.VOC_CompanyCgc = ""			&&		*- .VOC_CompanyCgc - Guarda o cgc da companhia corrente
		.VOC_CompanyIE = ""				&&		*- .VOC_CompanyIE - Guarda a inscrição estadual da companhia corrente
		.VOC_CompanyCity  = space(20)	&&		*- .VOC_CompanyCity - Guarda a cidade da companhia corrente
		.VOC_CompanyCityN = ""
		.VOC_CompanyState  = space(20)	&&		*- .VOC_CompanyState - Guarda o estado(UF) da companhia corrente
		.VOC_CompanyStateN = ""
		.VOC_CompanyUF = ""
		.VOC_CompanyCountry  = space(20)&&		*- .VOC_CompanyCountry - Guarda o país da companhia corrente
		.VOC_CompanyCountryN = ""
		.VOC_CompanyAddress = ""		&&		*- .VOC_CompanyAddress - Guarda o endereço da companhia corrente
		.VOC_CompanyNumber = ""			&&		*-  .VOC_CompanyNumber	 - Numero da empresa no endereço indicado em VOC_CompanyAdress.
		.VOC_CompanyPhone = ""			&&		*-	.VOC_CompanyPhone - Guarda o número do tefone da empresa.
		.VOC_CompanyDDDPhone = ""		&&		*-  .VOC_CompanyDDDPhone - Prefixo DDD do telefone da empresa.
		.VOC_CompanyNeighborhood = ""	&&		*- .VOC_CompanyNeighborhood - Guarda o bairro da companhia corrente
		.VOC_CompanyZipCode = ""		&&		*- .VOC_CompanyZipCode - Guarda o CEP da companhia corrente
		.VOC_CompanyJC = ""				&&		*- .VOC_CompanyJC - Guarda a Junta Comercial da companhia corrente
		.VOC_CompanyLogo = ""			&&		*- 	.VOC_CompanyLogo - Guarda o arquivo com o Logotipo da empresa para impressão em relatórios
		.VOC_MenuName = ""				&&		*- Variável que guarda o nome do menu a ser executado após o login
		.VON_TipoSenha = 0				&&		*- .VON_TipoSenha
		.VOC_UserName = ""				&&		*- .VOC_UserName - Guarda o nome do usuário corrente
		.VOC_UsrMsgAccount = ""			&&		*- Conta padrão para mensagens do usuário
		.VOC_UserSex = ""				&&		*- .VOC_UserSex - Guarda o sexo do usuário corrente
		.VOC_UserCivil = ""				&&		*- .VOC_UserCivil - Guarda o estado civil do usuário corrente
		.VOC_UserBorn = {}				&&		*- .VOC_UserBorn - Guarda a data de nascimento do usuário corrente
		.VOC_UserCode = ""				&&	*- .VOC_UserCode - Guarda o código do usuário corrente
		.VOC_GroupCode= space(5)
		.VOC_GroupName= ""
		.VOC_UserPassword = ""			&&		*- .VOC_UserPassword - Guarda o password do usuário corrente
		.VOC_WWWHomeStarSoft = .FOC_caption("www_homestar")	&&		*- .VOC_WWWHomeStarSoft - Guarda o endereço eletrônico da StarSoft
		.VOC_WWWHelpStarSoft = .FOC_caption("www_helpstar")	&&		*- .VOC_WWWHelpStarSoft - Guarda o endereço eletrônico do Suporte da StarSoft
		.VOA_CurrencyNames      = ""	&&		*- .VOA_CurrencyNames - Guarda os nomes de moedas
		.VOA_CurrencyNames[1]   =  "R$   "
		.VOA_CurrencyValues     = 0		&&		*- .VOA_CurrencyValues - Guarda os valores para conversão das moedas com índice > 1 para a com 1
		.VOA_CurrencyValues[1]  = 1
		.VOC_CurrencyPlural = "Reais"	&&		*- .VOC_CurrencyPlural - Guarda o nome plural da moeda
		.VOC_CurrencySingular = "Real"	&&		*- .VOC_CurrencySingular - Guarda o nome singular da moeda
		.VOC_DecimalPlural = "Centavos"	&&		*- .VOC_DecimalPlural - Guarda o nome plural da moeda
		.VOC_DecimalSingular = "Centavo"&&		*- .VOC_DecimalSingular - Guarda o nome singular da moeda
		.VOL_CurrencyHaveCents = .T.	&&		*- .VOL_CurrencyHaveCents - Indica se a moeda possui centavos
		.VON_CalcMode = 1				&&		*- .VON_CalcMode - Indica qual calculador os Spinners vão usar 1-Fox, 2-Windows, 3-Star
		.VOC_FormGroup = ""				&&		*- .VOC_FormGroup - Guarda o último grupo de arquivos aberto
		.VON_ThisCallType = FRM_NORMAL	&&		*- .VON_ThisCallType - Guarda o tipo de tela que está sendo chamada
		.VOC_Home = ""					&&		*- .VOC_Home - Guarda o diretório padrão do sistema

		.VOC_PresentModule = ""		&&		*- .VOC_PresentModule - Guarda o módulo atual do sistema
		.VON_ScreenObject = 0			&&		*- .VON_ScreenObject, .VOO_ScreenObject - Número e ponteiro ao objeto da tela
		.VOO_ScreenObject = _screen
		.VOL_True = .T.					&&		*- .VOL_True, .VOL_False - variáveis para os selects
		.VOL_False = .F.
		.VOL_Closeconnection = .T.		&&		*--	Indica se sempre fecha a conexão com o Banco de Dados no término da transação
		.VOL_ShowUniqueKey	= .T.		&&		*--	Indica se mostra o ícone de chave unica nas telas
		.VOC_Currency = space(5)		&&		*- .VOC_Currency - Guarda a moeda padrão do sistema
		.VOC_ScreenName = ""			&&		*- .VOC_ScreenName - Guarda o nome do programa na tela
		.VOD_EmptyDate = {}				&&		*- .VOD_EmptyDate - Data Vazia
		.VOD_MinValidDate = {^1900-01-01}					&&		*- .VOD_MinValidDate, .VOD_MaxValidDate - variáveis para valores máximos e mínimos para data
		.VOD_MaxValidDate = {^2079-01-01}
		.VOC_UserDirectory = "C:\Meus Documentos"			&&		*- .VOC_UserDirectory - Diretório padrão para documentos do usuário
		.VOL_SmartSpinner = .T.								&&		*- .VOL_SmartSpinner - indica se o spinner será tipo calculadora
		.VON_ActiveChar = 0
		.VON_ModemPort = -1									&&		*--	armazena variáveis para discagem via modem	&&	
		.VOC_TypeLine  = "T"
		.VOC_PrefixLine = ""

		.VON_Error = 0			&&		*- .VON_Error - Guarda os numeros referentes aos erros
	endwith
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Control : Campo de controle 
*/                    PLC_UserKey : Unikey do usuário 
*/                    PLL_OnlySize : executa encriptação somente do tamanho 
*/ Descrição        : Função para encriptação ou decriptação do controle do usuário 
*/ Retorno          : Retorna o controle executando um xor do UserKey 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_UserChangeControl
	LParameters PLC_Control, PLC_UserKey, PLL_OnlySize

	LOCAL VLN_i, VLN_j, VLC_NewControl, VLN_Size

	if PLL_OnlySize
		VLN_Size = len(PLC_Control)
	else
		VLN_Size = 100
	endif
	PLC_Control = padr(PLC_Control, VLN_Size)
	VLC_NewControl = ""
	VLN_i = 1
	VLN_j = 1

	*- Varre o controle e cálcula o xor do ukey0
	do while VLN_i <= VLN_Size
		for VLN_j = 1 to 5
			VLC_NewControl = VLC_NewControl + chr(bitxor(asc(substr(PLC_Control, VLN_i, 1)), ;
														 (asc(substr(PLC_UserKey, VLN_j, 1))*VLN_i*123)%32))
			VLN_i = VLN_i + 1											
			if VLN_i > VLN_Size
				exit
			endif
		endfor
	enddo

	return VLC_NewControl
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_oSource      : Referencia ao objeto                      		  
*/                  : PLN_ControlCount : Numero de elementos no Objeto oSource  
*/ Descrição        : Procura por objetos cujos valores estejam em branco 
*/ Retorno          : .T. :Se encontrar algum objeto vazio, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ChkEmptyValue
	LParameters PLO_oSource, PLN_ControlCount

	LOCAL VLN_i, VLN_j, VLO_Control, VLC_Baseclass, VLL_Empty, VLC_Macro

	for VLN_i = 1 to PLN_ControlCount

		VLO_Control = PLO_oSource.Controls[VLN_i]
	*--	Objetos que terao seus valores examinados, no caso de Pageframe e Container
	*--		os seus conteúdos são analisados através de recursividade
		VLC_BaseClass = upper(VLO_Control.Baseclass)
		if !(VLC_BaseClass $ "COMBOBOX/LISTBOX/TEXTBOX/PAGEFRAME/CONTAINER")
			loop
		endif

		do case
			case VLO_Control.Baseclass$"Pageframe"		&& Chamo esta funcão novamente com cada página
				for VLN_j = 1 to VLO_Control.PageCount
					if this.FOL_ChkEmptyValue(VLO_Control.pages[VLN_j], VLO_Control.pages[VLN_j].ControlCount)
						VLO_Control.pages[VLN_j].zorder(0)
						return .T.
					endif
				endfor

			case VLO_Control.Baseclass$"Container"		&& Chamo esta funcao novamente com o container
				if this.FOL_ChkEmptyValue(VLO_Control, VLO_Control.ControlCount)
					VLO_Control.zorder(0)
					return .T.
				endif
			other
				VLL_Empty = .F.
				if type('VLO_Control.VOL_ChkIfEmpty') = "L" and VLO_Control.VOL_ChkIfEmpty
					if type('VLO_Control.VOC_ControlSource') = "C"
						if !empty(VLO_Control.VOC_ControlSource)
							VLL_Empty = empty(eval(VLO_Control.VOC_ControlSource))
						endif
					else
						if !empty(VLO_Control.ControlSource)
							VLL_Empty = empty(eval(VLO_Control.ControlSource))
						endif
					endif

					if 	VLL_Empty
						this.FON_Msg(65)
						VLO_Control.SetFocus()
						return .T.
					endif
				endif
		endcase
	endfor

	return .F.
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Form     : Ponteiro ao form quando este existir 
*/                  : PLO_Source   : Ponteiro ao objeto quando este existir 
*/                  : PLC_CodeId  : Codigo principal para execucao do Trigger  
*/ Descrição        : Localiza no objeto de Triggers e executa o comando  
*/ Retorno          : a propriedade VOL_Ok do trigger 
*/----------------------------------------------------------------------------------------
procedure FOL_Triggers
	lparameters PLO_Form, PLO_Source, PLC_CodeId

	return this.VOO_Security.FOL_Triggers(PLO_Form, PLO_Source, PLC_CodeId)
endproc

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Source : Ponteiro ao objeto 
*/                    PLN_Tam    : Tamanho do retorno 
*/ Descrição        : Retorna o nome completo do objeto até o form  
*/ Retorno          : Nome completo do objeto  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_ObjectName
	LParameters PLO_Source, PLN_Tam

	LOCAL VLC_Name

	*- Se o source for null retorna string vazia
	if isnull(PLO_Source)
		return ""
	endif
	VLC_Name = PLO_Source.name
	do while upper(alltrim(PLO_Source.baseclass)) <> "FORM"
		PLO_Source = PLO_Source.parent
		VLC_Name = PLO_Source.name + "." + VLC_Name
	enddo

	if !empty(PLN_Tam)
		VLC_Name = padr(substr(VLC_Name, 1, PLN_Tam), PLN_Tam)
	endif

	return VLC_Name
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLO_Source : Ponteiro ao objeto 
*/ Descrição        : Retorna o objeto form em que o objeto está declarado  
*/ Retorno          : Objeto form  
*/----------------------------------------------------------------------------------------
PROCEDURE FOO_ThisForm
	LParameters PLO_Source

	do while upper(alltrim(PLO_Source.baseclass)) <> "FORM"
		PLO_Source = PLO_Source.parent
	enddo

	return PLO_Source
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_MenuID   	: Código do Menu 
*/ Descrição        : Verifica o acesso a um determinado item de menu 
*/ Retorno          : .F. : Se tem acesso, .T. : c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_AcessMenu
	lparameters PLC_MenuID

 	LOCAL VLL_Ret
	*- Verifica se o código do usuário é o administrador
	if this.VOC_UserCode == "STAR_"
		return .f.
	endif

	VLC_StringToFind = ";" + alltrim(PLC_MenuID) + ";"
	*- Verifico se o acesso foi dado ao usuário
	VLL_Ret = at(VLC_StringToFind, this.VOA_AccessCfg[1, 1]) > 0
	if !VLL_Ret
		*- Verifico se houve acesso liberado para o grupo
		VLL_Ret = at(VLC_StringToFind, this.VOA_AccessCfg[1, 3]) > 0
		if VLL_Ret
			*- Se foi liberado para o grupo, verifico se houve restrição para esse usuário
			VLL_Ret = !at(VLC_StringToFind, this.VOA_AccessCfg[1, 2]) > 0
		endif
	endif

	return !VLL_Ret


ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_MenuID   : Código do Menu 
*/                    PLC_UserCode : Código  do Usuário 
*/                    PLC_GroupId  : Código do Grupo de Usuários  
*/ Descrição        : Verifica o acesso para o toolbar  
*/ Retorno          : String de 4 posições: Acesso, Edita, Novo, Deleta 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_AcessToolbar
	LParameters PLC_MenuID, PLC_UserCode, PLC_GroupId

	LOCAL VLC_Ret, VLC_Ret1, VLC_Ret2, VLL_User
	*- Abre o arquivo que as definições dos Parâmetros
	*- Se não houver usuário ou grupo estabelecido, pego os padrões
	if empty(PLC_UserCode)
		PLC_UserCode = this.VOC_UserCode
	endif

	*- Verifica se o código do usuário é o administrador
	if PLC_UserCode == "STAR_"
		return "1111"
	else
		VLC_Ret = "2222"
	endif

	VLC_StringToFind = ";" + alltrim(PLC_MenuID) + ";"

	for VLN_Counter = 1 to 4
		*- Verifico se o acesso foi dado ao usuário
		VLL_Ret = at(VLC_StringToFind, this.VOA_AccessCfg[VLN_Counter, 1]) > 0
		if !VLL_Ret
			*- Verifico se houve acesso liberado para o grupo
			VLL_Ret = at(VLC_StringToFind, this.VOA_AccessCfg[VLN_Counter, 3]) > 0
			if VLL_Ret
				*- Se foi liberado para o grupo, verifico se houve restrição para esse usuário
				VLL_Ret = !at(VLC_StringToFind, this.VOA_AccessCfg[VLN_Counter, 2]) > 0
			endif
		endif

		if VLL_Ret
			VLC_Ret = stuff(VLC_Ret, VLN_Counter, 1, "1")
		endif
	endfor

	return VLC_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_String    : String para eliminar os acentos 
*/                    PLC_ClearToo  : Eliminar estes caracteres e trocar pelos  
*/                                    PLC_ClearWith  
*/                    PLC_ClearWith : Caracteres Novos para substituição dos PLC_ClearToo 
*/ Descrição        : Troca caracteres numa determinada string  
*/ Retorno          : O String Limpo 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_ClearString
	LParameters PLC_String, PLC_ClearToo, PLC_ClearWith

	LOCAL  VLN_c, VLN_pos, VLC_letra, VLC_newSt
	if empty(PLC_ClearToo)
		PLC_ClearToo = ""
	endif
	if empty(PLC_ClearWith)
		PLC_ClearWith= ""
	endif

*-	Caso esteja vazio retorno a variavel para não dar erro no for - Atorre
	if empty(nvl(PLC_String,""))
		return PLC_String
	endif

	VLC_newSt   = ""
	VLC_st1 	= "âáéíóúÁÉÍÓÚãõÂÃÕàÀçÇêôÊÔªº" + PLC_ClearToo
	VLC_st2 	= "aaeiouAEIOUaoAAOaAcCeoEO.." + PLC_ClearWith

	VLC_newSt = chrtran(PLC_String, VLC_st1, VLC_st2)
*- A função acima substitui esse for/next

*-		for VLN_c = 1 to Len( PLC_String )

*-			VLC_letra = substr(PLC_String,VLN_c, 1)
*-			VLN_pos   = at(VLC_letra , VLC_st1  )

*-			if VLN_pos > 0
*-				VLC_letra = substr(VLC_st2, VLN_pos, 1)
*-			endif
*-			VLC_newSt = VLC_newSt + VLC_letra

*-		next

	return VLC_newSt
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_Seconds : Segundos 
*/ Descrição        : Retorna o horário a partir de um número de segundos 
*/ Retorno          : Hora 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_SecondsinDayTime
	LParameters PLN_Seconds

	return trans( int(PLN_Seconds / 86400), "@l 9999") + " D " +;
		STRTRAN(str(int(mod(PLN_Seconds / 3600, 24)), 2) +;
		":" + str(int(mod(PLN_Seconds / 60, 60)), 2) +;
		":" + str(int(mod(PLN_Seconds, 60)), 2), " ", "0" )
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Data    : Data 
*/ Descrição        : Verifica se a data é válida e retorna uma válida  
*/                    para SQL-Server o limite é 01/01/1900 - 01/06/2079  
*/ Retorno          : Vazio - quando PLD_Data é vazia, NULL - quando PLD_Data é NULL, 
*/                    PLD_Data - quando PLD_Data está no intervalo válido, min/max - c.c. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_CheckDate
	LParameters PLD_Data

	if isnull(PLD_Data) or empty(PLD_Data) or between(PLD_Data, this.VOD_MinValidDate, this.VOD_MaxValidDate)
		return PLD_Data
	endif
	if PLD_Data < this.VOD_MinValidDate
		return this.VOD_MinValidDate
	endif
	return this.VOD_MaxValidDate
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Data    : Data 
*/                  : PLL_DiaUtil : Avança para o Primeiro dia util 
*/                  : PLN_Fator   : Fator avancar ou retornar o mes 
*/ Descrição        : Verifica o primeiro dia util da Data  
*/ Retorno          : Data 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_FirstDayDate
	LParameters PLD_Data, PLL_DiaUtil, PLN_Fator

	if empty(PLD_Data)
		PLD_Data = date()
	else
		PLD_Data = ctod(dtoc(PLD_Data))
	endif
	if empty(PLN_Fator)
		PLN_Fator = 0
	endif
	
	VLD_rt = PLD_Data - day( PLD_Data ) + 1
	
	if PLL_diaUtil and !empty(VLD_rt)
	   return this.FOD_ProximoDiaUtil( VLD_rt, "", "", "", 1 )
	else
	   return gomonth(VLD_rt, PLN_Fator)
	endif
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Data    : Data 
*/                    PLL_DiaUtil : Avança para o Último dia util 
*/                    PLN_Fator   : Fator avancar ou retornar o mês 
*/ Descrição        : Verifica o Último dia da Data 
*/ Retorno          : Data 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_LastDayDate
	LParameters PLD_Data, PLL_DiaUtil, PLN_Fator
	
	if empty(PLN_Fator)
		PLN_Fator = 0
	endif
	if empty(PLD_Data)
		PLD_Data = date()
	else
		PLD_Data = ctod(dtoc(PLD_Data))
	endif

	VLD_rt = gomonth(PLD_Data, 1) - day(gomonth(PLD_data, 1))
	
	if PLL_diaUTIL and !empty(VLD_rt)
	   return this.FOD_ProximoDiaUtil( VLD_rt, "", "", "", -1 )
	else
	   return gomonth(VLD_rt, PLN_Fator)
	endif
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Date : Data a ser convertida  
*/ Descrição        : Retorna a data no formato caracter AnoMes 
*/ Retorno          : Retorna a string           	  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_AnoMes
	LParameters PLD_Date

	return iif(!empty(nvl(PLD_Date, "")), substr(dtos(PLD_Date), 1, 6), space(6))
	
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros         PLD_Data         : Data  
*/                    PLC_pais         : País  
*/                    PLC_uf           : Estado  
*/                    PLC_cidade       : Cidade  
*/                    PLN_FatorFeriado : Fator avancar ou retornar em feriado (1=+)/(-1=-)
*/                    PLN_FatorWeekEnd : Fator avancar ou retornar o (1=SAB)/(-1=DOM) 
*/                    PLL_Feriado      : Considera feriado como dia útil  
*/                    PLL_WeekEnd      : Considera fim de semana como dia útil  
*/ Descrição        : Verifica o primeiro dia útil da Data  
*/ Retorno          : Data 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_ProximoDiaUtil
	LParameters PLD_Data, PLC_Pais, PLC_uf , PLC_Cidade, PLN_FatorFeriado, PLN_FatorWeekEnd, PLL_Feriado, PLL_WeekEnd

	if empty( PLC_Pais)
	   PLC_Pais = this.VOC_CompanyCountry
	endif
	if empty( PLC_uf)
	  PLC_uf   = this.VOC_CompanyState
	endif
	if empty( PLC_cidade)
	   PLC_cidade = this.VOC_CompanyCity
	endif

	if empty(PLN_fatorWeekEnd)
	   PLN_fatorWeekEnd = 1
	endif

	if PLN_fatorWeekEnd = 1
	   VLN_auxSAB = 2
	   VLN_auxDOM = 1
	else
	   VLN_auxSAB = -1
	   VLN_auxDOM = -2
	endif

	if empty(PLN_fatorFeriado)
	   PLN_fatorFeriado = 1
	endif

	if PLN_fatorFeriado = 1
	   VLN_auxSABFeriado = 2
	   VLN_auxDOMFeriado = 1
	else
	   VLN_auxSABFeriado = -1
	   VLN_auxDOMFeriado = -2
	endif

	if PLL_WeekEnd
	 VLN_Dia = dow(PLD_Data)                 && dia numerico da semana
	 if VLN_dia = 7
	       PLD_data = ctod(dtoc(PLD_data)) + VLN_auxSAB  && sabado avanca 2 dias ou Retorna 1
	 else
	    if VLN_dia = 1
	       PLD_data = ctod(dtoc(PLD_data)) + VLN_auxDOM  && domingo avanca 1 dia ou Retorna 2
	    endif
	 endif
	endif

	if PLL_Feriado
	 do while this.FOL_ChkDaTFeriado( PLC_pais, PLC_uf , PLC_cidade , PLD_data )
	     PLD_data = ctod(dtoc(PLD_data)) + PLN_fatorFeriado
	     VLN_dia  = dow(PLD_data)
	     if VLN_dia = 7
	           PLD_data = ctod(dtoc(PLD_data)) + VLN_auxSABFeriado
	     else
	        if VLN_dia = 1
	           PLD_data = ctod(dtoc(PLD_data)) + VLN_auxDOMFeriado
	        endif
	     endif
	 enddo
	endif

	return PLD_data
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Pais   : País  
*/                  : PLC_Uf     : Estado  
*/                  : PLC_Cidade : Cidade  
*/                  : PLD_Data   : Data a ser verificada  
*/                  : PLL_Ponte  : Verifica se é um Feriado Ponte 
*/ Descrição        : Verifica se existe feriado na data  
*/ Retorno          : .T. : Se for Feriado,  .F. : c.c. 
*/ Última Alteração : 21/09/99 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ChkDatFeriado
	LParameters PLC_Pais, PLC_UF, PLC_Cidade, PLD_Data, PLL_Ponte

	LOCAL VLN_Day, VLN_Month, VLN_DayOrder, VLN_Dow, VLN_TmpMonth, VLL_Return

	if empty(PLD_Data)
		PLD_Data = date()
	endif

	if empty( PLC_Pais)
	   PLC_Pais = this.VOC_CompanyCountry
	endif

	if empty( PLC_UF)
	  PLC_UF   =   this.VOC_CompanyState
	endif

	if empty( PLC_Cidade)
	   PLC_Cidade = this.VOC_CompanyCity
	endif

	VLN_Day			= day(PLD_Data)
	VLN_Month		= month(PLD_Data)
	VLN_DayOrder	= 0
	VLN_Dow			= dow(PLD_Data)
	VLN_TmpMonth	= VLN_Month

	do while VLN_TmpMonth == VLN_Month
		VLN_DayOrder	= VLN_DayOrder + 1
		VLN_TmpMonth = month(PLD_Data - (VLN_DayOrder * 7))
	enddo

	with this
		.FOL_SetParameter(1, PLC_Pais)
		.FOL_SetParameter(2, PLC_UF)
		.FOL_SetParameter(3, PLC_Cidade)
		.FOL_SetParameter(4, VLN_Day)
		.FOL_SetParameter(5, VLN_Month)
		.FOL_SetParameter(6, VLN_DayOrder)
		.FOL_SetParameter(7, VLN_Dow)
		.FOL_SetParameter(8, PLD_Data)
		.FOL_SetParameter(9, iif(PLL_Ponte, 1, 0))
		.FOL_ExecuteCursor("A26_CHKDATFER", "TMP_FERIADO", 9)
	endwith

	VLL_Return = (TMP_Feriado.reccount > 0)

	use in TMP_Feriado

	return VLL_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Ini    : Data inicial  
*/                    PLD_Fim    : Data final  
*/                    PLL_IniFim : Se calcula do início ao fim  
*/ Descrição        : Retorna o número de dsrs entre duas datas 
*/ Retorno          : número de dsrs             	  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Dsrs
	LParameters PLD_Ini, PLD_Fim, PLL_IniFim
	
	LOCAL VLN_Ret, VLD_Aux

	this.FOL_PushSelect()
	
	if empty(PLD_Ini) and empty(PLD_Fim)
		PLL_IniFim = .T.
	endif
	if empty(PLD_Ini)
		PLD_Ini = VPD_DataRH
	endif
	if empty(PLD_Fim)
		PLD_Fim = PLD_Ini
	endif
	
	if PLL_IniFim
		PLD_Ini = PLD_Ini - day(PLD_Ini) + 1
		PLD_Fim = gomonth(PLD_Fim, 1) - day(gomonth(PLD_Fim, 1))
	endif
	
	VLN_Ret = 0
	
	do while PLD_Ini <= PLD_Fim and !empty(PLD_Ini)
		if !seek(this.FOC_AnoMes(PLD_Ini), "A49", "A49_001_C")
			exit
		else
			VLD_Aux = gomonth(PLD_Ini - day(PLD_Ini) + 1, 1)-1
			if VLD_Aux > PLD_Fim
				VLD_Aux = PLD_Fim
	        endif
	
			VLN_Ret = VLN_Ret + occurs('1', substr( A49.A49_002_C, day(PLD_Ini), VLD_Aux - PLD_Ini + 1 ) )
	
	        PLD_Ini = VLD_Aux + 1
	 	endif
	enddo
	
	this.FOL_PopSelect()
	
	return VLN_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Data : data para conversão     	  
*/ Descrição        : Converte uma string data no formato dtos para uma variável tipo data
*/ Retorno          : Retorna o valor convertido 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_Stod
	LParameters PLC_Data

	local VLD_Date
	VLD_Date = {}

	if !empty(PLC_Data)

		VLN_year  = val(subs(PLC_Data, 1, 4))
		VLN_Month = val(subs(PLC_Data, 5, 2))
		VLN_Day   = val(subs(PLC_Data, 7, 2))
		
		if between(VLN_year, 100, 9999) and between(VLN_Month , 1, 12) and between(VLN_Day, 1, 31) 
			VLD_Date = date(VLN_year  ,VLN_Month , VLN_Day  ) 
		endif
	endif
	
	return VLD_Date
	
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_Valor   : Valor para conversão  
*/                    PLC_Morig   : Moeda original  
*/                    PLC_Mconv   : Moeda para conversão  
*/                    PLD_Data    : Data para conversão 
*/                    PLN_Round   : numero de Arredondamento ROUND  
*/                    PLL_SetNear : Procura a moeda na data próxima, se encontrar volta 
*/                                  a data em PLD_Data. 
*/                    PLL_ValProj : Valor a ser verificado se e o Projetado 
*/                    PLL_Warning : Indica se executa mensagens de warning  
*/ Descrição        : Converte os valores de uma moeda para outra moeda 
*/ Retorno          : Retorna o valor convertido 
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Vm
	LParameters PLN_Valor, PLC_MOrig, PLC_MConv, PLD_Data, PLN_Round, PLL_SetNear, PLL_ValProj,;
				PLL_Warning

	LOCAL VLN_Morig, VLN_Mconv, VLC_MOrig, VLC_MConv, VLC_DOrig, VLC_DConv, VLD_Fim, VLL_Inv,;
		  VLN_Aux, VLC_Tag, VLC_Data, VLL_Update, VLN_Parameters, VLN_Count, VLN_OldCount,;
		  VLN_CountA54, VLN_LastA54, VLN_CountA37, VLN_LastA37, VLN_OriginalValue, VLN_Projection, VLD_FirstDay, VLD_LastDay

	this.VOL_ExchangeError = .F.
	VLN_OriginalValue = PLN_Valor
	VLN_Projection = 0

	VLD_FirstDay = {}
	VLD_LastDay = {}

	VLN_Parameters = parameters()

	if empty(nvl(PLC_MOrig, "")) or empty(nvl(PLC_MConv, "")) or empty(PLN_Valor)
		return 0
	endif

	*- Guarda as moedas e datas
	VLC_MOrig = subs(PLC_MOrig, 1, 5)
	VLC_MConv = subs(PLC_MConv, 1, 5)
	VLC_DOrig = subs(PLC_MOrig, 6)
	VLC_DConv = subs(PLC_MConv, 6)
	if empty(PLD_Data)
		PLD_Data = iif(empty(VLC_DConv), date(), this.FOD_Stod(VLC_DConv))
	endif
	VLC_Data  = dtos(PLD_Data)

	if empty(VLC_DOrig)
		VLC_DOrig = VLC_Data
	endif

	if empty(VLC_DConv)
		VLC_DConv = VLC_Data
	endif

	*- Se não for igual considera a data passada
	if !(VLC_MOrig == VLC_MConv)
		if !empty(PLD_Data)
			VLC_DConv = VLC_Data
		endif
	endif

	*- Retorna o valor original se são iguais
	if VLC_MOrig == VLC_MConv and VLC_DOrig == VLC_DConv
		return PLN_Valor
	endif

	this.FOL_PushSelect()

	VLD_ConvDate = this.FOD_Stod(VLC_DConv)

	if VLC_DOrig < VLC_DConv
		VLD_OrigDate = this.FOD_Stod(VLC_DOrig)
		VLD_Fim = this.FOD_Stod(VLC_DConv)
		VLL_Inv = .F.
	else
		VLD_OrigDate = this.FOD_Stod(VLC_DConv)
		VLD_Fim = this.FOD_Stod(VLC_DOrig)
		VLL_Inv = .T.
	endif

	VLN_OldCount  = iif(empty(this.VOA_LastA54), 0, alen(this.VOA_LastA54, 1))
	VLN_LastA54 = Ceil( Ascan(this.VOA_LastA54, VLC_MOrig + VLC_DOrig + VLC_DConv) / 4 )
	if VLN_LastA54 = 0
		* Pesquisa conversão no próprio sistema monetário
		with this
			.FOL_SetParameter(1, VLC_MOrig)
			.FOL_SetParameter(2, VLD_OrigDate)
			.FOL_SetParameter(3, VLD_Fim)
			.FOL_ExecuteCursor("A54_CONVERT", "A54T", 3)
		endwith

		VLN_CountA54  = reccount("A54T")

		dimension  this.VOA_LastA54[iif(VLN_CountA54 = 0, 1, VLN_CountA54) + VLN_OldCount, 4]

		if VLN_CountA54 > 1
			select A54T
			go top in A54T
			for VLN_Count = 1 to VLN_CountA54
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 1] = VLC_MOrig + VLC_DOrig + VLC_DConv
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 2] = a54t.a54_006_n
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 3] = a54t.a54_005_b
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 4] = .f. or (VLL_Inv and VLN_Count = VLN_CountA54)
				skip in A54T
			next
		else
			this.VOA_LastA54[VLN_OldCount + 1, 1] = VLC_MOrig + VLC_DOrig + VLC_DConv
			this.VOA_LastA54[VLN_OldCount + 1, 4] = .t.
		endif
		VLN_LastA54 = VLN_OldCount + 1
    endif

	this.FOL_CloseTable("A54T")

    for VLN_Count = VLN_LastA54 to alen(this.VOA_LastA54, 1)
		if VLC_MOrig + VLC_DOrig + VLC_DConv <> this.VOA_LastA54[VLN_Count, 1] or;
			this.VOA_LastA54[VLN_Count, 4]
			exit
		endif
		do case
			case this.VOA_LastA54[ VLN_Count, 2] = 1  && -
				if VLL_Inv
					PLN_Valor = PLN_Valor + this.VOA_LastA54[ VLN_Count, 3]
				else
					PLN_Valor = PLN_Valor - this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 2  && +
				if VLL_Inv
					PLN_Valor = PLN_Valor - this.VOA_LastA54[ VLN_Count, 3]
				else
					PLN_Valor = PLN_Valor + this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 3  && *
				if VLL_Inv
					if this.VOA_LastA54[ VLN_Count, 3] <> 0
						PLN_Valor = PLN_Valor / this.VOA_LastA54[ VLN_Count, 3]
					else
						PLN_Valor = 0
					endif
				else
					PLN_Valor = PLN_Valor * this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 4  && /
				if VLL_Inv
					PLN_Valor = PLN_Valor * this.VOA_LastA54[ VLN_Count, 3]
				else
					if this.VOA_LastA54[ VLN_Count, 3] <> 0
						PLN_Valor = PLN_Valor / this.VOA_LastA54[ VLN_Count, 3]
					else
						PLN_Valor = 0
					endif
				endif
		endcase
    next

	if !(VLC_MOrig == VLC_MConv)
		VLN_OldCount  = iif(empty(this.VOA_LastA37), 0, alen(this.VOA_LastA37, 1))
		VLN_LastA37 = Ceil( Ascan(this.VOA_LastA37, Iif(this.VOL_CurrencyProjection, "2", iif(PLL_SetNear, "1", "0")) + VLC_MOrig + VLC_MConv + VLC_DConv ) / 5)
		if VLN_LastA37 = 0 or this.VOL_CurrencyProjection
			with this
				.FOL_SetParameter(1, VLC_MOrig)
				.FOL_SetParameter(2, VLD_Fim)
				.FOL_ExecuteCursor("A54_A36_UKEY", "A54T", 2)
				VLL_ConversaoMensal = !empty(a54t.a54_013_n)
				if !VLL_ConversaoMensal
					.FOL_SetParameter(1, VLC_MConv)
					.FOL_ExecuteCursor("A54_A36_UKEY", "A54T", 2)
					VLL_ConversaoMensal = !empty(a54t.a54_013_n)
				endif
				.FOL_CloseTable("A54T")

				if VLL_ConversaoMensal
					VLD_FirstDay = date(year(VLD_Fim), month(VLD_Fim), 1)
					VLD_LastDay = gomonth(VLD_Fim, 1) - day(gomonth(VLD_Fim, 1))
					.FOL_SetParameter(1, VLC_MOrig)
					.FOL_SetParameter(2, VLC_MConv)
					.FOL_SetParameter(3, VLD_FirstDay)
					.FOL_SetParameter(4, VLD_LastDay)
					.FOL_ExecuteCursor("A37_A36_UKEY01", "A37T", 4)
				else
					* Procura conversão entre sistemas monetários
					.FOL_SetParameter(1, VLC_MOrig)
					.FOL_SetParameter(2, VLC_MConv)
					.FOL_SetParameter(3, VLD_ConvDate)
					.FOL_ExecuteCursor("A37_A36_UKEY", "A37T", 3)
				endif
				select A37T
				
				*-- 
				if Eof("A37t") and this.VOL_CurrencyProjection
					VLN_Projection = this.FON_VmP(VLN_OriginalValue, PLC_MOrig, PLC_MConv, PLD_Data, PLN_Round, PLL_Warning)
				endif

				* Se não achou conversão com a data efetua pesquisa com Near
				if VLN_Projection = 0 and PLL_SetNear
					use in A37T
					.FOL_SetParameter(1, VLC_MOrig)
					.FOL_SetParameter(2, VLC_MConv)
					.FOL_SetParameter(3, VLD_ConvDate)
					.FOL_ExecuteCursor("A37_A36_NEAR", "A37T", 3)
				endif
				VLN_CountA37  = reccount("A37T")

				dimension  .VOA_LastA37[iif(VLN_CountA37 = 0, 1, VLN_CountA37) + VLN_OldCount, 5]

				VLN_LastA37 = VLN_OldCount + 1
				if VLN_Projection = 0
					.VOA_LastA37[VLN_LastA37, 1] = iif(PLL_SetNear, "1", "0") + VLC_MOrig + VLC_MConv + VLC_DConv
					.VOA_LastA37[VLN_LastA37, 2] = A37T.a37_002_b
					.VOA_LastA37[VLN_LastA37, 3] = A37T.a37_003_b
					.VOA_LastA37[VLN_LastA37, 4] = A37T.a37_005_n
					.VOA_LastA37[VLN_LastA37, 5] = VLN_CountA37 > 0
				else
					.VOA_LastA37[VLN_LastA37, 1] = .VOA_LastA70[VLN_Projection, 1]
					.VOA_LastA37[VLN_LastA37, 2] = .VOA_LastA70[VLN_Projection, 2]
					.VOA_LastA37[VLN_LastA37, 3] = .VOA_LastA70[VLN_Projection, 3]
					.VOA_LastA37[VLN_LastA37, 4] = .VOA_LastA70[VLN_Projection, 4]
					.VOA_LastA37[VLN_LastA37, 5] = .VOA_LastA70[VLN_Projection, 5]
				endif
			endwith
	    endif

		if this.VOA_LastA37[VLN_LastA37, 5]
			VLN_Aux = this.VOA_LastA37[VLN_LastA37, 2]

			do case
				case this.VOA_LastA37[VLN_LastA37, 4] = 1  && -
					PLN_Valor = PLN_Valor - VLN_Aux

				case this.VOA_LastA37[VLN_LastA37, 4] = 2  && +
					PLN_Valor = PLN_Valor + VLN_Aux

				case this.VOA_LastA37[VLN_LastA37, 4] = 3  && *
					PLN_Valor = PLN_Valor * VLN_Aux

				case this.VOA_LastA37[VLN_LastA37, 4] = 4  && /
					if VLN_Aux <> 0
						PLN_Valor = PLN_Valor / VLN_Aux
					else
						PLN_Valor = 0
					endif
			endcase
		else
			if PLL_Warning
				this.FON_Msg(262) && Conversão não cadastrada!
			endif
			this.VOL_ExchangeError = .T.  && Guarda o Valor do Msg - Para ser tratado por fora.
			PLN_Valor = 0
		endif

		this.FOL_CloseTable("A37T")
	endif

	PLN_Round = iif((VLN_Parameters > 4), PLN_Round, .F.)

	this.FOL_PopSelect()
	
	return this.FON_Round(PLN_Valor, PLN_Round)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_Valor   : Valor para conversão  
*/                    PLC_Morig   : Moeda original  
*/                    PLC_Mconv   : Moeda para conversão  
*/                    PLD_Data    : Data para conversão 
*/                    PLN_Round   : numero de Arredondamento ROUND  
*/                    PLL_SetNear : Procura a moeda na data próxima, se encontrar volta 
*/                                  a data em PLD_Data. 
*/                    PLL_Warning : Indica se executa mensagens de warning  
*/ Descrição        : Conversão sobre a projeção de moeda 
*/ Retorno          : Retorna o valor convertido 
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Vmp
	LParameters PLN_Valor, PLC_MOrig, PLC_MConv, PLD_Data, PLN_Round, PLL_Warning, PLL_InfoA37


	LOCAL VLN_Morig, VLN_Mconv, VLC_MOrig, VLC_MConv, VLC_DOrig, VLC_DConv, VLD_Fim, VLL_Inv,;
		  VLN_Aux, VLC_Tag, VLC_Data, VLL_Update, VLN_Parameters, VLN_Count, VLN_OldCount,;
		  VLN_CountA54, VLN_LastA54, VLN_CountA70, VLN_LastA70, VLD_FirstDay, VLD_LastDay


	this.VOL_ExchangeError = .F.

	VLN_Parameters = parameters()
	VLD_FirstDay = {}
	VLD_LastDay = {}

	if empty(nvl(PLC_MOrig, "")) or empty(nvl(PLC_MConv, "")) or empty(PLN_Valor)
		return 0
	endif

	*- Guarda as moedas e datas
	VLC_MOrig = subs(PLC_MOrig, 1, 5)
	VLC_MConv = subs(PLC_MConv, 1, 5)
	VLC_DOrig = subs(PLC_MOrig, 6)
	VLC_DConv = subs(PLC_MConv, 6)
	if empty(PLD_Data)
		PLD_Data = iif(empty(VLC_DConv), date(), this.FOD_Stod(VLC_DConv))
	endif
	VLC_Data  = dtos(PLD_Data)

	if empty(VLC_DOrig)
		VLC_DOrig = VLC_Data
	endif

	if empty(VLC_DConv)
		VLC_DConv = VLC_Data
	endif

	*- Se não for igual considera a data passada
	if !(VLC_MOrig == VLC_MConv)
		if !empty(PLD_Data)
			VLC_DConv = VLC_Data
		endif
	endif

	*- Retorna o valor original se são iguais
	if VLC_MOrig == VLC_MConv and VLC_DOrig == VLC_DConv
		return PLN_Valor
	endif

	this.FOL_PushSelect()

	VLD_ConvDate = this.FOD_Stod(VLC_DConv)

	if VLC_DOrig < VLC_DConv
		VLD_OrigDate = this.FOD_Stod(VLC_DOrig)
		VLD_Fim = this.FOD_Stod(VLC_DConv)
		VLL_Inv = .F.
	else
		VLD_OrigDate = this.FOD_Stod(VLC_DConv)
		VLD_Fim = this.FOD_Stod(VLC_DOrig)
		VLL_Inv = .T.
	endif

	VLN_OldCount  = iif(empty(this.VOA_LastA54), 0, alen(this.VOA_LastA54, 1))
	VLN_LastA54 = Ceil( Ascan(this.VOA_LastA54, VLC_MOrig + VLC_DOrig + VLC_DConv) / 4 )
	if VLN_LastA54 = 0
		* Pesquisa conversão no próprio sistema monetário
		with this
			.FOL_SetParameter(1, VLC_MOrig)
			.FOL_SetParameter(2, VLD_OrigDate)
			.FOL_SetParameter(3, VLD_Fim)
			.FOL_ExecuteCursor("A54_CONVERT", "A54T", 3)
		endwith

		VLN_CountA54  = reccount("A54T")

		dimension  this.VOA_LastA54[iif(VLN_CountA54 = 0, 1, VLN_CountA54) + VLN_OldCount, 4]

		if VLN_CountA54 > 1
			select A54T
			go top in A54T
			for VLN_Count = 1 to VLN_CountA54
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 1] = VLC_MOrig + VLC_DOrig + VLC_DConv
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 2] = a54t.a54_006_n
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 3] = a54t.a54_005_b
				this.VOA_LastA54[VLN_OldCount + VLN_Count, 4] = .f. or (VLL_Inv and VLN_Count = VLN_CountA54)
				skip in A54T
			next
		else
			this.VOA_LastA54[VLN_OldCount + 1, 1] = VLC_MOrig + VLC_DOrig + VLC_DConv
			this.VOA_LastA54[VLN_OldCount + 1, 4] = .t.
		endif
		VLN_LastA54 = VLN_OldCount + 1
    endif

	this.FOL_CloseTable("A54T")

    for VLN_Count = VLN_LastA54 to alen(this.VOA_LastA54, 1)
		if VLC_MOrig + VLC_DOrig + VLC_DConv <> this.VOA_LastA54[VLN_Count, 1] or;
			this.VOA_LastA54[VLN_Count, 4]
			exit
		endif
		do case
			case this.VOA_LastA54[ VLN_Count, 2] = 1  && -
				if VLL_Inv
					PLN_Valor = PLN_Valor + this.VOA_LastA54[ VLN_Count, 3]
				else
					PLN_Valor = PLN_Valor - this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 2  && +
				if VLL_Inv
					PLN_Valor = PLN_Valor - this.VOA_LastA54[ VLN_Count, 3]
				else
					PLN_Valor = PLN_Valor + this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 3  && *
				if VLL_Inv
					if this.VOA_LastA54[ VLN_Count, 3] <> 0
						PLN_Valor = PLN_Valor / this.VOA_LastA54[ VLN_Count, 3]
					else
						PLN_Valor = 0
					endif
				else
					PLN_Valor = PLN_Valor * this.VOA_LastA54[ VLN_Count, 3]
				endif

			case this.VOA_LastA54[ VLN_Count, 2] = 4  && /
				if VLL_Inv
					PLN_Valor = PLN_Valor * this.VOA_LastA54[ VLN_Count, 3]
				else
					if this.VOA_LastA54[ VLN_Count, 3] <> 0
						PLN_Valor = PLN_Valor / this.VOA_LastA54[ VLN_Count, 3]
					else
						PLN_Valor = 0
					endif
				endif
		endcase
    next
	
	if !(VLC_MOrig == VLC_MConv)
		VLN_OldCount  = iif(empty(this.VOA_LastA70), 0, alen(this.VOA_LastA70, 1))
		VLN_LastA70 = Ceil( Ascan(this.VOA_LastA70, "0" + VLC_MOrig + VLC_MConv + VLC_DConv ) / 5)
		if VLN_LastA70 = 0
			with this
				.FOL_SetParameter(1, VLC_MOrig)
				.FOL_SetParameter(2, VLD_Fim)
				.FOL_ExecuteCursor("A54_A36_UKEY", "A54T", 2)
				VLL_ConversaoMensal = !empty(a54t.a54_013_n)
				if !VLL_ConversaoMensal
					.FOL_SetParameter(1, VLC_MConv)
					.FOL_ExecuteCursor("A54_A36_UKEY", "A54T", 2)
					VLL_ConversaoMensal = !empty(a54t.a54_013_n)
				endif
				.FOL_CloseTable("A54T")

				if VLL_ConversaoMensal
					VLD_FirstDay = date(year(VLD_Fim), month(VLD_Fim), 1)
					VLD_LastDay = gomonth(VLD_Fim, 1) - day(gomonth(VLD_Fim, 1))
					.FOL_SetParameter(1, VLC_MOrig)
					.FOL_SetParameter(2, VLC_MConv)
					.FOL_SetParameter(3, VLD_FirstDay)
					.FOL_SetParameter(4, VLD_LastDay)
					.FOL_ExecuteCursor("A70_A36_UKEY01", "A70T", 4)
				else
					* Procura conversão entre sistemas monetários
					.FOL_SetParameter(1, VLC_MOrig)
					.FOL_SetParameter(2, VLC_MConv)
					.FOL_SetParameter(3, VLD_ConvDate)
					.FOL_ExecuteCursor("A70_A36_UKEY", "A70T", 3)
				endif
				select A70T

				* Se não achou conversão com a data efetua pesquisa com Near
				VLN_CountA70 = reccount("A70T")

				dimension  .VOA_LastA70[Max(1, VLN_CountA70) + VLN_OldCount, 5]

				VLN_LastA70 = VLN_OldCount + 1
				.VOA_LastA70[VLN_LastA70, 1] = "2" + VLC_MOrig + VLC_MConv + VLC_DConv
				.VOA_LastA70[VLN_LastA70, 2] = A70T.a70_002_b
				.VOA_LastA70[VLN_LastA70, 3] = 0
				.VOA_LastA70[VLN_LastA70, 4] = A70T.a70_004_n
				.VOA_LastA70[VLN_LastA70, 5] = VLN_Counta70 > 0
				
			endwith
	    endif

		if this.VOA_LastA70[VLN_LastA70, 5]
			VLN_Aux = this.VOA_LastA70[VLN_LastA70, 2]

			do case
				case this.VOA_LastA70[VLN_LastA70, 4] = 1  && -
					PLN_Valor = PLN_Valor - VLN_Aux

				case this.VOA_LastA70[VLN_LastA70, 4] = 2  && +
					PLN_Valor = PLN_Valor + VLN_Aux

				case this.VOA_LastA70[VLN_LastA70, 4] = 3  && *
					PLN_Valor = PLN_Valor * VLN_Aux

				case this.VOA_LastA70[VLN_LastA70, 4] = 4  && /
					if VLN_Aux <> 0
						PLN_Valor = PLN_Valor / VLN_Aux
					else
						PLN_Valor = 0
					endif
			endcase
		else
			if PLL_Warning
				this.FON_Msg(1244) && Conversão não cadastrada!
			endif
			this.VOL_ExchangeError = .T.  && Guarda o Valor do Msg - Para ser tratado por fora.
			PLN_Valor = 0
		endif

		this.FOL_CloseTable("A70T")
	endif

	PLN_Round = iif((VLN_Parameters > 4), PLN_Round, .F.)

	this.FOL_PopSelect()
	
	if PLL_InfoA37
		return this.FON_Round(PLN_Valor, PLN_Round)
	else
		return VLN_LastA70
	endif
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Moeda   : Moeda original  
*/                    PLC_What    : "1" - sigla, "2" - descrição  
*/ Descrição        : Retorna o nome da moeda  
*/ Retorno          : Nome da moeda  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_NomeMoeda
	LParameters PLC_Moeda, PLC_What

	LOCAL VLC_Ret, VLD_Date, VLC_Alias

	if empty(PLC_Moeda)
		return ""
	endif

	VLC_Ret   = ""
	VLD_Date  = dtot(this.FOD_StoD(substr(PLC_Moeda, 6)))

	if !(empty(VLD_Date) or isnull(VLD_Date))

		with this
			.FOL_SetParameter(1, VLD_Date)
			.FOL_SetParameter(2, VLD_Date)
			.FOL_SetParameter(3, substr(PLC_Moeda, 1, 5))
			.FOL_ExecuteCursor("CURRENCY_SQL", "CURRENCY1", 3)
		endwith


		if reccount("CURRENCY1") > 0
			do case
		*---	Sigla
				case empty(PLC_What) or PLC_What = "1"
					VLC_Ret = CURRENCY1.a54_001_c
		*---	Descrição
				case PLC_What = "2"
					VLC_Ret = CURRENCY1.a54_002_c
			endcase
		endif
	endif

	return VLC_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_DatI              : Data Inicial  
*/                    PLD_Datf              : Data Final  
*/                    PLL_NaoSomaMesInicial : Não soma o mês inicial  
*/ Descrição        : Calcula a diferença em meses entre as duas datas especificadas  
*/ Retorno          : Retorna a diferenca entre as datas  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_DifMeses
	LParameters PLD_datI, PLD_datF, PLL_NaoSomaMesInicial
	
	LOCAL VLN_nMes
	
	PLD_datI = this.FOD_FirstDayDate( PLD_datI )
	PLD_datF = this.FOD_FirstDayDate( PLD_datF )
	
	VLN_nMES = 0
	
	do while PLD_datI <= PLD_datF and !empty(PLD_datI)
		VLN_nMES = VLN_nMES + 1
	    PLD_datI = Gomonth( PLD_datI , 1 )
	Enddo
	
	return VLN_nMES + iiF(PLL_NaoSomaMesInicial,-1,0)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Nasc           : Data de nascimento 
*/                    PLD_Hoje           : Data Para comparação 
*/                    PLL_SoNoProximoMes : Fica mais velho somente um mes apos a data de  
*/                                         aniversario  
*/ Descrição        : Verifica a idade de uma pessoa  
*/ Retorno          : Retorna a diferenca entre as datas em anos  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Idade
	LParameters PLD_Nasc, PLD_hoje, PLL_soNOproximoMES

	LOCAL VLN_Idade
	
	if empty(PLD_hoje)
	   PLD_hoje = date()
	endif
	
	VLN_Idade = year(PLD_hoje) - year(PLD_Nasc)
	
	if !((month(PLD_hoje) > month(PLD_Nasc)) or ;
	     (month(PLD_hoje) = month(PLD_Nasc)  and;
	      day(PLD_hoje) >= day(PLD_Nasc)))
	
	      VLN_Idade = VLN_idade - 1
	
	else
	    if PLL_soNOproximoMES
	       if month(PLD_Nasc) = month(PLD_hoje) and day(PLD_hoje) >= day(PLD_Nasc)
	          VLN_idade = VLN_idade - 1
	       endif
	    endif
	endif
	
	return VLN_idade
ENDPROC


*/-----------------------------------------------------------------------------------------
*/ Descrição        : Inicializa as pictures e os formats definidas no arquivo Y16 e 	 
*/                    Pictsemp_sys e armazena na matriz VOA_PICTS  
*/ Retorno          : Sem retorno válido  
*/-----------------------------------------------------------------------------------------
PROCEDURE FOU_MakePicts
	LOCAL VLN_MaxArray

	if type("this.VOC_CompanyCode") = "U"
		this.VOC_CompanyCode = space(5)
	endif

	this.FOL_PushSelect()

	
	this.FOL_UseInPackage("y39", "y39")
	select y39
	set order to ukey
	go bottom in y39

	if !eof("Y39")
		VLN_MaxArray = val(y39.ukey)
		dimension this.VOA_Picts[VLN_MaxArray, 3]
		this.VOA_Picts = ""
	else
		dimension this.VOA_Picts[1, 3]
		this.VOA_Picts = ""
		return
	endif

	this.FOL_ExecuteCursor("Y40_ALL", "Y40", 0)
	select Y40
	index on y39_ukey tag y39_ukey

	select Y39
	scan
		VLN_MaxArray = val(y39.ukey)
		with this
			if seek(y39.ukey, "y40", "y39_ukey")
				VLC_FormatCgf 	= y40.y40_001_c
				VLC_IMaskCgf 	= y40.y40_002_c
				VLC_PictureCgf 	= y40.y40_003_c
			else
				VLC_FormatCgf 	= y39.y39_002_c
				VLC_IMaskCgf 	= y39.y39_003_c
				VLC_PictureCgf 	= y39.y39_005_c
			endif
			.VOA_Picts[VLN_MaxArray, 1] = alltrim(VLC_FormatCgf)
			.VOA_Picts[VLN_MaxArray, 2] = alltrim(VLC_IMaskCgf)
			.VOA_Picts[VLN_MaxArray, 3] = alltrim(VLC_PictureCgf)
		endwith
	endscan

	use in Y39
	use in Y40

	this.FOL_PopSelect()

	return
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_Name     : Nome do diretório a ser criado 
*/                    PLL_FileName : Indica se PLC_Name é o nome do arquivo ou diretório  
*/ Descrição        : Cria um diretório com o nome PLC_Name, se PLL_FileName é .F., 
*/                    caso contrário retira o nome do arquivo antes.  
*/ Retorno          : .T. : Se criou com sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_MakeDir
	LParameters PLC_Name, PLL_FileName

	LOCAL VLN_i, VLA_Aux

	*- Se for nome de arquivo retira o nome antes
	if PLL_FileName
		VLN_i = rat("\", PLC_Name)
		if VLN_i > 0
			PLC_Name = substr(PLC_Name, 1, VLN_i - 1)
		else
			PLC_Name = ""
		endif
	endif

	if !empty(PLC_Name)
	*--	Se não existe ainda o diretório cria
		if !directory(PLC_Name)
			mkdir (PLC_Name)
		endif
	endif

	return .T.
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Code : Codigo do Parametro  
*/ Descrição   		: Analisa os parametros do Sistema  
*/ Retorno     		: Resposta do Parâmetro 
*/----------------------------------------------------------------------------------------
PROCEDURE FOU_ChkSets
	LParameters PLC_Code
	LOCAL VLU_Return, VLN_Par
	VLN_Par = ascan(this.VOA_Sets, PLC_Code, -1, -1, 1)
	if VLN_Par > 0
		VLN_Par = asubscript(this.VOA_Sets, VLN_Par, 1)
		VLU_Return = this.VOA_Sets[VLN_Par, 2]
	endif	
	return VLU_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLO_Object : Objeto  
*/ Descrição   		: Traz o objeto mais interno dentro de PLO_Object 
*/ Retorno          : Objeto mais interno  
*/----------------------------------------------------------------------------------------
PROCEDURE FOO_GetActiveControl
	LParameters PLO_Object

	LOCAL VLO_Return, VLO_Object, VLN_Counter

	if 'grid' == lower(PLO_Object.baseclass)
		if PLO_Object.activecolumn > 0
			VLO_Object = PLO_Object.columns[PLO_Object.activecolumn]

			if !empty(VLO_Object.currentcontrol)
				for VLN_Counter = 1 to VLO_Object.controlcount
					if VLO_Object.currentcontrol = VLO_Object.controls[VLN_Counter].name
						VLO_Return = VLO_Object.controls[VLN_Counter]
						exit
					endif
				endfor
			else
				this.FON_Msg(264,program()) && "Erro fatal em " + program()
			endif
		endif
	else
		VLO_Return = PLO_Object
	endif

	return VLO_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Expression : indica a expressão em que se quer colocar o alias  
*/                    PLC_Alias1 : indica o alias 1 
*/ Descrição   		: Troca o alias na expressão com strtran  
*/ Retorno          : Expressão resultado  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_PutExprAlias
	LParameters PLC_Expression, PLC_Alias1

	*- verifica se existe %alias1 na expressão se sim substitui pelo alias, se não adiciona
	*-   o alias no começo
	PLC_Expression = lower(PLC_Expression)
	if at("%alias1", PLC_Expression) > 0
		PLC_Expression = strtran(PLC_Expression, '%alias1', PLC_Alias1)
	else
		if at(".", PLC_Expression) = 0 and !empty(PLC_Alias1)
			PLC_Expression = PLC_Alias1 + "." + PLC_Expression
		endif
	endif

	return PLC_Expression
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLO_oSource : objeto que deve ser atribuído o valor 
*/                    PLU_Value : valor atribuído ao objeto 
*/ Descrição   		: Atribui o valor ao objeto especificado  
*/ Retorno          : Expressão resultado  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_SetValue
	LParameters PLO_oSource, PLU_Value

	PLO_oSource.value = PLU_Value

	return .t.
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLD_Date - Data para escrever por extenso 
*/ Descrição   		: Retorna a data escrita por extenso  
*/ Retorno     		: Data por extenso 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_WriteLongDate
	LParameters PLD_Date

	LOCAL VLC_Return, VLC_DiaSemana, VLC_Mes, VLC_Dia, VLC_Ano

	VLC_DiaSemana = alltrim(this.VOA_021[dow(PLD_Date)])
	VLC_Mes = alltrim(this.VOA_020[month(PLD_Date)])
	VLC_Dia = alltrim(trans(day(PLD_Date), "@L 99"))
	VLC_Ano = alltrim(trans(year(PLD_Date), "@9 9999"))

	VLC_Return = VLC_DiaSemana+", "+VLC_Dia+" de "+VLC_Mes+" de "+VLC_Ano+"."

	return VLC_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Expression - Expressão que deve ser avaliada  
*/ Descrição   		: Retorna o check sum de uma expressão caractér 
*/ Retorno     		: check sum  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_CheckSum
	LParameters PLC_Expression

	LOCAL VLC_Return

	VLC_Return = sys(2007, PLC_Expression)

	return VLC_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Expression - Expressão que deve ser avaliada  
*/					  PLC_ValidCheckSum - Check sum já gravado  
*/ Descrição   		: Verifica se o check sum de uma expressão foi alterado 
*/ Retorno     		: .T. se estiver ok, .F. se houve modificação 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_VerifyCheckSum
	LParameters PLC_Expression, PLC_ValidCheckSum

	LOCAL VLL_Return

	VLL_Return = alltrim(sys(2007, PLC_Expression)) == alltrim(PLC_ValidCheckSum)

	return VLL_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_FileType : Tipo de arquivo (pode ser caracter indicando os  
*/                                   tipos de arquivos) 
*/                    PLL_CompletePath : indica se retornará o path completo  
*/ Descrição        : Retorna um nome de arquivo 
*/ Retorno          : O nome do arquivo  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_GetFile
	LParameters PLN_FileType, PLL_CompletePath

	LOCAL  VLC_File, VLC_Type

	if empty(PLN_FileType)
		PLN_FileType = FILE_ALL
	endif

	if type("PLN_FileType") = "N"
		do case
			case PLN_FileType = FILE_ALL
					VLC_TYPE = ""
			case PLN_FileType = FILE_PICTURE
					VLC_TYPE = "bmp|pcx"
			case PLN_FileType = FILE_PROGRAM
					VLC_TYPE = "fxp|prg|qpr"
			case PLN_FileType = FILE_SOUND
					VLC_TYPE = "wav"
			case PLN_FileType = FILE_FORM
					VLC_TYPE = "scx"
		endcase
	else
		VLC_Type = PLN_FileType
	endif

	if (type("PLN_FileType") = "C") or PLN_FileType != FILE_PICTURE
		VLC_File = upper(alltrim(getfile(VLC_TYPE)))
	else
		VLC_File = upper(alltrim(getpict()))		&& Trato Imagem como Imagem
	endif

	*- Só deixa selecionar arquivos do diretório corrente, caso contrário retorna vazio
	if !empty(VLC_File)
		if empty(PLL_CompletePath)
			if !(VLC_File = this.VOC_DirHome)
				this.FON_Msg(82)
				VLC_File = ""
			else
				VLC_File = strtran(VLC_File, this.VOC_DirHome, "")
			endif
		endif
	endif

	return VLC_File
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLC_File : nome do arquivo 
*/                    PLN_FileType : Tipo de arquivo (pode ser caracter indicando os  
*/                                   tipos de arquivos) 
*/                    PLL_CompletePath : indica se retornará o path completo  
*/ Descrição        : Grava um  arquivo com um nome especificado  
*/ Retorno          : o nome do arquivo  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_PutFile
	LParameters PLC_File, PLN_FileType, PLL_CompletePath

	LOCAL  VLC_File, VLC_Type

	if empty(PLN_FileType)
		PLN_FileType = FILE_ALL
	endif

	if type("PLN_FileType") = "N"
		do case
			case PLN_FileType = FILE_ALL
					VLC_TYPE = ""
			case PLN_FileType = FILE_PICTURE
					VLC_TYPE = "bmp|pcx"
			case PLN_FileType = FILE_PROGRAM
					VLC_TYPE = "fxp|prg|qpr"
			case PLN_FileType = FILE_SOUND
					VLC_TYPE = "wav"
			case PLN_FileType = FILE_FORM
					VLC_TYPE = "scx"
		endcase
	else
		VLC_Type = PLN_FileType
	endif

	VLC_File = upper(alltrim(putfile("", PLC_File, VLC_TYPE)))

	*- Só deixa selecionar arquivos do diretório corrente, caso contrário retorna vazio
	if empty(PLL_CompletePath)
		if !(VLC_File = this.VOC_DirHome)
			this.FON_Msg(82)
			VLC_File = ""
		else
			VLC_File = strtran(VLC_File, this.VOC_DirHome, "")
		endif
	endif

	return VLC_File
ENDPROC


*/--------------------------------------------------------------------------------------------
*/ Parametros		: PLC_ExpressionList : Memo que contém as macros(uma por linha) 		  
*/						a serem executadas       											  
*/ Descrição		: Executa as macros declaradas em cada linha de PLC_ExpressionList  
*/ Retorno			: .T. : Se executou com sucesso, .F. : c.c. 
*/ --- Modificada por Denis para executar a função execscript do VFP ---
*/--------------------------------------------------------------------------------------------
PROCEDURE FOL_DoMacros
	LParameters PLC_ExpressionList
	*- Apesar de ser um FOL_... poderá ter outros retornos, dependendo apenas do parâmetro

	local VLU_Return
	
	VLU_Return = .t.
	if !empty(nvl(PLC_ExpressionList,""))
		VLU_Return = execscript(PLC_ExpressionList)
	endif

*-		LOCAL  VLN_i, VLC_Macro, VLN_MaxLine, VLL_AppeLine, VLN_OldSetMemo, VLN_MLine

*-		VLN_OldSetMemo = set("MEMO")
*-		set memo to 1024

*-		VLN_MaxLine = memlines( PLC_ExpressionList )
*-		VLN_Mline = 0

*-		VLL_AppeLine = .F.

*-		*- Varre o campo memo e executa as macros
*-		for VLN_i = 1 to VLN_MaxLine
*-			_MLINE = VLN_Mline
*-			if VLL_AppeLine
*-				VLC_Macro = VLC_Macro + alltrim(mline( PLC_ExpressionList, 1, _MLINE))
*-			else
*-				VLC_Macro = alltrim(mline( PLC_ExpressionList, 1, _MLINE))
*-			endif
*-			VLN_Mline = _MLINE

*-			if substr(VLC_Macro, len(VLC_Macro), 1) = ";"
*-				VLC_Macro = trim(substr(VLC_Macro, 1, len(VLC_Macro)-1)) + " "
*-				VLL_AppeLine = .T.
*-			else
*-				if !empty(VLC_Macro) and VLC_Macro <> "*"
*-					&VLC_Macro
*-				endif
*-				VLL_AppeLine = .F.
*-			endif
*-		endfor

*-		set memo to VLN_OldSetMemo

	return VLU_Return
ENDPROC


*/--------------------------------------------------------------------------------------------
*/ Parametros		: PLM_PropertiesInformation : Memo com as informações de propriedades 
*/					  PLC_PropertieToSee		: Propriedade a ser consultada  
*/					  PLN_Position 				: Número da propriedade na linha ; 		  
*/														    (separada por vírgulas) 
*/					  PLL_NoWarning				: Não avisa se não existir a propriedade  
*/					  PLN_DefaultToReturn		: Se não existir retorna este Default 
*/					  PLL_NoExpression 			: Retorna a String sem converter para  	  
*/														qualquer expressão			   		  
*/ Descrição		: Procura uma propriedade no Memo PLM_PropertiesInformation 
*/ Retorno			: Valor da propriedade (pode ser de qualquer tipo)  
*/--------------------------------------------------------------------------------------------
PROCEDURE FOU_ChkProperties
	LParameters PLM_PropertiesInformation, PLC_PropertieToSee, PLN_Position, PLL_NoWarning,;
				 PLN_DefaultToReturn, PLL_NoExpression

	LOCAL VLC_LineFound, VLU_Return, VLN_Pos1, VLN_Pos2

	*- By ghviotto
	VLN_OldSetMemo = set("MEMO")
	set memo to 1024

	*- Obtém a linha que possui a propriedade PLC_PropertieToSee
	VLC_LineFound = mline(PLM_PropertiesInformation, atcline( PLC_PropertieToSee, PLM_PropertiesInformation ))+" "

	*- Procura a posição dentro da linha, se não encontra retorna .F.
	if !empty(VLC_LineFound)
		VLU_Return    = alltrim( substr( VLC_LineFound, at( "=", VLC_LineFound)+1) )
		if !empty(PLN_Position) and PLN_Position > 0
			if PLN_Position > 1
				VLN_Pos = at(";", VLU_Return, PLN_Position - 1)
				if VLN_Pos > 0
					VLU_Return = alltrim(substr(VLU_Return, VLN_Pos + 1))
				else
					VLU_Return = ""
				endif
			endif
			VLN_Pos = at(";", VLU_Return)
			if VLN_Pos > 0
				VLU_Return = alltrim(substr(VLU_Return, 1, VLN_Pos - 1))
			endif
		endif

		do case
	*---	Retorna sempre a String
			case PLL_NoExpression = .T.
				VLU_Return =  VLU_Return
	*---	Character
			case VLU_Return = '"'			
				VLU_Return = substr( VLU_Return, 2, len(VLU_Return) - 2)
	*---	Numbers
			case isdigit( VLU_Return )
				VLU_Return = val( strtran(strtran(VLU_Return, ",", ""), ".", set("point")) )
	*---	Dates
			case VLU_Return = "{"
				VLU_Return = evaluate( VLU_Return )
	*---	Logics
			case VLU_Return = "."			
				VLU_Return = evaluate( VLU_Return )
	*---	Macros&Expressions
			case VLU_Return = "&"			
				VLU_Return = evaluate( subs(VLU_Return,2) )
		endcase
	else
		if !PLL_NoWarning
			this.FON_Msg(275) && "Propertie not Found !"
		endif
	*-	Retorno padrão
		VLU_Return = PLN_DefaultToReturn
	endif
	set memo to VLN_OldSetMemo

	return VLU_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLU_MessageId : Identificador da mensagem 
*/                    PLC_AuxInfo1  : Mensagem opcional	  
*/                    PLC_AuxInfo2  : Mensagem opcional	2 
*/                    PLL_Wait      : Quando for Wait ou Nowait 
*/                    PLL_soMessage : Quando for Wait e o retorno e a mensagem  
*/ Descrição        : Executa um dialog box de acordo com o identificador  PLN_MessageId  
*/ Retorno          : 1 : OK, 2 : Cancel, 3 : Abort, 4 : Retry, 5 : Ignore, 
*/                    6 : Yes, 7 : No, -1 : Error 
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Msg
	LParameters PLU_MessageId, PLC_AuxInfo1, PLC_AuxInfo2, PLL_Wait, PLL_soMessage

	LOCAL VLN_DialogBoxButtons, VLN_Icon, VLN_DefaultButton, VLC_WindowTiTle, VLC_Message, ;
		  VLN_Return, VLN_WaitingTime, VLC_KeyboardDefault, VLU_Ret, VLC_MessageNumber

	if vartype(PLU_MessageId) = "N"
		VLC_MessageCode = "msg_" + trans(PLU_MessageId, "@L 99999")
		VLC_MessageNumber = " (" + trans(PLU_MessageId, "@L 99999") + ")"
	else
		VLC_MessageNumber = ""
		VLC_MessageCode = PLU_MessageId
	endif

	this.vol_lastSeek = _FALSE_		&& FOC_Caption() altera este Valor
	VLC_Message 	  = this.FOC_Caption(VLC_MessageCode)

	if this.vol_lastSeek
		VLC_Properties = this.VOC_LastCaptionProperties
	else
		VLC_Message    = ""
		VLC_Properties = ""
	endif
			
	with this
		VLC_Type 			 = .FOU_ChkProperties( VLC_Properties, "WarningType", 1, .T., ""  )
		VLN_DialogBoxButtons = .FOU_ChkProperties( VLC_Properties, "DialogBoxButtons", 1, .T. )
		VLN_Icon 			 = .FOU_ChkProperties( VLC_Properties, "Icon", 1, .T. )
		VLN_DefaultButton 	 = .FOU_ChkProperties( VLC_Properties, "DefaultButton", 1, .T. )
		VLC_WindowTiTle 	 = .FOU_ChkProperties( VLC_Properties, "WindowTiTle", 1, .T. )
		VLN_WaitingTime  	 = .FOU_ChkProperties( VLC_Properties, "WaitingTime", 1, .T. )
		VLC_KeyboardDefault  = .FOU_ChkProperties( VLC_Properties, "KeyboardDefault", 1, .T. )
	endwith

	if empty(PLC_AuxInfo1)
		PLC_AuxInfo1 = ""
	endif

	if empty(PLC_AuxInfo2)
		PLC_AuxInfo2 = ""
	endif

	VLC_Message = strtran(VLC_Message, "@1", PLC_AuxInfo1)
	VLC_Message = strtran(VLC_Message, "@2", PLC_AuxInfo2)
	VLC_Message = strtran(VLC_Message, "/n", chr(13))
	VLC_Message = strtran(VLC_Message, "/N", chr(13))
	VLC_Message = strtran(VLC_Message, "/d", dtoc(date()))
	VLC_Message = strtran(VLC_Message, "/D", dtoc(date()))
	VLC_Message = strtran(VLC_Message, "/h", time())
	VLC_Message = strtran(VLC_Message, "/H", time())
	VLC_Message = strtran(VLC_Message, "/t", chr(9))
	VLC_Message = strtran(VLC_Message, "/T", chr(9))

	if VLC_Type != "WAIT"

*--		Trata as posibilidade dos tipos de botões
		if empty(VLN_DialogBoxButtons)
			VLN_DialogBoxButtons = ""
		endif
		VLN_DialogBoxButtons = upper( VLN_DialogBoxButtons )

		do case
			case VLN_DialogBoxButtons  == "OK"
				VLN_DialogBoxButtons = 0
			case VLN_DialogBoxButtons == "OKCANCEL"
				VLN_DialogBoxButtons = 1
			case VLN_DialogBoxButtons == "ABORTRETRYIGNORE"
				VLN_DialogBoxButtons = 2
			case VLN_DialogBoxButtons == "YESNOCANCEL"
				VLN_DialogBoxButtons = 3
			case VLN_DialogBoxButtons == "YESNO"
				VLN_DialogBoxButtons = 4
			case VLN_DialogBoxButtons == "RETRYCANCEL"
				VLN_DialogBoxButtons = 5
			other
				VLN_DialogBoxButtons = 0
		endcase

*--		Trata as posibilidade dos ícones
		if empty(VLN_Icon)
			VLN_Icon = ""
		endif
		VLN_Icon = upper( VLN_Icon )
		do case
			case VLN_Icon = "STOP"
				VLN_Icon = 16
			case VLN_Icon = "?"
				VLN_Icon = 32
			case VLN_Icon = "!"
				VLN_Icon = 48
			case VLN_Icon = "INFORMATION"
				VLN_Icon = 64
			other
				VLN_Icon = 0
		endcase

*--		Verifica qual o botão inicial
		if empty(VLN_DefaultButton)
			VLN_DefaultButton = 0
		endif
		do case
			case VLN_DefaultButton = 1
				VLN_DefaultButton = 0
			case VLN_DefaultButton = 2
				VLN_DefaultButton = 256
			case VLN_DefaultButton = 3
				VLN_DefaultButton = 512
			other
				VLN_DefaultButton = 0
		endcase

*--		Verifica o texto e o título do aviso
		if empty( VLC_WindowTiTle )
			VLC_WindowTiTle = ""
		endif
		if empty( VLC_Message )
			VLC_Message = ""
		endif

	endif

	VLU_Ret = -1

	this.VOC_LastMessage = VLC_Message

	if !this.VOL_ServiceMode
		*- Executa como retorno o da função MESSAGEBOX
		if VLC_Type != "WAIT"
			VLU_Ret = messagebox( VLC_Message, VLN_DialogBoxButtons + VLN_Icon + VLN_DefaultButton,  this.FOC_Caption(VLC_WindowTiTle) + VLC_MessageNumber )
		else
			if PLL_soMEssage
				return (VLC_Message)
			else
				if !PLL_Wait
					wait window (VLC_Message) nowait
				else
					wait window (VLC_Message)
				endif
			endif
		ENDIF
	ENDIF
	
	return VLU_Ret
ENDPROC


*/---------------------------------------------------------------------------------------------
*/ Parametros  :      PLC_Alias     : Arquivo onde a sequência será pesquisada 
*/				      PLC_Tag       : Tag que deve conter a Sequência para busca 
*/				      PLC_Filter    : Expressão que será procurada na PLC_Tag  
*/				      PLC_Field     : Campo onde está armazenado a Ordem 
*/				      PLN_Skip      : Ajuste do pulo da ordem 10 em 10 - 10, 5 em 5 - 5  
*/ Descrição        : Retorna a ultima sequência utilizada mais o pulo 
*/ Retorno          : Próxima sequência da numeração com pict @L 99999 
*/ Ultima alteração : 21/12/1998  
*/---------------------------------------------------------------------------------------------
PROCEDURE FOC_NextSeqNumber
	LParameters PLC_Alias, PLC_Tag, PLC_Filter, PLC_Field, PLN_Skip, VLC_Tag

	LOCAL VLC_Retorno, VLC_Alias

	VLC_Alias = select()
	VLC_Tag = tag()
	VLC_NewAlias = "RDO_"+PLC_Alias

	use dbf(PLC_Alias) again alias (VLC_NewAlias) in 0
	select (VLC_NewAlias)
	set order to PLC_Tag descending
	if seek(PLC_Filter, VLC_NewAlias)
		VLC_Retorno = transform(val(evaluate("RDO_"+PLC_Field)) + PLN_Skip, [@L 99999])
	else
		VLC_Retorno = transform(PLN_Skip, [@L 99999])
	endif

	use in (VLC_NewAlias)
	select (VLC_Alias)
	set order to (VLC_Tag)

	return VLC_Retorno
ENDPROC


*/---------------------------------------------------------------------------------------------
*/ Parametros       : PLN_OperatorStr : Número do comparador na matriz VGA_232 
*/ Descrição        : Analisa o conteúdo da string e retorna o elemento correspondente na  
*/					  array VGA_232 
*/---------------------------------------------------------------------------------------------
PROCEDURE FOC_Operator
	LPARAMETERS PLN_OperatorStr

	LOCAL VLC_Return

	*- Matriz VGA_232
	*-	01-Igual
	*-	02-Maior
	*-	03-Menor
	*-	04-Maior ou Igual
	*-	05-Menor ou Igual
	*-	06-Iniciado por
	*-	07-Terminado por
	*-	08-Contém
	*-	09-Está contido
	*-	10-Entre

	do case
		case PLN_OperatorStr = 1
			VLC_Return = "="
		case PLN_OperatorStr = 2
			VLC_Return = ">"
		case PLN_OperatorStr = 3
			VLC_Return = "<"
		case PLN_OperatorStr = 4
			VLC_Return = ">="
		case PLN_OperatorStr = 5
			VLC_Return = "<="
		case PLN_OperatorStr = 6
			VLC_Return = "%LIKE"
		case PLN_OperatorStr = 7
			VLC_Return = "LIKE%"
		case PLN_OperatorStr = 8
			VLC_Return = "%LIKE%"
		case PLN_OperatorStr = 9
			VLC_Return = "IN"
		case PLN_OperatorStr = 10
			VLC_Return = "BETWEEN"
	endcase

	return VLC_Return
ENDPROC


*/---------------------------------------------------------------------------------------------
*/ Parametros       : PLC_OperatorStr : String a ser analisada 
*/ Descrição        : Analisa o conteúdo da string e retorna o elemento correspondente na  
*/					  array VGA_232 
*/---------------------------------------------------------------------------------------------
PROCEDURE FON_Operator
	LPARAMETERS PLC_OperatorStr

	LOCAL VLN_Return

	PLC_OperatorStr = alltrim(upper(PLC_OperatorStr))

	*- Matriz VGA_232
	*-	01-Igual
	*-	02-Maior
	*-	03-Menor
	*-	04-Maior ou Igual
	*-	05-Menor ou Igual
	*-	06-Iniciado por
	*-	07-Terminado por
	*-	08-Contém
	*-	09-Está contido
	*-	10-Entre


	do case
		case PLC_OperatorStr == "="
			VLN_Return = 1
		case PLC_OperatorStr == ">"
			VLN_Return = 2
		case PLC_OperatorStr == "<"
			VLN_Return = 3
		case PLC_OperatorStr == ">="
			VLN_Return = 4
		case PLC_OperatorStr == "<="
			VLN_Return = 5
		case PLC_OperatorStr == "%LIKE"
			VLN_Return = 6
		case PLC_OperatorStr == "LIKE%"
			VLN_Return = 7
		case PLC_OperatorStr == "%LIKE%"
			VLN_Return = 8
		case PLC_OperatorStr == "IN"
			VLN_Return = 9
		case PLC_OperatorStr == "BETWEEN"
			VLN_Return = 10
	endcase

	return VLN_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  : PLN_Number : Numero para ser escrito 
*/ Descrição   : Escreve um Numero por extenso 
*/ Retorno     : O extenso Formatado conforme parametrizacao  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_Extenso
	LPARAMETERS PLN_Number, PLN_Len, PLL_OnlyInt, PLC_Unit, PLC_CurrencySingular, PLC_CurrencyPlural,;
	PLC_DecimalSingular, PLC_DecimalPlural, PLL_CurrencyHaveCents

	#DEFINE STARCHAR 		this.VOA_007_SIGNS[1]
	#DEFINE ENDCHAR 		this.VOA_007_SIGNS[2]
	#DEFINE FILLCHAR 		this.VOA_007_SIGNS[3]
	#DEFINE JUNCTIONSTRING 	this.VOA_007_SIGNS[4]
	#DEFINE JUNCTIONSIGN 	this.VOA_007_SIGNS[5]
	#DEFINE JUNCTIONNAME 	this.VOA_007_SIGNS[6]
	#DEFINE JUNCTIONSPACE 	this.VOA_007_SIGNS[7]
	#DEFINE UPPERCASEMODE 	this.VOA_007_SIGNS[8]
	#DEFINE _100_to_199 	this.VOA_007_SIGNS[9]
	#DEFINE THOUSANDSing 	this.VOA_007_SIGNS[10]
	#DEFINE THOUSANDPlur 	this.VOA_007_SIGNS[11]
	#DEFINE MILLIONSing 	this.VOA_007_SIGNS[12]
	#DEFINE MILLIONPlur 	this.VOA_007_SIGNS[13]
	#DEFINE BILLIONSing 	this.VOA_007_SIGNS[14]
	#DEFINE BILLIONPlur 	this.VOA_007_SIGNS[15]
	#DEFINE CURRENCY_Sing		PLC_CurrencySingular
	#DEFINE CURRENCY_Plur		PLC_CurrencyPlural
	#DEFINE CURRENCY_Dec_Sing	PLC_DecimalSingular
	#DEFINE CURRENCY_Dec_Plur	PLC_DecimalPlural
	#DEFINE CURRENCY_Cents   	PLL_CurrencyHaveCents
	#DEFINE CURRENCYARRAYNUMBER	7

	LOCAL VLC_Extenso, VLL_UnitDate, VLC_InitStringUnit, VLD_DateUnit
	LOCAL _PARTES, _VALOR, i, auxNUM, PLN_Number, _RT, _cercarINI, _cercarFIN, _completarCOM, VLN_Len



	VLN_Len = iif(empty(PLN_Len), 200, PLN_Len)
	VLD_DateUnit = date()
	VLC_InitStringUnit = ""
	VLL_UnitDate = .F.

	
	*- Abre o arquivo que as definicoes dos arrays
	this.FOL_PushSelect()

	dimension this.VOA_007_NUMBERS[36], this.VOA_007_SIGNS[15]

	with this
		.VOA_007_NUMBERS[01] = .FOC_Caption("um")
		.VOA_007_NUMBERS[02] = .FOC_Caption("dois")
		.VOA_007_NUMBERS[03] = .FOC_Caption("tres")
		.VOA_007_NUMBERS[04] = .FOC_Caption("quatro")
		.VOA_007_NUMBERS[05] = .FOC_Caption("cinco")
		.VOA_007_NUMBERS[06] = .FOC_Caption("seis")
		.VOA_007_NUMBERS[07] = .FOC_Caption("sete")
		.VOA_007_NUMBERS[08] = .FOC_Caption("oito")
		.VOA_007_NUMBERS[09] = .FOC_Caption("nove")
		.VOA_007_NUMBERS[10] = .FOC_Caption("dez")
		.VOA_007_NUMBERS[11] = .FOC_Caption("onze")
		.VOA_007_NUMBERS[12] = .FOC_Caption("doze")
		.VOA_007_NUMBERS[13] = .FOC_Caption("treze")
		.VOA_007_NUMBERS[14] = .FOC_Caption("quatorze")
		.VOA_007_NUMBERS[15] = .FOC_Caption("quinze")
		.VOA_007_NUMBERS[16] = .FOC_Caption("dezesseis")
		.VOA_007_NUMBERS[17] = .FOC_Caption("dezessete")
		.VOA_007_NUMBERS[18] = .FOC_Caption("dezoito")
		.VOA_007_NUMBERS[19] = .FOC_Caption("dezenove")
		.VOA_007_NUMBERS[20] = .FOC_Caption("vinte")
		.VOA_007_NUMBERS[21] = .FOC_Caption("trinta")
		.VOA_007_NUMBERS[22] = .FOC_Caption("quarenta")
		.VOA_007_NUMBERS[23] = .FOC_Caption("cinquenta")
		.VOA_007_NUMBERS[24] = .FOC_Caption("sessenta")
		.VOA_007_NUMBERS[25] = .FOC_Caption("setenta")
		.VOA_007_NUMBERS[26] = .FOC_Caption("oitenta")
		.VOA_007_NUMBERS[27] = .FOC_Caption("noventa")
		.VOA_007_NUMBERS[28] = .FOC_Caption("cem")
		.VOA_007_NUMBERS[29] = .FOC_Caption("duzentos")
		.VOA_007_NUMBERS[30] = .FOC_Caption("trezentos")
		.VOA_007_NUMBERS[31] = .FOC_Caption("quatrocentos")
		.VOA_007_NUMBERS[32] = .FOC_Caption("quinhentos")
		.VOA_007_NUMBERS[33] = .FOC_Caption("seiscentos")
		.VOA_007_NUMBERS[34] = .FOC_Caption("setecentos")
		.VOA_007_NUMBERS[35] = .FOC_Caption("oitocentos")
		.VOA_007_NUMBERS[36] = .FOC_Caption("novecentos")
	endwith

	*	Propriedade para formacao do Extenso

	this.VOA_007_SIGNS[ 1] = "[ "
	this.VOA_007_SIGNS[ 2] = " ] "
	this.VOA_007_SIGNS[ 3] = "/"
	this.VOA_007_SIGNS[ 4] = this.FOC_Caption("e")
	this.VOA_007_SIGNS[ 5] = ","
	this.VOA_007_SIGNS[ 6] = this.FOC_Caption("de")
	this.VOA_007_SIGNS[ 7] = " "
	this.VOA_007_SIGNS[ 8] = ""
	this.VOA_007_SIGNS[ 9] = this.FOC_Caption("cento")
	this.VOA_007_SIGNS[10] = this.FOC_Caption("mil")
	this.VOA_007_SIGNS[11] = this.FOC_Caption("mil")
	this.VOA_007_SIGNS[12] = this.FOC_Caption("milhao")
	this.VOA_007_SIGNS[13] = this.FOC_Caption("milhoes")
	this.VOA_007_SIGNS[14] = this.FOC_Caption("bilhao")
	this.VOA_007_SIGNS[15] = this.FOC_Caption("bilhoes")

	*	Definicao do Extenso básico das Unidades
	if empty(PLC_Unit)
		if empty(PLC_CurrencySingular)
			PLC_CurrencySingular	= this.FOC_Caption("real")
			PLC_CurrencyPlural      = this.FOC_Caption("reais")
			PLC_DecimalSingular     = this.FOC_Caption("centavo")
			PLC_DecimalPlural       = this.FOC_Caption("centavos")
			PLL_CurrencyHaveCents   = .t.
		endif
	else
		PLL_CurrencyHaveCents   = .t.
		*- Verifica se é unidade de medida (T02).
		if len(PLC_Unit) = 20 
			if this.FOL_FindExpression("T02_UKEY", "TMP_T02", PLC_Unit)
				PLC_CurrencySingular	= alltrim(tmp_t02.t02_004_c)
				PLC_CurrencyPlural      = alltrim(tmp_t02.t02_005_c)
				PLC_DecimalSingular     = alltrim(tmp_t02.t02_006_c)
				PLC_DecimalPlural       = alltrim(tmp_t02.t02_007_c)
				
				PLL_OnlyInt			    = tmp_t02.t02_011_n = 0
			endif
			this.FOL_CloseTable("TMP_T02")
		else
			*- Tratamento da data antes de passar para o ExecuteCursor
			*- Problema de passar string é porque o Oracle não aceita.
			VLL_UnitDate = !empty(right(padr(PLC_Unit, 13), 8))
			if VLL_UnitDate
				VLC_InitStringUnit = right(padr(PLC_Unit, 13), 8)
				VLD_DateUnit = ctod(substr(VLC_InitStringUnit, 7, 2) + '/' + substr(VLC_InitStringUnit, 5, 2) + '/' + substr(VLC_InitStringUnit, 1, 4))
			endif
			this.FOL_SetParameter(1, left(PLC_Unit, 5))
			this.FOL_SetParameter(2, VLD_DateUnit)
			this.FOL_ExecuteCursor("A54_A36_UKEY", "TMP_A54", 2)

			PLC_CurrencySingular	= alltrim(tmp_a54.a54_008_c)
			PLC_CurrencyPlural      = alltrim(tmp_a54.a54_009_c)
			PLC_DecimalSingular     = alltrim(tmp_a54.a54_011_c)
			PLC_DecimalPlural       = alltrim(tmp_a54.a54_012_c)

			PLL_OnlyInt			    = tmp_a54.a54_010_n = 0

			this.FOL_CloseTable("TMP_A54")
		endif
	endif

	this.FOL_PopSelect()

	VLC_AuxNumber = trans(PLN_Number,'999,999,999.099')
	VLC_Extenso   = space(VLN_Len)

	declare VLA_Parts[4]
	declare VLA_Value[4]
	for i = 1 to 4
		VLA_Parts[i] = this.FOC_NumToString(substr(VLC_AuxNumber,(i * 4) -3,3))
		VLA_Value[i]  = val(substr(VLC_AuxNumber,(i * 4) -3,3))
	next
	if VLA_Value[1] > 0
		if VLA_Value[1] > 1
			VLC_AuxExtenso = JUNCTIONSPACE + MILLIONPlur
		else
			VLC_AuxExtenso = JUNCTIONSPACE+MILLIONSing
		endif

		if VLA_Value[2] = 0 .and. VLA_Value[3] = 0
			VLC_AuxExtenso = VLC_AuxExtenso + JUNCTIONSPACE+JUNCTIONNAME+JUNCTIONSPACE
		else
			if VLA_Value[2] # 0 .or. (val(subs(VLC_AuxNumber,9,1)) > 0 .and.;
					val(subs(VLC_AuxNumber,10,1)) > 0) .or. (val(subs(VLC_AuxNumber,9,1)) = 0 .and.;
					val(subs(VLC_AuxNumber,10,1)) > 1 .and. val(subs(VLC_AuxNumber,11,1)) > 0)
				VLC_AuxExtenso = VLC_AuxExtenso + JUNCTIONSIGN + JUNCTIONSPACE
			else
				if ( VLA_Value[2] = 0 .and. VLA_Value[3]  >0 )
					VLC_AuxExtenso = VLC_AuxExtenso +JUNCTIONSPACE+JUNCTIONSTRING+JUNCTIONSPACE
				endif
			endif
		endif
		VLA_Parts[1] = VLA_Parts[1] + VLC_AuxExtenso
	endif
	if VLA_Value[2] > 0
		store JUNCTIONSPACE + THOUSANDSing TO VLC_AuxExtenso
		
		
		if VLA_Value[3] # 0
			store JUNCTIONSPACE + THOUSANDSing + JUNCTIONSIGN + JUNCTIONSPACE TO VLC_AuxExtenso
		endif
		if VLA_Value[2] > 0 .AND. VLA_Value[3] > 0
			store JUNCTIONSPACE + THOUSANDSing + JUNCTIONSPACE + JUNCTIONSTRING + JUNCTIONSPACE TO VLC_AuxExtenso
		endif
		if val(substr(VLC_AuxNumber,10,2)) # 0 .AND. VLA_Value[3] > 100
			STORE JUNCTIONSPACE + THOUSANDSing + JUNCTIONSIGN + JUNCTIONSPACE TO VLC_AuxExtenso
		endif
		VLA_Parts[2] = VLA_Parts[2] + VLC_AuxExtenso
	endif
	if PLN_Number >= 1
		if int(PLN_Number) > 1
			VLA_Parts[3] = VLA_Parts[3] + JUNCTIONSPACE + trim(CURRENCY_Plur)
		else
			VLA_Parts[3] = VLA_Parts[3] + JUNCTIONSPACE + trim(CURRENCY_Sing)
		endif
	endif
	if VLA_Value[4] > 0 AND !PLL_OnlyInt .and. CURRENCY_Cents
		if PLN_Number > 1
			VLA_Parts[4] = JUNCTIONSPACE + JUNCTIONSTRING + JUNCTIONSPACE + VLA_Parts[4]
		endif
		if VLA_Value[4] > 1
			VLA_Parts[4] = VLA_Parts[4] + JUNCTIONSPACE + CURRENCY_Dec_Plur
		else
			VLA_Parts[4] = VLA_Parts[4] + JUNCTIONSPACE + CURRENCY_Dec_Sing
		endif
	else
		VLA_Parts[4] = ""
	endif
	VLC_Extenso = STARCHAR + VLA_Parts[1] + VLA_Parts[2] + VLA_Parts[3] + VLA_Parts[4]
	VLC_Extenso = strtran(alltrim(substr(VLC_Extenso, 1, VLN_Len)), JUNCTIONSPACE + JUNCTIONSPACE,JUNCTIONSPACE)

	if !empty(PLN_Len)
		VLC_Extenso = this.FOC_STinSize( VLC_Extenso, PLN_Len, 1, .t. ) + ;
					  this.FOC_STinSize( VLC_Extenso, PLN_Len, 2, .t. ) + ;
					  this.FOC_STinSize( VLC_Extenso, PLN_Len, 3, .t. ) + ;
					  this.FOC_STinSize( VLC_Extenso, PLN_Len, 4, .t. ) + ;
					  this.FOC_STinSize( VLC_Extenso, PLN_Len, 5, .t. )
	endif
	VLC_Extenso = padr(alltrim(VLC_Extenso) + ENDCHAR , VLN_Len, FILLCHAR) + JUNCTIONSPACE



	return VLC_Extenso
ENDPROC

*/----------------------------------------------------------------------------------------/*
*/ Function         : FOC_STinSize               /*
*/ Parâmetros       : PLC_String         : Campo Memo de origem                           /*
*/                    PLN_Size           : Tamanho da linha                               /*
*/                    PLN_nRow           : Número da linha desejada                       /*
*/                    PLL_ForceWidth     : Obriga que o retorno tenha o tamanho de        /*
*/                                         PLN_Size                                       /*
*/                    PLL_returnRowCount : Retorna o número de linhas do campo memo       /*
*/ Descrição        : Escreve um pedaço do campo memo                                     /*
*/ Retorno          : Pedaço do campo memo ou tamanho do campo memo em linhas             /*
*/ Grupo            : GENERIC                    /*
*/ Procura por      : Campo Memo                 /*
*/ Última alteração : 24/11/97                   /*
*/ Alterado por     : Marcelo Ris                /*
*/ Versão           : 1                          /*
*/----------------------------------------------------------------------------------------/*
PROCEDURE FOC_StinSize
	LParameters PLC_String, PLN_Size, PLN_nRow, PLL_ForceWidth, PLL_returnRowCount
	 
	LOCAL VLN_OldSetMemo, VLC_AuxReturn, VLN_RowCount
	 
		VLN_OldSetMemo = set("MEMO")
		set memo to max(8, min(1024, PLN_Size))
		 
		VLN_RowCount= memline( PLC_String )
		 
		if PLL_ForceWidth
			VLC_Auxreturn = padr(mline(PLC_String, PLN_nRow), PLN_Size)
		else
			VLC_Auxreturn = mline(PLC_String, PLN_nRow)
		endif
		set memo to VLN_OldSetMemo
	 
	return iif( PLL_returnRowCount, VLN_RowCount, VLC_AuxReturn)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  : PLC_Part : Numer basico para ser Escrito 
*/ Descrição   : Escreve um Numero por extenso 
*/ Retorno     : O extenso Formatado conforme parametrizacao  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_NumtoString
	LPARAMETERS PLC_Part

	private VLC_Digit_1, VLC_Digit_2, VLC_Digit_3, VLC_Dg1, VLC_Dg2, VLC_Dg3, VLC_Dezena

	store "" to VLC_Digit_1, VLC_Digit_2, VLC_Digit_3



	VLC_Dg1 	= val(substr(PLC_Part,1,1))
	VLC_Dg2 	= val(substr(PLC_Part,2,1))
	VLC_Dg3 	= val(substr(PLC_Part,3,1))
	VLC_Dezena 	= val(substr(PLC_Part,2,2))

	if VLC_Dg1 > 0
		VLC_Digit_1 = this.VOA_007_NUMBERS(VLC_Dg1 + 27)
		if val(PLC_Part) > 100 .AND. VAL(PLC_Part) <= 199
			store JUNCTIONSPACE + _100_to_199 + JUNCTIONSPACE to VLC_Digit_1
		endif
	endif
	if VLC_Dezena < 20 .AND. VLC_Dezena > 1
		VLC_Digit_2 = this.VOA_007_NUMBERS(VLC_Dezena)
	else
		if VLC_Dg2 > 0
			VLC_Digit_2 = this.VOA_007_NUMBERS(VLC_Dg2 + 18)
		endif
		if VLC_Dg3 > 0
			VLC_Digit_3 = this.VOA_007_NUMBERS(VLC_Dg3)
		endif
	endif

	if (VLC_Dg1 > 0 .AND. VLC_Dg2 > 0)  .OR. (VLC_Dg3 > 0 .AND. VLC_Dg1 > 0)
		store trim(VLC_Digit_1) + JUNCTIONSPACE + JUNCTIONSTRING + JUNCTIONSPACE to VLC_Digit_1
	endif
	if VLC_Dg2 > 0 .AND. VLC_Dezena > 19 .AND. VLC_Dg3 > 0
		store trim(VLC_Digit_2)+ JUNCTIONSPACE + JUNCTIONSTRING + JUNCTIONSPACE TO VLC_Digit_2
	endif

	return VLC_Digit_1 + VLC_Digit_2 + VLC_Digit_3
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_String : Código para análise         				  
*/ Descrição   		: Retira todos os separadores de PLC_String 
*/ Retorno     		: Retorna a String Limpa 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_ClearSeparator
	LParameters PLC_String
	LOCAL VLC_Caracter, VLN_Counter, VLC_ClearString

	VLC_ClearString  = ""

	*- Retira os caracteres de separação de níveis
	for VLN_Counter = 1 to len(PLC_String)
		VLC_Caracter = subs(PLC_String, VLN_Counter, 1)
		VLN_Pos = 0
		if isdigit(VLC_Caracter) or isalpha(VLC_Caracter)
			VLC_ClearString = VLC_ClearString + VLC_Caracter
		endif
	endfor

	return VLC_ClearString
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_String : Código para análise         				  
*/               	  PLC_Pict   : Mascara para conversão 
*/ Descrição   		: Pesquisa o numero de níveis em PLC_String utilizando PLC_Pict 
*/ Retorno     		: Retorna o numero de níveis da String  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_LevelNumber
	LParameters PLC_String, PLC_Pict
	LOCAL VLN_Counter, VLC_Code, VLC_Caracter, VLN_Occurs, VLC_ClearString

	*- Se a String estiver em branco retorna 0 (sem níveis)
	if empty(PLC_String)
		return 0
	endif

	VLC_Code = PLC_String
	VLC_ClearString = ""

	*- Retira os caracteres de separação de níveis
	VLC_ClearString = this.FOC_ClearSeparator(VLC_Code)

	*- Coloca mascara que esta na Matriz VGA_PICTS
	VLC_Code = transform(VLC_ClearString, PLC_Pict)

	VLL_Ok = .T.
	VLC_ClearString = ""
	VLN_Occurs = 1

	*- Varre toda String pesquisando cada nível do código
	for VLN_Counter = 1 to len(VLC_Code)
		VLC_Caracter = subs(VLC_Code, VLN_counter, 1)
	*--	Se o caracter pesquisado não for Número ou Letra é considerado separador
		VLN_Pos = 0
		if !isdigit(VLC_Caracter) and !isalpha(VLC_Caracter)
			VLN_Pos = VLN_Counter -1
			VLN_Occurs = VLN_Occurs + 1
		else
	*---	String com o código sem mascara
			VLC_ClearString = VLC_ClearString + VLC_Caracter
		endif

	endfor

	return VLN_Occurs
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: Nenhum 
*/ Descrição		: Executa rotina para saída do sistema         	                  	  
*/ Retorno			: .T. : Se sucesso, .F. : c.c.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_ShutDown

	*- Verifica se não existe nenhuma tela ativa no sistema
	if type("VGO_GEN") = "O"
		if VGO_ToolBar.VON_FormInstanceCount = 0
			if this.FON_Msg(31) = 6
				on shutdown

				this.FOL_WorkResource(1, "USR", this.VOC_UserCode, "", .F.)
				this.FOL_Cleanup()
			endif
		else
			*- Verifica se a tela de preview de relatório foi ativada
			if wexist("report designer")
				release window "report designer"
			else
				this.FON_Msg(32)
			endif
		endif

		this.FOL_CloseSemaphoreHandle()
		return .T.
	else
		return .F.
	endif
ENDPROC

*/---------------------------------------------------------------------------------------------
*/ Parametros       : PLU_What           - O que será avaliado 
*/ Descrição        : Verifica o número máximo de níveis 
*/ Retorno          : Retorna o número máximo de níveis  
*/---------------------------------------------------------------------------------------------
PROCEDURE FON_LastNivel
	LParameters PLU_What

	LOCAL VLN_Counter, VLN_Pos, VLC_Caracter, VLN_Size

	VLN_Pos = 1
	VLN_Size = len(PLU_What)
	for VLN_Counter = 1 to VLN_Size
		VLC_Caracter = substr(PLU_What, VLN_counter, 1)
	*--	Se o caracter pesquisado não for Número ou Letra é considerado separador
		if !isdigit(VLC_Caracter) and !isalpha(VLC_Caracter)
			VLN_Pos = VLN_Pos + 1
		endif
	endfor

	return VLN_Pos
ENDPROC

*/--------------------------------------------------------------------------------------------
*/ Parametros       : PLC_Table        : Nome da tabela a ser avaliada  
*/                    PLC_FieldName    : Campo que possui o código da conta 
*/                    PLC_LevelParent  : Nível em que está a conta anterior a que está sendo  
*/                                       cadastrada				  
*/                    PLN_Level        : Nível que está a conta com nível anterior  
*/                    PLC_UkeyChild    : Armazena a ukey da conta com nível anterior  
*/                    PLC_Select       : Select utilizada para pesquisa do nivel anterior  					
*/                    PLL_StandardPlan : Indica que a verificação está sendo feita para o 
*/                                       plano padrão, e a função não deve avaliar  
*/                                       sintético ou analítico 
*/                    PLN_PositionPic  : Indica qual a posição da matriz de picture está  
*/                                       a picture utilizada para consistir o código  
*/ Descrição        : Verifica se a conta com nível anterior é sintética  
*/ Retorno          : Retorna .T. se for sintética ou .F. caso contrário  
*/--------------------------------------------------------------------------------------------
PROCEDURE FOL_IsSintetica
	LParameters PLC_Table, PLC_FieldName, PLC_LevelParent, PLN_Level, PLC_UkeyChild, PLC_Select,;
	PLL_StandardPlan, PLN_PositionPic

	LOCAL VLC_Code, VLN_Niveis, VLL_Ok, VLC_ClearString, VLC_, VLN_Pos, VLN_Counter, VLC_Field ,;
	VLN_NiveisCaracter, VLL_TMP

	*- Amazena a quantidade de níveis da conta com nível inferior. Por exemplo :
	*- se a conta a ser cadastrada for 1.11, o nível a ser armazenado será o
	*- antecessor dessa conta, será nível 1
	VLC_Niveis = transform(PLN_Level -1, "@L 99")


	VLC_Field = evaluate(PLC_Table + "." + PLC_FieldName)

	select (PLC_Table)
	*- Limpa o campo para que este não conste na pesquisa
	blank field (PLC_FieldName)

	*- Coloca mascara no valor do campo a ser pesquisado
	VLC_Code = transform(VLC_Field, "@" + this.VOA_Picts(PLN_PositionPic,1) + " " + ;
	                     this.VOA_Picts(PLN_PositionPic, 2))

	VLL_Ok = .T.
	VLC_ClearString = ""

	*- Varre toda string pesquisando cada nível do código
	for VLN_Counter = 1 to len(VLC_Code)
		VLC_Caracter = substr(VLC_Code, VLN_counter, 1)
		VLN_Pos = 0
	*--	Se o caracter pesquisado não for Número ou Letra é considerado separador
		if !isdigit(VLC_Caracter) and !isalpha(VLC_Caracter)
			VLN_Pos = VLN_Counter - 1
		else
	*---	String com o código sem máscara
			VLC_ClearString = VLC_ClearString + VLC_Caracter
		endif

	*--	Desvio. Se a variável VLN_Pos tiver valor significa que foi encontrado um
	*--	novo nível no código
		if VLN_Pos > 0 and VLC_Niveis >= "01"
			if  this.FOL_FindExpression(PLC_Select, "TMP_Select", VLC_ClearString)
				if !PLL_StandardPlan
	*-----			Se encontrou o nível anterior armazena o tipo de conta (Sintética ou Analítica)
					VLL_Ok = TMP_Select.array_017 = 1
				endif
	*----		Se o nível encontrado for igual ao nível anterior da conta a ser cadastrada
				if evaluate("TMP_Select." + PLC_LevelParent) = VLC_Niveis
	*-----			Armazena a ukey do nível anterior da	conta a ser cadastrada
					PLC_UkeyChild = TMP_Select.ukey
	*-----			Abandona o looping no registro que possui a conta com nível anterior
					exit
				endif
				 this.FOL_CloseTable("TMP_Select")
			endif
		endif
	endfor

	replace (PLC_FieldName) with VLC_Field

	return VLL_Ok
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Alias			 - Arquivo Filho  
*/ 				 	  PLC_Tag			 - Tag do Arquivo Filho 
*/				 	  PLD_Parent 		 - Data do registro pai   						  
*/				 	  PLD_FieldChild	 - Campo Data do Arquivo filho	  
*/				 	  PLC_UkeyParent 	 - Ukey do registro Pai 
*/				 	  PLC_FieldUkeyChild - Campo Ukeyp do arquivo Filho 				  
*/					  PLL_Day			 - Se joga a diferença de dias p/ 				  
*/							   	           algum campo do filho referente a esta dif. 	  
*/					  PLN_Diference		 - Diferença de dias do arquivo pai 
*/					  PLC_FieldChild2	 - Campo de dif. de dias do arquivo filho		  
*/					  PLC_ExcludFields	 _ Campos que devem ser excluídos no 			  
*/					  CreateSqlString													  
*/ Descrição   		: Atualização da Datas do arquivo filho conforme as datas do arquivo  
*/					  pai  
*/ Retorno     		: .t., Se Atualizou com sucesso , .f. 
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_UpdateDateChild
	LParameters PLC_Alias, PLC_Tag, PLD_Parent, PLD_FieldChild, PLC_UkeyParent, PLC_FieldUkeyChild, ;
				PLL_Day, PLN_Diference, PLC_FieldChild2, PLC_ExcludFields

	LOCAL VLC_ExcludFields

	 this.FOL_PushSelect()

	VLC_ExcludFields = iif(empty(PLC_ExcludFields), "", PLC_ExcludFields)

	*- Verifica se o UKEY fornecido existe no arquivo filho
	if !seek(PLC_UkeyParent, PLC_Alias, PLC_Tag)
		this.FON_Msg(48)
		return .F.
	endif

	PLL_Day = iif(empty(PLL_Day), .F., PLL_Day)

	if !PLL_Day
		select (PLC_Alias)
		 this.FOL_PushSelect()
	*-- Varre a Tabela Filha e atualiza as datas dos registros conforme a data do registro pai
		scan while &PLC_FieldUkeyChild = PLC_UkeyParent
			replace &PLD_FieldChild with PLD_Parent
			replace mycontrol  with "1"
			VLL_Add = (left(status,1) <> "W")
			 this.FOL_CreateSQLString(PLC_Alias, left(PLC_Alias,3), VLL_Add, VLL_Add, VLC_ExcludFields)
		endscan
	else
		select (PLC_Alias)
		this.FOL_PushSelect()
	*-- Varre a Tabela Filha e atualiza as datas dos registros conforme a data do registro pai
		scan while &PLC_FieldUkeyChild = PLC_UkeyParent
			replace &PLD_FieldChild with PLD_Parent
			replace	&PLC_FieldChild2 with PLN_Diference
			replace mycontrol  with "1"
			VLL_Add = (left(status,1) <> "W")
			 this.FOL_CreateSQLString(PLC_Alias, left(PLC_Alias,3), VLL_Add, VLL_Add, VLC_ExcludFields)
		endscan
	endif

	*- retorna 2 níveis
	 this.FOL_PopSelect()
	 this.FOL_PopSelect()

	return .T.
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros       : PLD_Data : Data a partir da qual será iniciado o cálculo  
*/                  : PLN_Dias : Dias a serem acrescentados a data inicial  
*/                  : PLL_ConsFeriado : Feriado é considerado dia útil? 
*/                  : PLL_ConsFimSemana : Fim de Semana é considerado dia útil? 
*/			        : PLN_ActionWeekend : Ação para o fim de semana. 1-Retorna; 2-Avança; 
*/			        :	3-Mantem.														  
*/		    	    : PLN_ActionHoliday : Ação para o feriado. 1- Retorna; 2-Avança; 	  
*/					: 	3-Mantem  														  
*/		     	    : PLN_Action: 1 - 1º dia útil antes da data atual se a mesma for 
*/						feriado															  
*/					ou fim de semana. 2 -1º dia útil depois da data atual se a mesma for  
*/					feriado ou fim de semana											  
*/                  : PLC_Pais : Ukey do País a ser pesquisado nos feriados 
*/                  : PLC_UF : Ukey do Estado a ser pesquisado nos feriados 
*/                  : PLC_Cidade : Ukey da Cidade a ser pesquisada nos feriados 
*/ Descrição        : Acrescenta dias na data inicial, tratando dias úteis  
*/ Retorno          : Data 
*/----------------------------------------------------------------------------------------
PROCEDURE FOD_VerificaData
	LParameters PLD_Data, PLN_Dias, PLL_ConsFeriado, PLL_ConsFimSemana, PLN_ActionWeekend,;
		PLN_ActionHoliday, PLN_Action, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey

	LOCAL  VLN_DiasPassados, VLD_DataRetorno, VLN_DiasContados, VLN_DayNumber, VLL_Holiday ,;
			VLN_ActionWeekend, VLN_ActionHoliday, VLL_SunSaturday

	VLN_DiasPassados  = 0
	VLN_DiasContados  = 0
	VLD_DataRetorno	  = PLD_Data
	VLN_ActionWeekend = PLN_ActionWeekend
	VLN_ActionHoliday = PLN_ActionHoliday

	do while VLN_DiasPassados <= PLN_Dias
		VLD_DataRetorno = ctod(dtoc(PLD_Data)) + VLN_DiasContados
		VLN_DayNumber = dow(VLD_DataRetorno)
		VLL_SunSaturday = inlist(VLN_DayNumber,1,7)

	*-- Se não considerar feriados e o dia for diferente de sábado e domingo.
		if empty(PLL_ConsFeriado) and !VLL_SunSaturday and (VLN_DiasPassados>=1)
			VLL_Holiday = this.FOL_ChkDatFeriado(PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, VLD_DataRetorno)
		endif

	*-- Se não considerar fins-de-semana ou não considerar feriados.
		if (!PLL_ConsFimSemana and VLL_SunSaturday) or VLL_Holiday
			VLN_DiasPassados = VLN_DiasPassados - 1
		endif
		VLL_Holiday = .F.
		VLN_DiasPassados = VLN_DiasPassados + 1
		VLN_DiasContados = VLN_DiasContados + 1
	enddo

	*-- Se considerar fim de semana.
	if PLL_ConsFimSemana and VLL_SunSaturday
		do case
	*---	Caso a data de retorno cair em um fim de semana retorna.
			case (VLN_ActionWeekend = 1)
				if VLN_DayNumber = 1
	*-----			(Domingo) - Retorna dois dias.
					VLD_DataRetorno = ctod(dtoc(VLD_DataRetorno))-2
				else
					if VLN_DayNumber = 7
	*----- 				(Sábado) - Retorna um dia.
						VLD_DataRetorno = ctod(dtoc(VLD_DataRetorno))-1
					endif
				endif
	*---	Caso a data de retorno cair em um fim de semana avança.
			case (VLN_ActionWeekend = 2)
				if VLN_DayNumber = 1
	*-----			(Domingo) - Avança um dia.
					VLD_DataRetorno = ctod(dtoc(VLD_DataRetorno))+1
				else
	 				if VLN_DayNumber = 7
	*------			 (Sábado) - Avança dois dias.
						VLD_DataRetorno = ctod(dtoc(VLD_DataRetorno))+2
					endif
				endif
	*-- Caso VLN_ActionWeekend = 3 a data será mantida.
		endcase
	endif

	*- Verifica se a data do vencimento é feriado.
	VLL_Holiday = this.FOL_ChkDatFeriado(PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, VLD_DataRetorno)

	*- Se considerar feriado e for feriado.
	if !empty(PLL_ConsFeriado) and VLL_Holiday
		do case
			case (VLN_ActionHoliday = 1)
	*----		Retorna para a primeira data antes do feriado independete de ser fim de semana.
				VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, -1, -1, .T., .F.)
				VLN_DayNumber = dow(VLD_DataRetorno) && Guarda o dia da semana.

				if PLL_ConsFimSemana and (inlist(VLN_DayNumber, 1, 7)) and (VLN_ActionWeekend=2)
	*-----			PLN_Action=1 - 1º dia util antes do feriado e fim de semana.
	*-----			PLN_Action=2 - 1º dia util após o feriado e fim de semana.
					VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, iif((PLN_Action=1), -1,1), iif((PLN_Action=1), -1,1), .T., PLL_ConsFimSemana)
				else
					if PLL_ConsFimSemana and (inlist(VLN_DayNumber, 1, 7)) and (VLN_ActionWeekend=1)
						VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, -1, -1, .T., .T.)
					else
						if !PLL_ConsFimSemana
							VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, -1, -1, .T., .T.)
						endif
					endif
				endif
	*---	Caso a data de retorno cair em um feriado vai para o próximo dia util.
			case (VLN_ActionHoliday = 2)
				VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, 1, 1, .T., .F.)
				VLN_DayNumber = dow(VLD_DataRetorno)

				if PLL_ConsFimSemana and (inlist(VLN_DayNumber, 1, 7)) and (VLN_ActionWeekend=1)
	*-----			PLN_Action=1 - 1º dia util antes do feriado e fim de semana.
	*-----			PLN_Action=2 - 1º dia util após o feriado e fim de semana.
					VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, iif((PLN_Action=1), -1,1), iif((PLN_Action=1), -1,1), .T., PLL_ConsFimSemana)
				else
					if PLL_ConsFimSemana and (inlist(VLN_DayNumber, 1, 7)) and (VLN_ActionWeekend=2)
						VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, 1, 1, .T., .T.)
					else
						if !PLL_ConsFimSemana
							VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, 1, 1, .T., .T.)
						endif

					endif
				endif
	*-- Caso VLN_ActionHoliday = 3 a data será mantida.
		endcase
	else
		if VLL_Holiday
			VLD_DataRetorno = this.FOD_ProximoDiaUtil(VLD_DataRetorno, PLC_A22Ukey, PLC_A23Ukey, PLC_A24Ukey, iif((VLN_ActionWeekend=1 and PLL_ConsFimSemana),-1, 1), iif((VLN_ActionWeekend=1 and PLL_ConsFimSemana), -1,1), .T., (VLN_ActionWeekend<>3))
		endif
	endif
	return VLD_DataRetorno
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros  : PLN_ValueToRound : Valor que será arredondado  
*/				 PLN_Regra     : Número da Matriz de Arredondamento Padrão  
*/				 PLL_OnlyRound : Em vez de arredondar, fornece a diferença no processo  
*/				 PLN_FatorAux  : Critério de arredondamento específico  
*/ Descrição   : Análise de uma condição de pagamento do tipo NORMAL  
*/ Retorno     : .T. se a condição for avaliada + matrizes VGA_AVenc, VGA_AParc 
*/ Data da Ultima alteração : 24/02/1998 - Gustavo Haidar Viotto  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Arredonda
	LParameters PLN_ValueToRound, PLN_Regra, PLL_OnlyRound, PLN_FatorAux

	LOCAL  VLN_RoundBase, VLN_Return

	with this
		if between(PLN_Regra, 1, alen(.VOA_014, 1))
			if type(".VOA_014[PLN_Regra]") = "N"
				VLN_RoundBase = .VOA_014[PLN_Regra]
			else
				VLN_RoundBase = VAL(.VOA_014[PLN_Regra])
			endif
		else
			VLN_RoundBase = 0
		endif
	endwith

	if !empty(PLN_FatorAux)
		VLN_RoundBase = PLN_FatorAux
	endif
	
	if(VLN_RoundBase # 0)
		VLN_return = iif((PLN_ValueToRound/VLN_RoundBase) - int(PLN_ValueToRound/VLN_RoundBase) > 0,;
	           int(PLN_ValueToRound/VLN_RoundBase) + 1, ;
	           int(PLN_ValueToRound/VLN_RoundBase)) * VLN_RoundBase
		if PLL_OnlyRound
			return VLN_return - PLN_ValueToRound
		else
			return VLN_Return
		endif

	else
		if PLL_OnlyRound
			return 0
		else
			return PLN_ValueToRound
		endif
	endif
ENDPROC


*/--------------------------------------------------------------------------------------------
*/ Parametros		: PLM_CommandsInformation   : Memo com as do comando  
*/					  PLC_CommandToSee		    : Propriedade a ser consultada  
*/					  PLN_Position 				: Número do comando 
*/                    PRC_CommandName           : nome do comando 
*/                    PRN_Position              : guarda a última posição encontrada  
*/ Descrição		: Retorna memo com os comandos separados por \\BEGIN <PLC_CommandToSee> 
*/                    \\END <PLC_CommandToSee> 
*/ Retorno			: Conteúdo do campo memo 
*/--------------------------------------------------------------------------------------------
PROCEDURE FOC_ChkCommands
	LParameters PLM_CommandsInformation, PLC_CommandToSee, PLN_Position, PRC_CommandName, PRN_Position

	LOCAL VLN_i, VLN_j, VLC_Ret

	if empty(PLN_Position)
		PLN_Position = 1
	endif

	if !empty(PLC_CommandToSee)
		PLC_CommandToSee = " " + upper(alltrim(PLC_CommandToSee))
	endif
	PRC_CommandName  = ""

	*- procura o comando
	PRN_Position = at("\\BEGIN" + PLC_CommandToSee, upper(PLM_CommandsInformation), PLN_Position)
	if PRN_Position = 0
		return ""
	endif
	VLC_Ret = subs(PLM_CommandsInformation, PRN_Position + 7 + len(PLC_CommandToSee))
	VLN_j = at(chr(13), VLC_Ret)

	*- pega o nome do comando
	PRC_CommandName = upper(alltrim(subs(VLC_Ret, 1, VLN_j - 1)))
	VLC_Ret = subs(VLC_Ret, VLN_j + 1)
	if subs(VLC_Ret, 1, 1) = chr(10)
		VLC_Ret = subs(VLC_Ret, 2)
	endif

	VLN_j = at("\\END" + PLC_CommandToSee, upper(VLC_Ret))
	if VLN_j > 0
		VLC_Ret = subs(VLC_Ret, 1, VLN_j - 1)
	endif

	*- retira até o último chr(13)
	VLN_i = rat(chr(13), VLC_Ret)
	if VLN_i > 0
		VLC_Ret = subs(VLC_Ret, 1, VLN_i - 1)
	endif


	return VLC_Ret
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_Num : número a ser convertido 
*/ Descrição		: Converte um número em sua representação romana                 	  
*/ Retorno			: representação romana do número  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_Dectoroman
	LParameters PLN_Num

	LOCAL VLN_j, VLC_Unidade, VLC_Ret

	VLN_j = 1
	VLC_Unidade = "IVXLCDMQWY"
	VLC_Ret = ""

	*- insere os caracteres à esquerda
	do while PLN_Num > 0
		VLN_Mod = mod(PLN_Num, 10)
		do case
	        case inlist(VLN_Mod, 1, 2, 3)
	        	VLC_Ret = replicate(subs(VLC_Unidade, VLN_j, 1), VLN_Mod) + VLC_Ret

	        case inlist(VLN_Mod, 5, 6, 7, 8)
	        	VLC_Ret = replicate(subs(VLC_Unidade, VLN_j, 1), VLN_Mod - 5) + VLC_Ret
	        	VLC_Ret = subs(VLC_Unidade, VLN_j + 1, 1) + VLC_Ret

	        case VLN_Mod = 4
	        	VLC_Ret = subs(VLC_Unidade, VLN_j + 1, 1) + VLC_Ret
	        	VLC_Ret = subs(VLC_Unidade, VLN_j, 1) + VLC_Ret

	        case VLN_Mod = 9
	        	VLC_Ret = subs(VLC_Unidade, VLN_j + 2, 1) + VLC_Ret
	        	VLC_Ret = subs(VLC_Unidade, VLN_j, 1) + VLC_Ret
		endcase
		PLN_Num = int(PLN_Num / 10)
		VLN_j = VLN_j + 2
	enddo

	return VLC_Ret
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_Num : número (no format ABC) a ser convertido 
*/ Descrição		: Converte um número com a representação ABC para o formato numerico
*/ Retorno			: O número  convertido
*/----------------------------------------------------------------------------------------
PROCEDURE FON_abcToDec
	LParameters PLC_Num

	local VLN_Return, VLN_Max, vln_i, VLC_Char
	VLN_Return = 0
	VLN_Max    = len(PLC_Num)

	for vln_i = VLN_Max to 1 Step -1

		VLC_Char = substr(PLC_Num, VLN_i, 1)
		if isAlpha(VLC_Char)
			VLN_Num  = asc( upper(VLC_Char) ) - 64
			VLN_Return = VLN_Return + (VLN_Num  * (26^(VLN_Max - vln_i)) )
		else
			VLN_Return = 0
			Exit
		endif
		
		
	next
	Return VLN_Return 

ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_Num : número a ser convertido 
*/ Descrição		: Converte um número em sua representação ABC 	                  	  
*/ Retorno			: representação ABC do número 
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_Dectoabc
	LParameters PLN_Num

	LOCAL VLN_j, VLC_Ret

	VLC_Ret = ""
	VLN_j = 0

	*- adiciona os caracteres à esquerda
	do while PLN_Num > 0
		VLN_j = mod(PLN_Num, 26)
	    if VLN_j = 0
			VLC_Ret = 'Z' + VLC_Ret
			PLN_Num = int(PLN_Num / 26) - 1
	    else
			VLC_Ret = chr(VLN_j + 64) + VLC_Ret
			PLN_Num = int(PLN_Num / 26)
		endif
	enddo

	return VLC_Ret
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros  : PLN_Valor     : Valor a ser zerado						  
*/ Descrição   : Zera o valor se o registro em livros estiver cancelado 
*/ Retorno     : O valor validado									  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_MudaVal
	LPARAMETER PLN_Valor

	*-	Verifica se é um registro cancelado para zerar
	if h03_006_n = 1
		PLN_Valor = 0
	endif

	return PLN_Valor
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_ClassDefault : nome da classe padrão, ex: adodb.connection  
*/              	  PLC_ClassCustom  : nome da classe que contém os eventos customizados
*/ Descrição   		: Cria um objeto da classe PLC_ClassDefault com os eventos customiza- 
*/                    dos da classe PLC_ClassCustom.  
*/ Retorno     		: ponteiro ao objeto criado, ou .NULL. - caso contrário.  
*/----------------------------------------------------------------------------------------
PROCEDURE FOO_CreateObjectWithEvents
	LParameters PLC_ClassDefault, PLC_ClassCustom

	LOCAL VLO_Object, VLO_Custom, VLN_Error

	VLO_Object = createobject(PLC_ClassDefault)
	VLO_Custom = createobject(PLC_ClassCustom, VLO_Object)

	VLN_Error = this.VOO_VFPCOMUTIL.BindEvents(VLO_Object, VLO_Custom)

	return VLO_Object
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_Id	   : Código de identificação do Tipo	    	    	  
*/					  PLC_T08Id	   : Código de identificação da Totalização				  
*/ 					  PLC_J06Par   : Código da nota			  
*/ 					  PLC_J06UkeyP : Ukey da Nota		  
*/ Descrição		: Procura os totais da Nota para impressão das Bases dos Impostos	  
*/ Retorno			: O Total das bases		 
*/----------------------------------------------------------------------------------------
PROCEDURE FOB_CalcTotNF
	LParameters PLN_Id, PLC_T08Id, PLC_J06Par, PLC_J06UkeyP

	LOCAL VLN_Id, VLC_T08Id, VLC_J06Par, VLC_J06UkeyP, VLB_Return
	
	VLN_Id	 	 = PLN_Id
	VLC_T08Id	 = upper(PLC_T08Id)
	VLC_J06Par	 = upper(PLC_J06Par)
	VLC_J06UkeyP = PLC_J06UkeyP

	if substr(VLC_J06Par,1,1) = "J"
		if VLN_Id = 0
			with this
				.FOL_SetParameter(1, VLN_Id)
				.FOL_SetParameter(2, VLC_T08Id)
				.FOL_SetParameter(3, VLC_J06Par)
				.FOL_SetParameter(4, VLC_J06UkeyP)
				.FOL_ExecuteCursor("J06_TOT", "J06TT", 4)
			endwith
		else
			with this
				.FOL_SetParameter(1, VLN_Id)
				.FOL_SetParameter(2, VLC_J06Par)
				.FOL_SetParameter(3, VLC_J06UkeyP)
				.FOL_ExecuteCursor("J06_TOT1", "J06TT", 3)
			endwith
		endif
		VLB_Return = J06TT.J06_001_B
	else
		if VLN_Id = 0
			with this
				.FOL_SetParameter(1, VLN_Id)
				.FOL_SetParameter(2, VLC_T08Id)
				.FOL_SetParameter(3, VLC_J06Par)
				.FOL_SetParameter(4, VLC_J06UkeyP)
				.FOL_ExecuteCursor("E04_TOT", "E04TT", 4)
			endwith
		else
			with this
				.FOL_SetParameter(1, VLN_Id)
				.FOL_SetParameter(2, VLC_J06Par)
				.FOL_SetParameter(3, VLC_J06UkeyP)
				.FOL_ExecuteCursor("E04_TOT1", "E04TT", 3)
			endwith
		endif
		VLB_Return = E04TT.E04_001_B
	endif

	return VLB_Return
ENDPROC

*-- Função que ira retornar em uma string as parcelas da futura
*/----------------------------------------------------------------------------------------/*
*/ Função Objeto	: FOC_FatParc       	                                              /*
*/ Parâmetros		: PLC_J05Par   : Parcela da NF de entrada ou saida                    /*
*/ 					  PLC_J05UkeyP : Ukey da parcela                                      /*
*/ Descrição		: Procura as parcelas da notas fiscal    							  /*
*/ Retorno			: As parcelas                   		                              /*
*/ Grupo			: Database                                                            /*
*/ Procura por		: Parcelas                                                            /*
*/ Última alteração	: 10/02/00   														  /* 
*/ Alterado por		: Alexandre Torre													  /*
*/ Versão			: 2																	  /*
*/----------------------------------------------------------------------------------------/*
PROCEDURE FOC_FatParc
	LParameters PLC_J05Par, PLC_J05UkeyP

	LOCAL VLC_J05Par, VLC_J05UkeyP, VLC_Return, VLC_Alias

	VLC_Alias	  = alias()
	VLC_Return	 = ""
	VLC_J05Par	 = upper(PLC_J05Par)
	VLC_J05UkeyP = PLC_J05UkeyP

	if substr(VLC_J05Par,1,1) = "J"
		with this
			.FOL_SetParameter(1, VLC_J05Par)
			.FOL_SetParameter(2, VLC_J05UkeyP)
			.FOL_ExecuteCursor("J05_FRM", "PARJ05", 2)
		endwith

		select PARJ05
		scan
			VLC_Return = VLC_Return + PARJ05.J05_001_C + DTOC(PARJ05.J05_005_D);
						 + TRANS(PARJ05.J05_004_B,"999,999,999,999.99") + CRLF
		endscan

		use in PARJ05
	else
		with this
			.FOL_SetParameter(1, VLC_J05Par)
			.FOL_SetParameter(2, VLC_J05UkeyP)
			.FOL_ExecuteCursor("E05_FRM", "PARE05", 2)
		endwith

		select PARE05
		scan
			VLC_Return = VLC_Return + PARE05.E05_001_C + DTOC(PARE05.E05_005_D);
						 + TRANS(PARE05.E05_004_B,"999,999,999,999.99") + CRLF
		endscan

		use in PARE05
	endif

	if !empty(VLC_Alias)
		select (VLC_Alias)
	endif

	RETURN VLC_Return
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_Id	   : Código de identificação do Tipo 1- Aliqüota		  
*/									 , 2- Valor											  
*/ 					  PLC_J22Par   : Código da item			  
*/ 					  PLC_J22UkeyP : Ukey do item		  
*/					  PLC_A40Id	   : Código de identificação do Imposto    				  
*/ Descrição		: Procura as aliqüotas dos itens da nota 							  
*/ Retorno			: A Aliqüota do Imposto ou Valor		  
*/----------------------------------------------------------------------------------------
PROCEDURE FOB_AliqImp
	LParameters PLN_Id, PLC_J22Par, PLC_J22UkeyP, PLC_A40Id

	LOCAL VLN_Id, VPLC_J22Par, VPLC_J22UkeyP, VPLC_A40Id, VLC_Alias, VLB_Return

	VLC_Alias	  = alias()
	VLN_Id	 	  = iif(empty(PLN_Id),1,iif((PLN_Id > 2), 2,PLN_Id))
	VLC_A40Id	  = upper(PLC_A40Id)
	VLC_J22Par	  = upper(PLC_J22Par)
	VLC_J22UkeyP  = PLC_J22UkeyP

	if SUBSTR(VLC_J22Par,1,1) = "J"
		with this
			.FOL_SetParameter(1, VLC_J22Par)
			.FOL_SetParameter(2, VLC_J22UkeyP)
			.FOL_SetParameter(3, VLC_A40Id)
			.FOL_ExecuteCursor("J22_TOTALQ", "J22_TOTALQ", 3)
		endwith

		if reccount("J22_TOTALQ") > 0
			if VLN_Id = 1
				if J22_TOTALQ.array_036 = 1
					VLB_Return = J22_TOTALQ.j22_003_b
				else
					VLB_Return = 0
				endif
			else
				VLB_Return = J22_TOTALQ.j22_004_b
			endif
		else
			VLB_Return = 0
		endif

		use in J22_TOTALQ
	else
		with this
			.FOL_SetParameter(1, VLC_J22Par)
			.FOL_SetParameter(2, VLC_J22UkeyP)
			.FOL_SetParameter(3, VLC_A40Id)
			.FOL_ExecuteCursor("E12_TOTALQ", "E12_TOTALQ", 3)
		endwith

		if reccount("E12_TOTALQ") > 0
			if VLN_Id = 1
				if E12_TOTALQ.array_036 = 1
					VLB_Return = E12_TOTALQ.e12_003_b
				else
					VLB_Return = 0
				endif
			else
				VLB_Return = E12_TOTALQ.e12_004_b
			endif
		else
			VLB_Return = 0
		endif

		use in e12_TOTALQ
	endif

	if !empty(VLC_Alias)
		select (VLC_Alias)
	endif

	return VLB_Return
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros  : PLN_Minute     : Campo de Minutos para ser convetido 
*/ Descrição   : Retorna a conversão de minutos em horas  
*/ Retorno     : Horas 
*/ Data da Ultima alteração : 12/01/2000 - Alexandre Torre  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_ConvertHour
	LParameters PLN_Minute, PLN_BLANK

	LOCAL VLC_RETURN, VLC_HORA, VLC_MINUTOS, VON_Minute
	IF EMPTY(PLN_BLANK)
		PLN_BLANK = 0
	ENDIF
	VON_Minute = PLN_Minute

	IF TYPE('VON_Minute') = 'C'
		VLC_HORA    = LEFT(VON_Minute, AT(':', VON_Minute) - 1)
		VLC_MINUTOS = SUBSTR(VON_Minute, AT(':', VON_Minute) + 1, 2)
		VLC_RETURN  = INT(VAL(VLC_MINUTOS) + (VAL(VLC_HORA) * 60))
	ELSE
		IF PLN_BLANK = 0 AND VON_Minute = 0
			VLC_RETURN = ''
		ELSE
			VLC_HORA    = PADL(INT(VON_Minute / 60), 3, '0')
			IF AT('0', VLC_HORA, 1) = 1
				VLC_HORA = STRTRAN(VLC_HORA, '0', '', 1, 1)
			ENDIF
			VLC_MINUTOS = PADL(MOD(VON_Minute, 60), 2, '0')
			VLC_RETURN  = VLC_HORA + ":" + VLC_MINUTOS
		ENDIF
	ENDIF

	RETURN VLC_RETURN
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Parametros		: PLN_ValueToRound : Valor a ser arredondado            	    	  
*/ 					  PLN_Accurracy : Precisão após a vírgula 
*/					  PLN_RoundType : 0-Standard (+.5), 1-Baixo, 2-Cima 
*/ Descrição		: Arredonda valores no sistema, evitando problemas de parcelamento  
*/ Retorno			: Valor arredondado  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Round
	LParameters PLN_ValueToRound, PLN_Accurracy, PLN_RoundType

	LOCAL VLN_Return, VLN_AuxValue, VLN_Accurracy, VLN_RoundType

	VLN_Accurracy = iif(empty(PLN_Accurracy) and vartype(PLN_Accurracy) = "L", this.VON_Accurracy, PLN_Accurracy)
	VLN_RoundType = iif(empty(PLN_RoundType) and vartype(PLN_RoundType) = "L", this.VON_RoundType, PLN_RoundType)

	VLN_AuxValue = int(10 ^ VLN_Accurracy)

	do case
		case empty(VLN_RoundType) 	&& Arredondamento padrão do FOX
			*- É preciso consistir se o número de casas decimais é até 4, porque se for mais a função NTOM não vai
			*- considerar
			if VLN_Accurracy <= 4
				*- É preciso eliminar as características da varíavel antes de arredondar
				VLN_Return = round(mton(ntom(PLN_ValueToRound)), VLN_Accurracy)
			else
				VLN_Return = round(PLN_ValueToRound, VLN_Accurracy)
			endif
		case VLN_RoundType = 1		&& Arredondamento para baixo, ignorando demais casas
			*- Added by Denis
			*- Problema no fox quando uma variável é proveniente de um campo tipo B (Dupla Precisão)
			*- Exemplo: create cursor teste (number b(18))
			*-			append blank
			*-			replace number with 1000*1.003 -> Deveria ser 1003
			*-			? floor(number) -> resulta 1002
			VLN_ValueToRound = ntom(PLN_ValueToRound) && Limitado a 4 posições, espero que não seja necessário mais que isso
			*- End

			VLN_Return = floor(VLN_ValueToRound * VLN_AuxValue)/VLN_AuxValue

		case VLN_RoundType = 2
			*- Added by Denis
			*- Problema no fox quando uma variável é proveniente de um campo tipo B (Dupla Precisão)
			*- Exemplo: create cursor teste (number b(18))
			*-			append blank
			*-			replace number with 1000*1.003 -> Deveria ser 1003
			*-			? floor(number) -> resulta 1002
			VLN_ValueToRound = ntom(PLN_ValueToRound) && Limitado a 4 posições, espero que não seja necessário mais que isso
			*- End
			VLN_Return = ceil(VLN_ValueToRound * VLN_AuxValue)/VLN_AuxValue
	endcase

	return VLN_Return
ENDPROC


*/-----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Code	     : Código para análise 
*/               	  PLC_Pict       : Mascara para conversão  
*/               	  PLC_Alias      : Arquivo a ser pesquisado  
*/                    PLC_CursorName : Armazena o nome do cursor a ser executado do  
*/                                     caso se for SQL 
*/ Descrição   		: Checa a Consistência de níveis 
*/ Retorno     		: se código válido retorna o numero de níveis, senão retorna 0 
*/-----------------------------------------------------------------------------------------
PROCEDURE FON_ChkConsistNiveis
	LParameters PLC_Code, PLC_Pict, PLC_Alias, PLC_CursorName

	LOCAL VLC_Code, VLC_ClearString, VLC_Condition, VLN_Counter, VLC_Caracter, VLN_Pos,;
		  VLL_Ok, VLN_Param, VLC_Alias

	VLN_Param = parameters()
	VLC_Condition = ""
	VLC_Code = PLC_Code
	VLC_ClearString = ""

	*- Retira os caracteres de separação de níveis
	VLC_ClearString = this.FOC_ClearSeparator(VLC_Code)

	*- Coloca mascara que esta na Matriz VGA_PICTS
	VLC_Code = transform(VLC_ClearString, PLC_Pict)
	store "" to VLC_Parent, VLC_ClearString
	VLN_Pos = 0

	*- Varre toda String pesquisando cada nível do código
	for VLN_Counter = len(VLC_Code) to  1  step -1
		VLC_Caracter = substr(VLC_Code, VLN_Counter, 1)
	*--	Se o caracter pesquisado não for Número ou Letra é considerado separador
		if !isdigit(VLC_Caracter) and !isalpha(VLC_Caracter)
			VLC_Parent = substr(VLC_Code, 1, VLN_Counter)
			exit
		endif
	endfor

	*- Retira os caracteres de separação de níveis
	VLC_ClearString = this.FOC_ClearSeparator(VLC_Parent)

	*- Verifica se codigo tem um pai
	if this.FON_LevelNumber(PLC_Code, PLC_Pict) = 1
	*---Se existir somente um nível a consistência é OK
		VLL_Ok = .T.
	else
	*-	Armazena .T. se econtrou o código com nível anterior, ou .F. caso contrário
		 this.FOL_SetParameter(1, VLC_ClearString)
		 this.FOL_ExecuteCursor(PLC_CursorName, PLC_CursorName, 1)
		VLL_Ok = !eof(PLC_CursorName)
	endif

	return iif(VLL_Ok, this.FON_LevelNumber(PLC_Code, PLC_Pict), 0)
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_ReportUkey : ukey do relatório                           		  
*/ Descrição		: Procura os nomes dos relatórios        							  
*/ Retorno			: O Nome do relatório           		  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_NameReport
	lparameters PLC_ReportUkey

	local VLC_Return, VLN_RepoCounter, VLN_Counter, VLC_Separator
	VLC_Return = ""
	VLC_Separator = ""
	local array VLA_RetArray[1]

	VLN_RepoCounter = this.FON_GetReports(1, PLC_ReportUkey, @VLA_RetArray)
	
	*- Inverto o scan para agilizar a busca
	for VLN_Counter = VLN_RepoCounter to 1 step -1
		if PLC_ReportUkey == VLA_RetArray[VLN_Counter, 1]
			VLC_Return = VLA_RetArray[VLN_Counter, 11]
		endif
	endfor

	return VLC_Return
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_What	  : Armazena quais são os números a ser feito o MMC 
*/                                  separados por vírgula 
*/ Descrição		: Calcula MMC  
*/ Retorno			: MMC  
*/----------------------------------------------------------------------------------------
PROCEDURE FON_Mmc
	LParameters PLC_What

	LOCAL VLN_Mult, VLN_Counter1, VLN_Times, VLL_Divisivel, VLA_Array, VLN_Counter2

	store 1 to VLN_Mult, VLN_Times
	VLN_Counter1 = 2

	PLC_What = alltrim(PLC_What)

	if right(PLC_What, 1) <> ","
		PLC_What = PLC_What + ","
	endif

	*- Verifica quantos números será preciso colocar na Array
	VLN_Occurs = occurs(",", PLC_What)
	dimension VLA_Array[VLN_Occurs]
	for VLN_Counter2 = 1 to VLN_Occurs
		VLC_Substr  = substr(PLC_What,1, at(",", PLC_What)-1)
		VLA_Array[VLN_Counter2] = val(VLC_Substr)
	*--	Armazena o maior número a ser tirado o MMC
		VLN_Times = max(VLN_Times, VLA_Array[VLN_Counter2])
		PLC_What = right(PLC_What,len(PLC_What)-(len(VLC_Substr)+1))
	endfor

	do while VLN_Counter1 <= VLN_Times
		VLL_Divisivel = .F.
		for VLN_Counter2 = 1 to VLN_Occurs
	*---	Se for maior do que e divisível
			if VLA_Array[VLN_Counter2] > 0 and mod(VLA_Array[VLN_Counter2], VLN_Counter1) = 0
				if !VLL_Divisivel
					VLN_Mult = VLN_Counter1 * VLN_Mult
				endif
				VLA_Array[VLN_Counter2] = VLA_Array[VLN_Counter2] / VLN_Counter1
				VLL_Divisivel = .T.
			endif
		endfor
		if !VLL_Divisivel
			VLN_Counter1 = VLN_Counter1 + 1
		endif
	enddo

	return VLN_Mult
ENDPROC


*/---------------------------------------------------------------------------------------------
*/ Parametros		: PLC_VersionStr  - String com a versão do produto
*/					  PLN_TypeVersion - Indica o tipo de Versão: Comercial, Apresentação, Beta 
*/ Descrição		: Cria o arquivo de versão do produto
*/ Retorno			: Lógico
*/---------------------------------------------------------------------------------------------
PROCEDURE FOL_WriteVersion
	lparameters PLC_VersionStr, PLN_TypeVersion, PLC_SSPCode

	VLL_Return = .T.
	
	if empty(PLN_TypeVersion)
		PLN_TypeVersion = 0
	endif

	if !empty(PLC_VersionStr)

		VLC_Version   = alltrim(PLC_VersionStr)
		VLC_SSPCode   = alltrim(PLC_SSPCode)
		VLC_SSPCode   = this.FOC_UserChangeControl(VLC_SSPCode, "SISCORP2000", .t.)
		VLC_Version   = this.FOC_UserChangeControl(VLC_Version, "SISCORP2000", .t.)
		VLC_Version   = VLC_Version + CRLF + str(PLN_TypeVersion) + CRLF + VLC_SSPCode

		VLL_Return    = strtofile(VLC_Version, 'version.txt') > 0
		
	endif

	return VLL_Return
ENDPROC



*/----------------------------------------------------------------------------------------
*/ Descrição		: Retorna a versão do sistema corrente
*/ Retorno			: Versão do sistema
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_GetVersion
	local VLC_Version, VLC_SPVersion
	VLC_Version = mline(this.VOC_VersionCfgFile, 1)
	VLC_Version =  this.FOC_UserChangeControl(VLC_Version, "SISCORP2000", .t.)

	this.VON_TypeVersion = int(val(mline(this.VOC_VersionCfgFile, 2)))

	VLC_SPVersion = mline(this.VOC_VersionCfgFile, 3)
	VLC_SPVersion =  this.FOC_UserChangeControl(VLC_SPVersion, "SISCORP2000", .t.)

	this.VOC_SPVersion = VLC_SPVersion

	return VLC_Version
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Descrição		: Retorna a versão do sistema corrente
*/ Retorno			: Versão do sistema
*/----------------------------------------------------------------------------------------
PROCEDURE FON_TypeVersion
	return this.VON_TypeVersion
ENDPROC


PROCEDURE Destroy
	*-- Executa o Cleanup se necessário.
	IF !this.VOL_IsClean
	  this.FOL_CleanUp()
	ENDIF
ENDPROC


*-- Verifica se há registros na tabela passda como parâmetro.
PROCEDURE FOL_NoneRecno
	*/----------------------------------------------------------------------------------------/*
	*/ Função Objeto    : FOL_NoneRecno                                                       /*
	*/ Parâmetros  		: PLC_Table        - Tabela a ser avaliada                            /*
	*/ Descrição   		: Verifica se há registros na tabela passada como parâmetro           /*
	*/ Retorno     		: .T. Se estiver em EOF  ou .F. c.c                                   /*
	*/ Grupo       		: Database                                                            /*
	*/ Procura por 		: Tabela                                                              /*
	*/ Última alteração	: 31/07/1998                                                          /*
	*/ Alterado por 	: Rodrigo                                                             /*
	*/ Versão 			: 1                                                                   /*
	*/----------------------------------------------------------------------------------------/*
	LParameters PLC_Table

	LOCAL VLL_Return

	this.FOL_PushSelect()

	select (PLC_Table)
	this.FOL_PushSelect()
	go top in (PLC_Table)
	VLL_Return = eof(PLC_Table)
	this.FOL_PopSelect()
	this.FOL_PopSelect()

	return VLL_Return
ENDPROC

	*-- Mensagem fiscal da Nota Fiscal de Venda, baseado nas mensagens dos impostos.
	PROCEDURE FOC_NfMessage
		LPARAMETERS PLC_J10Ukey, PLC_BetweenChar, PLL_All

		*- O parâmetro PLL_All indica se a função deve fazer o link com todas as mensagens, ou deve
		*-   verificar se há a repetição de ukey's
		*- O parâmetro PLC_BetweenChar armazena o caractér que será colocado entre as mensagens, se
		*-   nada for passado utiliza 1 espaço.

		LOCAL VLC_CurAlias, VLC_Return, VLC_LastH02Ukey, VLC_BetChar, VLN_Parameters, VLN_Recc


		VLN_Parameters = parameter()
		VLC_CurAlias = alias()
		VLC_LastH02Ukey = space(20)
		VLC_BetChar = iif(VLN_Parameters <= 1, " ", PLC_BetweenChar)

		with this
			.FOL_SetParameter(1, PLC_J10Ukey)
		*-	Esse cursor traz todos os registros de impostos de uma NF (apenas J10)
			.FOL_ExecuteCursor("J10_MSG", "J10_MSG", 1)
		endwith

		VLC_Return = ""
		VLN_Recc = recc()

		select J10_MSG
		index on h02_ukey tag code
		set order to code

		VLN_Recc = recc()

		scan
		*-	Respeitando o parâmetro passado
			if (VLC_LastH02Ukey <> j10_msg.h02_ukey) or PLL_All or ( empty(j10_msg.h02_ukey) or isnull(j10_msg.h02_ukey) )
				if !empty(j10_msg.j22_012_m) and !isnull(j10_msg.j22_012_m)
					VLC_Return = VLC_Return + alltrim(j10_msg.j22_012_m)
					if recno() <> VLN_Recc
						VLC_Return = VLC_Return + VLC_BetChar
					endif
				endif
			endif
			VLC_LastH02Ukey = nvl(j10_msg.h02_ukey, "")
		endscan

		*- Fecha o arquivo antes de sair e retorna ao alias anterior
		use in J10_MSG

		if !empty(VLC_CurAlias)
			select (VLC_CurAlias)
		endif

		return VLC_Return
	ENDPROC


	*-- Função que serapa os caracteres com space, utilizado para formulários de preenchimento. ex. Seguro Desemprego
	PROCEDURE FOC_Separator
		*/----------------------------------------------------------------------------------------/*
		*/ Função Objeto	: FOC_Separator			                                              /*
		*/ Parâmetros		: VLC_String  : String a ser separada			          	    	  /*
		*/ 					  VLN_Spaces  : Número de espaços entre os caracteres				  /*
		*/ Descrição		: Gera uma String com separacao de letras                      		  /*
		*/ Retorno			: String separada		                                              /*
		*/ Grupo			: Report, SCX                                                         /*
		*/ Procura por		: String                                                              /*
		*/ Última alteração	: 31/12/2000 														  /* 
		*/ Alterado por		: Tânia Sayumi Kondo											      /*
		*/ Versão			: 2																	  /*
		*/----------------------------------------------------------------------------------------/*
		LPARAMETERS VLC_String, VLN_Spaces
		LOCAL  VLN_Count, VLC_StrReturn
		if empty(VLN_Spaces)
		 	VLN_Spaces = 1
		endif
		 
		if empty(VLC_String) 
		 	VLC_String = ""
		endif
		 
		VLC_StrReturn = ""

		if !isnull(VLC_String)
			for VLN_Count = 1 to len(VLC_String)
				VLC_StrReturn = VLC_StrReturn + substr(VLC_String, VLN_Count, 1)
				if VLN_Count < len(VLC_String)
					VLC_StrReturn = VLC_StrReturn + space(1*VLN_Spaces)
				endif
			next
		endif
		 
		return VLC_StrReturn
	ENDPROC


*-- Grava a resposta do parâmetro no arquivo (Y41)
	PROCEDURE FOL_SetSets
		*/----------------------------------------------------------------------------------------/*
		*/ Função Objeto	: FOL_SetSets			                                              /*
		*/ Parâmetros		: PLC_Y08Ukey : Id do Parâmetro      			          	    	  /*
		*/ 					  PLC_Content : Conteúdo a ser gravado sempre em formato caracter     /*
		*/ Descrição		: Grava a resposta do parâmetro                               		  /*
		*/ Retorno			: .T. ou .F. c.c.                                                     /*
		*/ Grupo			: PARÂMETROS                                                          /*
		*/ Procura por		: Parêmtros                                                           /*
		*/ Última alteração	: 23/02/2001 														  /* 
		*/ Alterado por		: Rodrigo														      /*
		*/ Versão			: 1																	  /*
		*/----------------------------------------------------------------------------------------/*
				
		LPARAMETERS PLC_Y08Ukey, PLC_Content

		LOCAL VLN_Return, VLU_ParContent, VLN_Par, VLL_Return

		VLU_ParContent = .null.

		this.FOL_UseInPackage("y08", "y08")

		VLL_Return = .T.

		with this
			if seek(PLC_Y08Ukey, "y08", "ukey")
				if .FOL_BeginTransaction("Y41")
					if !.FOL_FindExpression("Y41_Y08_UKEY", "Y41T", PLC_Y08Ukey)
						.FOL_NewRecordToForm("Y41T")
						replace y41t.y08_ukey with y08.ukey in Y41T
					endif
					
					do case
						*- Tipo caracter
						case y08.y08_012_c = "CHARACTER"
							replace y41t.y41_001_m with PLC_Content in Y41T
							VLU_ParContent = nvl(y41t.y41_001_m, "")
						*- Tipo lógico
						case y08.y08_012_c = "LOGIC"
							replace y41t.y41_003_n with iif(PLC_Content = ".T.", 1, 0) in Y41T
							VLU_ParContent = y41t.y41_003_n = 1
						*- Tipo data
						case y08.y08_012_c = "DATE"
							replace y41t.y41_004_d with ctod(PLC_Content) in Y41T
							VLU_ParContent = y41t.y41_004_d
						*- Tipo numérico
						case y08.y08_012_c = "NUMERIC"
							replace y41t.y41_002_b with val(PLC_Content) in Y41T
							VLU_ParContent = y41t.y41_002_b
						*- Tipo matriz
						case y08.y08_012_c = "ARRAY"
							replace y41t.y41_002_b with val(PLC_Content) in Y41T
						*- Tipo texto
						case y08.y08_012_c = "TEXT"
							replace y41t.y41_001_m with PLC_Content in Y41T
						*- Tipo multiplas opções
						case y08.y08_012_c = "OPTIONS"
							replace y41t.y41_002_b with val(PLC_Content) in Y41T
					endcase
					replace y41t.y41_006_c with this.FOC_CheckSum(alltrim(y41t.y41_001_m)), ;
							y41t.mycontrol with "1" in Y41T
					.FOL_CreateSQLString("Y41T", "Y41", empty(y41t.status), .T.)
					VLL_Return = .FOL_SaveRecordSQL("Y41T", "Y41", "UKEY")
					.FOL_CloseTable("Y41T")
					if VLL_Return
						VLL_Return = .FOL_CommitTrans("Y41")
						if !isnull(VLU_ParContent)
							VLN_Par = ascan(this.VOA_Sets, PLC_Y08Ukey, -1, -1, 1)
							if VLN_Par > 0
								VLN_Par = asubscript(this.VOA_Sets, VLN_Par, 1)
								this.VOA_Sets[VLN_Par, 2] = VLU_ParContent
							endif	
						endif
					else
						.FOL_RollBack("Y41")
					endif
				endif
			else
				*- Parâmetro não encontrado
				this.FON_Msg(1443)
			endif
			
			use in y08

		endwith
		return VLL_Return
	ENDPROC


*-- Retorna toda a hierarquia de um objeto
	PROCEDURE FOC_ParString
		*/----------------------------------------------------------------------------------------/*
		*/ Função Objeto	: FOC_ParString			                                              /*
		*/ Parâmetros		: PLO_Object : Referência do objeto 			          	    	  /*
		*/					  PLC_Separator : Indica o separador dos objetos, se vazio coloca _   /*
		*/ Descrição		: Retorna toda a hierarquia de um objeto                      		  /*
		*/ Retorno			: Cadeia hierarquica do objeto até antes do form                      /*
		*/ Grupo			: Objeto, hierarquia                                                  /*
		*/ Procura por		: Objeto                                                              /*
		*/ Última alteração	: 14/12/2001 														  /* 
		*/ Alterado por		: Denis  														      /*
		*/ Versão			: 2																	  /*
		*/----------------------------------------------------------------------------------------/*
				
		LPARAMETERS PLO_Object, PLC_Separator

		local VLC_Return, VLO_Parent, VLO_CurObj
		
		VLC_Return = upper(sys(1272, PLO_Object))

		if empty(PLC_Separator)
			PLC_Separator = "_"
		endif

		if PLC_Separator <> "."
			VLC_Return = strtran(VLC_Return, ".", PLC_Separator)
		endif

		return VLC_Return
	ENDPROC


*-- Retorna o tratamento do campo usr_note dos registros
	PROCEDURE FOC_UsrNote
		*/----------------------------------------------------------------------------------------/*
		*/ Função Objeto	: FOC_UsrNote			                                              /*
		*/ Parâmetros		: PLC_Note : Texto do campo sem tratamento		          	    	  /*
		*/          		: PLN_Option : Número da opção desejada de 1 a 5         	    	  /*
		*/ Descrição		: Retorna o tratamento do campo usr_note dos registros        		  /*
		*/ Retorno			: Campo de Observação tratado, sem as tags de separação               /*
		*/ Grupo			: STRING       						                                  /*
		*/ Procura por		: Comentário, usr_note	                                              /*
		*/ Última alteração	: 22/03/2001 														  /* 
		*/ Alterado por		: Denis  														      /*
		*/ Versão			: 1																	  /*
		*/----------------------------------------------------------------------------------------/*
				
		LPARAMETERS PLC_Note, PLN_Option

		local VLC_Value, VLC_Number, VLN_Counter
		
		VLC_Value  = nvl(PLC_Note, "")

		if empty(PLN_Option)

			for VLN_Counter = 2 to 5
				VLC_Number = alltrim(str(VLN_Counter))
				VLC_Value = strtran(VLC_Value, "<werp.option" + VLC_Number + ">", "")
				VLC_Value = strtran(VLC_Value, "</werp.option" + VLC_Number + ">", "")
			endfor

		else

			if PLN_Option = 1

				VLN_FOption	= at("<werp.option", VLC_Value)

				if VLN_FOption > 0
					VLC_Value = subs(VLC_Value, 1, VLN_FOption - 1)
				endif
			
			else
				VLC_Number 	= alltrim(str(PLN_Option))
				VLN_IOption	= at("<werp.option" + VLC_Number + ">", VLC_Value)
				VLN_FOption	= at("</werp.option" + VLC_Number + ">", VLC_Value)

				if VLN_IOption > 0
					VLC_Value = subs(VLC_Value, VLN_IOption + 14, VLN_FOption - VLN_IOption - 14)
				else
					VLC_Value = ""
				endif
			endif
		endif

		return VLC_Value
	ENDPROC

	PROCEDURE FOC_TagContent
		*/----------------------------------------------------------------------------------------/*
		*/ Função Objeto	: FOC_TagContent		                                              /*
		*/ Parâmetros		: PLC_String : Texto para ser tratado			          	    	  /*
		*/          		: PLC_Tag : Indica a tag a ser tratada no formato <tag></tag>    	  /*
		*/								obs.: passar apenas o texto da tag sem <, >, ou /		  /*
		*/					: PLN_Occur : Ocorrência da tag dentro da expressão					  /*
		*/ Descrição		: Retorna o conteúdo de uma tag dentro de uma string          		  /*
		*/ Grupo			: STRING       						                                  /*
		*/ Procura por		: Tag, String         	                                              /*
		*/ Última alteração	: 29/03/2001 														  /* 
		*/ Alterado por		: Denis  														      /*
		*/ Versão			: 1																	  /*
		*/----------------------------------------------------------------------------------------/*
				
		lparameters PLC_String, PLC_Tag, PLN_Occur

		local VLC_Value, VLC_TagStart, VLN_Start, VLC_TagEnd, VLN_End, VLN_Size, VLC_Return
		local VLN_RStart
		
		if empty(PLN_Occur)
			PLN_Occur = 1
		endif

		PLC_String  	= nvl(PLC_String, "")
		VLC_TagStart	= "<" + PLC_Tag + ">"
		*-VLN_RStart		= at(VLC_TagStart, PLC_String, PLN_Occur)
		*-VLN_Start		= VLN_RStart + len(PLC_Tag) + 2
		VLC_TagEnd		= "</" + PLC_Tag + ">"
		*-VLN_REnd		= at(VLC_TagEnd, PLC_String, PLN_Occur)
		*-VLN_End			= VLN_REnd - 1
		*-VLN_Size		= VLN_End - VLN_Start + 1
		VLC_Return		= strextract(PLC_String, VLC_TagStart, VLC_TagEnd, PLN_Occur, 1)
		
		*-if VLN_RStart > 0 and VLN_REnd > 0 and VLN_Size > 0
		*-	VLC_Return = subs(PLC_String, VLN_Start, VLN_Size)
		*-endif
			
		return VLC_Return
	ENDPROC


*/----------------------------------------------------------------------------------------
*/ Descrição   		:  Monta um array publico com o código e a resposta dos parâmetros 
*/ Retorno     		: .t.
*/----------------------------------------------------------------------------------------
PROCEDURE FOL_StartSets

*- Cursores para substituir os DBFS
this.FOL_ExecuteCursor("Y41_ALL", "Y41", 0)
select y41
index on y08_ukey tag y08_ukey

this.FOL_UseInPackage("y08", "y08")

select y08

VLN_Set = 0
scan
	VLC_TypeParameter 	= y08.y08_012_c
	if !seek(y08.ukey, "Y41", "y08_ukey")
		VLC_Props			= nvl(y08.y08_999_m, "")
		do case
			case VLC_TypeParameter = "CHARACTER"   && Caracter
				VLU_StdRet = ""

			case VLC_TypeParameter = "LOGIC"       && Lógico
				VLU_StdRet = .T.

			case VLC_TypeParameter = "DATE"        && Data
				VLU_StdRet = {}

			case VLC_TypeParameter = "NUMERIC"     && Numérico
				VLU_StdRet = 0
		endcase
		VLU_ParContent = this.FOU_ChkProperties(VLC_Props, "DEFAULT", 0, .T., "")

	else
		do case
			case VLC_TypeParameter = "CHARACTER"   && Caracter
				VLU_ParContent = nvl(y41.y41_001_m, "")

			case VLC_TypeParameter = "LOGIC"       && Lógico
				VLU_ParContent = y41.y41_003_n = 1

			case VLC_TypeParameter = "DATE"        && Data
				VLU_ParContent = y41.y41_004_d

			case VLC_TypeParameter = "NUMERIC"     && Numérico
				VLU_ParContent = y41.y41_002_b
		endcase
	endif

	VLN_Set = VLN_Set + 1
	dimension this.VOA_Sets[VLN_Set, 2]
	this.VOA_Sets[VLN_Set, 1] = y08.ukey
	this.VOA_Sets[VLN_Set, 2] = VLU_ParContent
endscan

use in Y41
use in Y08

return .t.


*/----------------------------------------------------------------------------------------/*
*/ Função Objeto	: FOU_AnalyzeForm                                                     /*
*/ Parâmetros		: PLC_FileLog 	- Arquivo de Log                                      /*
*/                  : PLC_FileResult- Arquivo de Resultado                                /*
*/ Descrição		: Analise do arquivo de LOG para apontar a performance da tela        /*
*/ Retorno			: Arquivo com as totalizações de ocorrências de chamadas a métodos,   /*
*/                  : eventos e funcões                                                   /*
*/----------------------------------------------------------------------------------------/*
*-- Analise do arquivo de LOG para apontar a performance da tela.
PROCEDURE FOU_AnalyzeForm
	LParameters PLC_FileLog, PLC_FileResult

	LOCAL VLC_String, VLL_Valid, VLN_Line, VLC_Intermed, VLC_Total

	VLC_Total = sys(2015)
	VLC_Intermed = sys(2015)

	*- Tabela que recebera os dados do LOG
	create cursor tablelog ;
		(duration n(7,3), class c(30), 	procedure c(60), line i, file c(60), number n(6))

	*- Tabela com as occorrencias dos metodos, eventos e funções
	create table (VLC_Intermed) ;
		(duration n(7,3), class c(30), 	procedure c(60), line i, file c(60), number n(6))
	select (VLC_Intermed)
	use (VLC_Intermed) alias intermed
	index on class+procedure+file tag filter
	set order to filter

	*- Recebe arquivo de LOG
	select tablelog
	append from (PLC_FileLog) type delimiters
	replace all number with 1

	select class, procedure, file, sum(duration) as duration from tablelog group by class,procedure,file into table (VLC_Total)
	select (VLC_Total)
	use (VLC_Total) alias total
	index on class+procedure+file tag class

	*- Varre a tabela de LOG e separa as ocorrências
	select tablelog
	VLC_String = "[]"
	scan
		VLL_Valid = .f.
		VLN_Line  = tablelog.line
		VLC_String = class+procedure+file
		scatter memvar memo
		select intermed
	*--	Verifica se a ocorrencia esta cadastrada
		if seek(VLC_String)
	*---	Valida somente se o ocorrência for igual a primeira cadastrada
			if VLN_Line = intermed.line
				VLL_Valid = .t.
			endif
		else
	*----	Registra a ocorrência se não existir 
			VLL_Valid = .t.
		endif

		if VLL_Valid
			append blank
			gather memvar memo
		endif

		select tablelog
	endscan

	*- Totaliza as ocorrencias
	select class, procedure, file , sum(number) as occurs, duration  from intermed group by class, procedure into table (PLC_FileResult)
	select (PLC_FileResult)
	scan
		if seek(class+procedure+file, "total", "class")
			select (PLC_FileResult)
			replace duration with total.duration
		endif
	endscan

	select (PLC_FileResult)
	index on occurs tag occurs
	index on duration tag duration
	index on procedure tag procedure

	*- Elimina os arquivos temporários
	this.FOL_CloseTable("tablelog")
	this.FOL_CloseTable("intermed")
	this.FOL_CloseTable("total")
	erase &VLC_Intermed..*
	erase &VLC_Total..*

	return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros       : PLN_Lin		: Numero da linha da célula
*/                    PLC_Col		: Nome (Letra) da coluna da célula
*/ Descrição        : Retorna o valor de um campo do Perfil de fórmulas
*/ Retorno          : Caracter  
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_GetFieldValue
	LPARAMETERS PLN_Lin, PLC_Col
	LOCAL VLC_Return

	select TMP_FlexGrid
	if (PLN_Lin > 0 and PLN_Lin <= reccount("TMP_FlexGrid")) and type('PLC_Col') <> "U"
		go PLN_Lin
		VLC_Return = evaluate("TMP_FlexGrid." + PLC_Col)
	else
		VLC_Return = ""
	endif
	return VLC_Return
ENDPROC


*/----------------------------------------------------------------------------------------
*/ Parametros		: PLC_String : Comando SQL para ser alterado
*/ Descricao		: Deixa o comando SQL pronto para ser executado
*/ Retorno			: O comando pronto para ser executado
*/----------------------------------------------------------------------------------------
procedure FOC_ChangeString

LParameters PLC_String

LOCAL VLN_i, VLN_Pos, VLC_Aux, VLC_TableName, VLC_Subst, VLC_KeyWords, VLC_NextWord, VLN_StrSize, ;
	VLN_InitNextWord, VLC_Command, VLL_Initied, VLC_RefTableName, VLN_LenToStuff, VLC_AliasToUse, ;
	VLN_LenNextWord, VLL_NoLockExpr, VLL_From, VLC_KeyWords2

VLC_KeyWords = "WHERE//JOIN//LEFT//OUTER//CROSS//IN//RIGHT//AS//INNER//,//ON//(NOLOCK)"
VLC_KeyWords2 = VLC_KeyWords+"//ORDER//GROUP"

VLL_From = at("FROM", PLC_String) > 0 and upper(substr(alltrim(PLC_String), 1, 6)) = "SELECT"
VLC_UpperStr = upper(PLC_String)
VLL_IsSelect = (at("DELETE", VLC_UpperStr) + at("UPDATE", VLC_UpperStr) + at("INSERT", VLC_UpperStr) <= 0)

*- altera o nome das tabelas colocando os seus respectivos database names
do while .T.
	VLN_i = at("STAR_DATA@", PLC_String)
	if VLN_i = 0
		exit
	endif

	VLN_LenToStuff = 10
	VLC_Command = substr(PLC_String, VLN_i + 10)
	VLN_Pos = 1
	VLN_StrSize = len(VLC_Command)
	VLL_NoLockExpr = .f.

	*-	Acha a primeira palavra com uma letra a mais (é o nome da tabela sempre)
	do while VLN_Pos <= VLN_StrSize
		VLC_Aux = subs(VLC_Command, VLN_Pos, 1)
		if asc(VLC_Aux) < 48 
			exit
		endif
		VLN_Pos = VLN_Pos + 1
		VLN_LenToStuff = VLN_LenToStuff + 1
	enddo

	*-	Preciso da próxima palavra, então estou pegando a próxima letra
	VLN_InitNextWord = VLN_Pos

	*-	Tiramos a última letra
	VLC_TableName = substr(VLC_Command, 1, VLN_Pos - 1)
	VLL_Initied = .F.
	VLN_LenNextWord = 0

	VLC_RefTableName = upper(alltrim(VLC_TableName))
	
	VLL_FilteredTable = .F.
	VLC_FixedFilter = ""
	for VLN_Counter = 1 to this.VON_TableStdFilterCount
		if at("//" + VLC_RefTableName + "//", this.VOA_TableStdFilter[VLN_Counter, 2]) > 0
			VLL_FilteredTable = .t.
			VLC_FixedFilter = this.VOA_TableStdFilter[VLN_Counter, 1]
			exit
		endif
	endfor

	VLL_Initied = asc(substr(VLC_Command, VLN_Pos, 1)) <> 32

	*-	Agora preciso da próxima palavra
	do while VLN_Pos <= VLN_StrSize
		VLC_Aux = substr(VLC_Command, VLN_Pos, 1)
		VLN_CharNo = asc(VLC_Aux)

		if VLL_Initied
			if VLN_CharNo < 48 
				VLN_LenNextWord = VLN_LenNextWord + 1
				exit
			endif
		else
			if VLN_CharNo <> 32	&& Espaço, só começo a contar depois dos espaços
				VLL_Initied = .T.
			endif
		endif

		VLN_Pos = VLN_Pos + 1
		VLN_LenNextWord = VLN_LenNextWord + 1
	enddo

	VLC_NextWord = upper(ltrim(subs(VLC_Command, VLN_InitNextWord, VLN_Pos - VLN_InitNextWord + 1)))

	if at(alltrim(VLC_NextWord), VLC_KeyWords) > 0 and VLL_IsSelect
		*- Não existe alias, usarei o nome da tabela como alias padrão
		VLC_AliasToUse = VLC_TableName

		if at("(NOLOCK)", VLC_NextWord) > 0
			VLL_NoLockExpr = .T.
			VLN_LenToStuff = VLN_LenToStuff + VLN_LenNextWord
		endif
	else
		VLN_LenToStuff = at("STAR_DATA@" + upper(VLC_Command), upper(PLC_String)) + VLN_Pos - VLN_I + 10
		*- Preciso usar o alias definido na select
		VLC_AliasToUse = VLC_NextWord
		if (alltrim(upper(VLC_NextWord)) $ VLC_KeyWords2) and VLL_From
			VLC_AliasToUse = " " + VLC_TableName + " " + VLC_AliasToUse
		endif

		*- Se o BD for MS-SQL Server e a próxima palavra não foi (NOLOCK), preciso verificar se a palavra seguinte é (NOLOCK)
		if this.VON_DBType = 1 && Sql Server
			VLN_LenNextWord = 0
			VLL_Initied = .f.
			*-	Agora preciso da próxima palavra
			do while VLN_Pos <= VLN_StrSize
				VLC_Aux = substr(VLC_Command, VLN_Pos, 1)
				VLN_CharNo = asc(VLC_Aux)

				if VLL_Initied
					if VLN_CharNo < 48 
						VLN_LenNextWord = VLN_LenNextWord + 1
						exit
					endif
				else
					if VLN_CharNo <> 32	&& Espaço, só começo a contar depois dos espaços
						VLL_Initied = .t.
					endif
				endif

				VLN_Pos = VLN_Pos + 1
				VLN_LenNextWord = VLN_LenNextWord + 1
			enddo

			VLC_NextWord = upper(alltrim(substr(VLC_Command, VLN_InitNextWord, VLN_Pos - VLN_InitNextWord + 1)))

			if at("(NOLOCK)", VLC_NextWord) > 0
				VLL_NoLockExpr = .t.
				VLN_LenToStuff = VLN_LenToStuff + VLN_LenNextWord
			endif
		endif
	endif

	VLC_Subst = ""

	VLC_RefTableIdent = this.FOC_TableIdent(VLC_RefTableName) && Identificador da tabela
	VLC_Subst = VLC_RefTableIdent + VLC_TableName

	if VLL_FilteredTable and VLL_IsSelect
		VLC_Subst = "(SELECT * FROM " + VLC_Subst + iif(VLL_NoLockExpr, " (NOLOCK) ", "") +	VLC_FixedFilter + ") " + VLC_AliasToUse
	else
		VLC_Subst = VLC_Subst + " " + VLC_AliasToUse + iif(VLL_NoLockExpr, " (NOLOCK) ", "")
	endif

	PLC_String = stuff(PLC_String, VLN_i, VLN_LenToStuff, VLC_Subst)
enddo

*- Preciso trocar para ter certeza que estará maiusculo, o tratamento acontecerá na FOL_SQLEXEC
PLC_String = strtran(PLC_String, "vpa_seek[", "VPA_SEEK[", -1, -1, 1)

return PLC_String

ENDPROC

*-----------------------------------------------------------------------------------------------------------------------
*- Carrega uma array com os identificadores das tabelas do sistema para a substituição do STAR_DATA@
*- Parâmetros - PLC_TabId : Id da Tabela desejada
*- Retorno   : .T. se conseguir carregar, .F. c.c.
*- Denis Carrizo em 11/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_TableIdent
lparameters PLC_TabId

local VLN_Counter

VLC_Return = ""
with this
	if .VON_TableIdentCount > 1
		for VLN_Counter = 1 to .VON_TableIdentCount
			if "//" + PLC_TabId + "//" $ .VOA_TableIdent[VLN_Counter, 2]
				VLC_Return = .VOA_TableIdent[VLN_Counter, 1]
				exit
			endif
		endfor
	else
		VLC_Return = .VOA_TableIdent[1, 1]
		if empty(VLC_Return)
			VLC_Return = ""
		endif
	endif
endwith

return VLC_Return

endproc

*/----------------------------------------------------------------------------------------/*
*/ Parâmetros  		: PLC_CiaUkey : Ukey da Empresa
*/					  PLC_UsrPsw : Password do usuário
*/ Descrição   		: Tenta efetuar o Login na empresa Default (_STAR) atravez de um usuario que
*/					  possua uma "conexao segura". Associando o usuario logado no sistema 
*/					  operacional a um usuario do sistema. apos acessar este usurio esta funcao
*/					  procura a empresa padra do mesmo.
*/					  
*/ Retorno          : -1: Dados Incorretos, 1: Login Ok
*/ Grupo       		: GENERAL
*/ Procura por 		: Cia, Usuário, senha, login
*/ Última alteração	: 16/04/2003
*/ Alterado por 	: wlc
*/ Versão 			: 1
*/----------------------------------------------------------------------------------------/*
procedure FON_TrustedUser
Lparameters PLC_CiaUkey, PLC_UserLogin, PLL_NoLogin, PRA_CiaUserInfo

	Local VLN_Size, VLR_Buffer, VLN_return, VLL_CIAFound 
	
	If Empty(PLC_CiaUkey)
		PLC_CiaUkey = "STAR_"	&& Assume que esta e a empresa padrao do Sistema
	Endif

	*	Caso eu nao informe o usuario (O que e comum nesta funcao) procuro o nome do
	*	usuario logado no sistema operacional.
	If Empty(PLC_UserLogin)

		PLC_UserLogin = ""
	    VLN_Size = 250 
	    VLR_Buffer = SPACE(VLN_Size ) 

		 DECLARE INTEGER GetUserName IN advapi32; 
		        STRING  @ lpBuffer, INTEGER @ nSize 


	    IF GetUserName (@VLR_Buffer, @VLN_Size ) > 0 
	        VLR_Buffer= STRTRAN (ALLTRIM(SUBSTR (VLR_Buffer, 1, VLN_Size )), Chr(0),"") 
	        PLC_UserLogin  = Alltrim(VLR_Buffer)
	    ENDIF 
	    
	    CLEAR DLLS  "GetUserName"
	endif

	VLN_return = -1
	VLL_CIAFound =  this.FOL_FindExpression("CIA_UKEY", "CIA", PLC_CiaUkey)

	if VLL_CIAFound
		*- Carrega os identificadores de tabelas do sistema

		this.VON_LastSql = 0
		dimension this.VOA_LastSql[1, 2]
		this.VON_LastSqlOptions = 0
		dimension this.VOA_LastSqlOptions[1, 2]

		*- Carrega os identificadores de tabelas do sistema
		this.FOL_LoadTblId(cia.ukey)
			
		If this.FOL_FindExpression("USR_CODE", "USR", padr(PLC_UserLogin, 20))
			If !Empty(usr.usr_021_n) And !Empty(usr.usr_024_c)

				VLC_Control= this.FOC_UserChangeControl(usr.usr_016_m, usr.ukey, .t.)
				VLC_PassWord = alltrim(subs(VLC_Control, 17))
				VLC_Login    = Alltrim(usr.Usr_CODE)
				VLC_UkeyCia = Substr(usr.usr_024_c,1, Len(usr.ukey) )
				
				If Empty(PLL_NoLogin)
					VLN_return = this.FON_SetCiaUsr(VLC_UkeyCia , PLC_UserLogin, VLC_PassWord)
				Else
					VLN_return = 2
					Dimension PRA_CiaUserInfo(5)
					PRA_CiaUserInfo(1) = cia.cia_001_C
					PRA_CiaUserInfo(2) = cia.ukey
					PRA_CiaUserInfo(3) = Alltrim(usr.Usr_CODE)
					PRA_CiaUserInfo(4) = Alltrim(usr.ukey)
					PRA_CiaUserInfo(5) = VLC_PassWord
				endif
			Endif
			
		Endif

	endif
	
	*- Zera o array de comandos SQL porque o caminho das tabelas dos selects executados acima
	*- pode mudar caso o login não tenha sido feito.
	this.VON_LastSql = 0
	dimension this.VOA_LastSql[1, 2]
	this.VOA_LastSql = .f.
	this.VON_LastSqlOptions = 0
	dimension this.VOA_LastSqlOptions[1, 2]
	this.VOA_LastSqlOptions = .f.

	Return VLN_return 
endproc

*/----------------------------------------------------------------------------------------/*
*/ Parâmetros  		: PLC_CiaUkey : Ukey da Empresa
*/                    PLC_UsrLogin : Login ID do usuário
*/					  PLC_UsrPsw : Password do usuário
*/ Descrição   		: Efetua o login do usuário na empresa desejada, inicializando o ambiente
*/ Retorno          : -1: Dados Incorretos, 1: Login Ok
*/ Grupo       		: GENERAL
*/ Procura por 		: Cia, Usuário, senha, login
*/ Última alteração	: 08/02/2002
*/ Alterado por 	: Denis Carrizo
*/ Versão 			: 3
*/----------------------------------------------------------------------------------------/*
procedure FON_SetCiaUsr
lparameters PLC_CiaUkey, PLC_UsrLogin, PLC_UsrPsw

LOCAL VLC_OldLang, VLL_CIAFound, VLL_USRFound, VLN_MaxWidth, VLN_MinWidth, VLC_Control, VLC_LastCiaUkey
local VLN_Return, VLN_ChangePasswd, VLC_SqlCommand, VLN_RowsAffect, VLL_ChangePsw, VLL_StepVerify, VLN_UsrLicClass

*- Inicializo com o flag de dados incorretos
VLN_Return = -1
this.VOL_LoggedOn = .f.
VLN_UsrLicClass = 0

if !this.VOL_AutomatedLogin
	this.FOL_CloseAllTable()
endif

VLC_LastCiaUkey = this.VOC_CompanyCode

VLL_CIAFound = .f.
*- Tento achar a cia pela ukey
if len(PLC_CiaUkey) = 5
	VLL_CIAFound = this.FOL_FindExpression("CIA_UKEY", "CIA", PLC_CiaUkey)
endif

*- Se não achou pela ukey tento procurar pela descrição
if !VLL_CIAFound
	VLL_CIAFound = this.FOL_FindExpression("CIA_001_C", "CIA", PLC_CiaUkey)
endif

this.VON_LastSql = 0
dimension this.VOA_LastSql[1, 2]
this.VON_LastSqlOptions = 0
dimension this.VOA_LastSqlOptions[1, 2]


dimension this.voa_LastA37[1, 5]
store .f. to this.voa_LastA37

if VLL_CIAFound
	*- Carrega os identificadores de tabelas do sistema
	this.FOL_LoadTblId(cia.ukey)
endif

VLL_USRFound = this.FOL_FindExpression("USR_CODE", "USR", padr(PLC_UsrLogin, 20))

if !VLL_USRFound and VLL_CIAFound and len(PLC_UsrLogin) = 5 and this.VOL_AutomatedLogin
	VLL_USRFound = this.FOL_FindExpression("USR_UKEY", "USR", PLC_UsrLogin)
	if VLL_USRFound and empty(usr.usr_034_n)
		VLL_StepVerify = .t.
	endif
endif

if VLL_StepVerify
	VLN_Return = 1
else
	if VLL_CIAFound and VLL_USRFound
		*- Verifica se possui acesso à empresa se não for administrador
		VLC_Aux = this.FOC_UserChangeControl(usr.usr_018_m, usr.ukey, .t.)
		if usr.ukey == "STAR_" or at("__ALL;", VLC_Aux) > 0 or at(cia.ukey + ";", VLC_Aux) > 0

			VLC_Control = this.FOC_UserChangeControl(usr.usr_016_m, usr.ukey, .t.)

			if (alltrim(subs(VLC_Control, 17)) == alltrim(PLC_UsrPsw))
				VLN_Return = 1
				VLL_ChangePsw = .f.

				*- Critérios para troca de senha
				if subs(VLC_Control, 1, 1) <> "S"
					if empty(subs(VLC_Control, 17)) and subs(VLC_Control, 5, 1) <> "S"
						VLL_ChangePsw = .t.
					else
						if val(subs(VLC_Control, 6, 2)) > len(alltrim(subs(VLC_Control, 17)))
							VLL_ChangePsw = .t.
						else
							if subs(VLC_Control, 8, 1) = "S"
								VLL_ChangePsw = .t.
							else
								VLN_I = val(subs(VLC_Control, 2, 3))
								if VLN_I > 0
									if this.FOD_Stod(subs(VLC_Control, 9, 8)) + VLN_I < date()
										VLL_ChangePsw = .t.
									endif
								endif
							endif
						endif
					endif
				endif

				if VLL_ChangePsw
					*- Troca de senha
					this.FOL_DoNormalForm("PASSNEW", VLC_Control)
					this.VOU_FormReturn = subs(VLC_Control, 1, 7) + "N" + dtos(date()) + subs(this.VOU_FormReturn, 17)
					replace usr.usr_016_m with this.FOC_UserChangeControl(this.VOU_FormReturn, usr.ukey, .t.)
					with this
						.FOL_TableUpdate("USR")
						if .FOL_FindExpression("USR_UKEY", "USR_TMP", usr.ukey, .f., .t.)
							select USR_TMP
							replace mycontrol with "1",;
									usr_016_m with usr.usr_016_m in USR_TMP

							.FOL_CreateSqlString("USR_TMP", "USR", .f., .f.)
							if .FOL_BeginTransaction("USR")
								VLL_Return = .FOL_SaveRecordSql("USR_TMP", "USR", "UKEY")
								if VLL_Return
									VLL_Return = .FOL_CommitTrans("USR")
								else
									.FOL_RollBack("USR")
								endif
							endif
							use in usr_tmp
						endif
					endwith
				endif

			endif
		endif
	endif
endif

this.VOC_CompanyCode	= padr(nvl(cia.Ukey, ""), 5)
this.VOC_UserCode   	= padr(nvl(usr.ukey, ""), 5)

*- Se o login estiver em progresso
if VLN_Return = 1
	VLN_Return = this.VOO_Security.FON_CheckLogin(usr.usr_020_n, usr.ukey)
endif

*- Se a condição de login estiver OK
if VLN_Return = 1

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("reset")
		VGO_ToolBar.FOL_ProgressBar("setmax", 4, "Logging On")

		*- Seta o agente
		if usr.array_236 > 1 && 1=NENHUM

			if !this.VOL_ServiceMode and file(getenv("WINDIR") + "\MSAGENT\AGENTCTL.DLL")	&&		*-	inicia a classe de agents
				if vartype(this.VOO_Agent) <> "O" 
					this.VOO_Agent = createobject("CGO_Agent")
				endif
			endif

			if !isnull(this.VOO_Agent)
				this.VON_ActiveChar = usr.array_236 - 1

				do case
					Case this.VON_ActiveChar = 1
						VLC_AuxName= "GENIE"
						
					Case this.VON_ActiveChar = 2
						VLC_AuxName= "PEEDY"
						
					Case this.VON_ActiveChar = 3
						VLC_AuxName= "MERLIN"
						
					Case this.VON_ActiveChar = 4
						VLC_AuxName= "ROBBY"
				endcase

				If this.VOO_Agent.FOL_loadAgent(VLC_AuxName, VLC_AuxName + ".acs") 
				    this.VOO_Agent.FOL_setActiveAgent(VLC_AuxName)
				endif
			endif
		else
			this.VOO_Agent = .NULL.			
		endif
	endif

	this.VOC_CompanyName    		= alltrim( nvl(cia.cia_001_c, ""))
	this.VOC_CompanyCgc     		= alltrim( nvl(cia.cia_006_c, ""))
	this.VOC_CompanyIE      		= alltrim( nvl(cia.cia_007_c, ""))
	this.VOC_CompanyAddress 		= alltrim( nvl(cia.cia_003_c, ""))
	this.VOC_CompanyNumber 			= alltrim( nvl(cia.cia_067_c, ""))
	this.VOC_CompanyNeighborhood	= alltrim( nvl(cia.cia_002_c, ""))
	this.VOC_CompanyZipCode			= alltrim( nvl(cia.cia_004_c, ""))
	this.VOC_CompanyJC				= alltrim( nvl(cia.cia_009_c, ""))
	
	* Verifico se existe o logotipo cadastrado no CIA.
	this.VOC_CompanyLogo 			= alltrim( nvl(cia.cia_046_m, ""))
	if !file(this.VOC_CompanyLogo)
		this.VOC_CompanyLogo = ""
	endif
	this.VOC_CompanyPhone 			= alltrim( nvl(cia.cia_018_c, ""))
	this.VOC_CompanyDDDPhone 		= alltrim( nvl(cia.cia_017_c, ""))
	this.VOC_CompanyComplement 		= alltrim( nvl(cia.cia_068_c, ""))

	this.VOC_GroupCode 				= nvl(usr.grp_ukey, "")
	this.VOC_UserName  				= alltrim( nvl(usr.usr_001_C, ""))
	this.VOC_UserBorn  				= nvl(usr.usr_019_d, {})
	this.VOL_FillCurrency			= !empty(nvl(usr.usr_017_n, 0))


	*-Preencho com o diretório padrão do usuário cadastrado em integrantes
	this.FOL_FindExpression("C04_USR_UKEY1", "C04_USR", usr.ukey)
	this.VOC_UserDirectory = nvl(c04_usr.c04_012_m, "")
	this.FOL_CloseTable("C04_USR")

	if !empty(this.VOC_GroupCode) and ;
		this.FOL_FindExpression("GRP_UKEY", "GRP", usr.grp_ukey)
		this.VOC_GroupName = alltrim( nvl(grp.grp_001_C, "") )
	endif

	if !empty(nvl(usr.t77_ukey, "")) and this.FOL_FindExpression("T77_UKEY", "T77", usr.t77_ukey)
		this.VOC_UsrT77Ukey = t77.ukey
		this.VOC_UsrMsgAccount = alltrim(t77.t77_001_C)
		this.VOC_ReplyAddress = this.VOC_UsrMsgAccount
		if !empty(t77.t77_006_n) and !empty(t77.t77_015_c)
			this.VOC_ReplyAddress = alltrim(t77.t77_015_c)
		endif
	else
		this.VOC_UsrT77Ukey = .null.
		this.VOC_UsrMsgAccount = ""
		this.VOC_ReplyAddress = ""
	endif

	if used("t77")
		use in t77
	endif

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
	endif

	*- Muda o idioma
	VLC_OldLang = this.VOC_USER_Language
	do case
		case alltrim(usr.usr_025_c) = "055"   && Português
			this.VOC_USER_Language = "055"
		case alltrim(usr.usr_025_c) = "001"   && Inglês
			this.VOC_USER_Language = "001"
		case alltrim(usr.usr_025_c) = "034"   && Espanhol
			this.VOC_USER_Language = "034"
		case alltrim(usr.usr_025_c) = "033"   && Francês
			this.VOC_USER_Language = "033"
		case alltrim(usr.usr_025_c) = "049"   && Alemão
			this.VOC_USER_Language = "049"
		case alltrim(usr.usr_025_c) = "039"   && Italiano
			this.VOC_USER_Language = "039"
		otherwise
			this.VOC_USER_Language = "055"
	endcase

	if VLC_OldLang <> this.VOC_USER_Language
		this.FOL_InitArray()
	endif

	*- Verifico se houve troca de empresa para uma reinicialização mais rápida
	if !VLC_LastCiaUkey == this.VOC_CompanyCode
		this.FOU_MakePicts()		&&		*- inicia as picts do sistema
		this.FOL_StartSets()		&&		*- Joga os parâmetros em uma matriz publica
	endif

	this.VOC_UserSex   				= this.VOA_022[max(1,usr.Array_022)]
	this.VOC_UserCivil 				= this.VOA_033[max(1,usr.Array_033)]

	set status bar off
	set message to ""

	this.VOC_UserCia = this.FOC_UserChangeControl(usr.usr_018_m, usr.ukey, .T.)
	this.VOC_UserCia = substr(alltrim(this.VOC_UserCia), 1, len(alltrim(this.VOC_UserCia)) - 1)
	this.VOC_UserCia = "'" + strtran(this.VOC_UserCia, ";", "', '") + "'"

	if("__ALL" $ this.VOC_UserCia)
		this.FOL_ExecuteCursor("CIA_ALL", "CIAUSER", 0)
	else
		this.FOL_SetParameter(1, this.FOC_ClearString(this.VOC_UserCia, ",;'", ""))
		this.FOL_ExecuteCursor("CIA_USER", "CIAUSER", 1)
	endif

	VLN_i = 0

	select ciauser
	scan
		VLN_i = VLN_i + 1
		dimension this.VOA_CiaCodes[VLN_i, 2]
		this.VOA_CiaCodes[VLN_i, 1] = ciauser.ukey
		this.VOA_CiaCodes[VLN_i, 2] = ciauser.cia_001_c
	endscan
	use in ciauser

	with this
		.VOL_LimitCare = !empty(usr.usr_030_n)
		.VON_MaxRecordsAtTime = usr.usr_032_n
		.VON_MaxRecordsForWarning = usr.usr_031_n
	endwith

	if !this.VOL_ServiceMode
		*- Pano de Fundo
		if !empty(nvl(usr.usr_012_m, "")) .and. file(nvl(usr.usr_012_m, ""))
			_screen.picture = usr.usr_012_m
		else
			_screen.picture = ""
		endif

		*- Foto do Usuario
		VLC_Type = type("_screen.usrFOTO")
		if !empty(usr.usr_015_n) .and. file(nvl(usr.usr_010_m, ""))
			if VLC_Type !=  "O"
				_screen.addobject("usrFOTO", "cgo_usrfoto")
				_screen.usrFOTO.top  = 10
				_screen.usrFOTO.left = _screen.width - _screen.usrFOTO.width - 10
			endif
			VLC_foto = usr.usr_010_m
			_screen.usrFOTO.picture = VLC_Foto
			_screen.usrFOTO.visible = .T.
		else
			if VLC_Type =  "O"
				_screen.usrFOTO.visible = .F.
			endif
		endif

		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
	endif

*--	Indica se subdivisão de áreas está ativa ou não.
	this.VOL_SubMenus = this.FOU_ChkSets("STARPAR-L-00277")

*--	Indica se sempre fecha a conexão com o Banco de Dados no término da transação
	this.VOL_CloseConnection	= !this.FOU_ChkSets("STARPAR-L-00162")

*--	Indica se mostra o ícone de chave unica nas telas 
	this.VOL_ShowUniqueKey = this.FOU_ChkSets("STARPAR-L-00161")

*--	Cor das letras desabilitadas igual a das letras habilitadas 
	this.VOL_ChangeForeColor = this.FOU_ChkSets("STARPAR-L-00163")

*--	Indica se o sistema trabalha com Grade
	this.VOL_WorkWithGrade = this.FOU_ChkSets("STARPAR-L-00053")

*--	Indica se o sistema trabalha com Rastreabilidade
	this.VOL_WorkWithRaster = this.FOU_ChkSets("STARPAR-L-00054")

*-- (Projetos) Indica que o documento só poder ser visto, alterado, editado ou apagado
*--	pelo responsavel pelo projeto e que somente os usuários que fazem parte de um
*--	grupo de responsáveis poderá criar um novo documento.
	this.VOL_JustResponsible = this.FOU_ChkSets("STARPAR-L-00172")

*-- Indica se a empresa trabalha com o controle de apropriação de despesas do módulo de projetos.
	this.VOL_WorkWithProjects = this.FOU_ChkSets("STARPAR-L-00181")

*--	Permite editar documentos gerados por integração.
	this.VON_EditDocument = this.FOU_ChkSets("STARPAR-N-00204")

	if !empty(Cia.A24_Ukey)
		if this.FOL_FindExpression("A24_UKEY", "A24", cia.a24_ukey)
			this.VOC_CompanyCity     = Cia.A24_Ukey
			this.VOC_CompanyCityN    = A24.a24_001_c
		endif
	endif

	if !empty(cia.A23_Ukey)
		if this.FOL_FindExpression("A23_UKEY", "A23", cia.a23_ukey)
			this.VOC_CompanyState    = Cia.A23_Ukey
			this.VOC_CompanyStateN   = A23.a23_001_c
			this.VOC_CompanyUF		= A23.a23_002_c
		endif
	endif

	if !empty(Cia.A22_Ukey)
		if this.FOL_FindExpression("A22_UKEY", "A22", cia.a22_ukey)
			this.VOC_CompanyCountry  = Cia.A22_Ukey
			this.VOC_CompanyCountryN = A22.a22_001_c
		endif
	endif

	if !empty(cia.a36_ukey)
		if this.FOL_FindExpression("A36_UKEY", "A36", cia.a36_ukey)
			this.VOC_Currency = a36.a36_code
		endif
	else
		if this.FOL_FindExpression("A36_A36_004_N", "A36", 1)
			this.VOC_Currency = a36.a36_code
		endif
	endif

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
	endif

	with this
		VLC_Parameter = .FOU_ChkSets("STARPAR-C-00002")	&&		*-- Armazena os SM. para utilização na contabilidade apuração de saldos
		dimension .VOA_B_Currency[3]
		.VOA_B_Currency[1] = substr(VLC_Parameter,  1, 5)
		.VOA_B_Currency[2] = substr(VLC_Parameter,  7, 5)
		.VOA_B_Currency[3] = substr(VLC_Parameter, 13, 5)

		.VOC_CepStarDirectory = nvl( .FOU_ChkSets("STARPAR-C-00189"), "")  &&		*--	Armazena o Diretorio de instalacao do PlugIn CEPSTAR
		.VON_Accurracy = .FOU_ChkSets("STARPAR-N-00089")							&&		*--	Quantidade de casas decimais
		.VON_RoundType = .FOU_ChkSets("STARPAR-N-00088") - 1						&&		*-- Tipo de arredondamento padrão
		.VOC_WebServer = alltrim(.FOU_ChkSets("STARPAR-C-00091"))					&&		*--	Nome do server
		.VOC_ClipartDirectory = alltrim(.FOU_ChkSets("STARPAR-C-00092"))			&&		*--	Diretório padrão do Clipart para Web
		.VOC_BannerDirectory = alltrim(.FOU_ChkSets("STARPAR-C-00093"))				&&		*--	Diretório padrão do Banner para Web

		VLC_Parameter = .FOU_ChkSets("STARPAR-C-00052")	&&		*--	OBS : Vide array 324	&&		*--	Verifica os módulos que utilizam ou não Crd I e Crd II

		VLN_Content = val(substr(VLC_Parameter, 1, 1))	&&		*--	Contabilidade
		.VOL_Crd1B = VLN_Content > 0
		.VOL_Crd2B = .VOL_Crd1B and VLN_Content = 2

		VLN_Content = val(substr(VLC_Parameter, 2, 1))	&&		*--	Estoque
		.VOL_Crd1D = VLN_Content > 0
		.VOL_Crd2D = .VOL_Crd1D and VLN_Content = 2

		VLN_Content = val(substr(VLC_Parameter, 3, 1))	&&		*--	Folha de Pagamento
		.VOL_Crd1M = VLN_Content > 0
		.VOL_Crd2M = .VOL_Crd1M and VLN_Content = 2

		VLN_Content = val(substr(VLC_Parameter, 4, 1))	&&		*--	Financeiro
		.VOL_Crd1F = VLN_Content > 0
		.VOL_Crd2F = .VOL_Crd1F and VLN_Content = 2

		VLN_Content = val(substr(VLC_Parameter, 5, 1))	&&		*--	PCP
		.VOL_Crd1K = VLN_Content > 0
		.VOL_Crd2K = .VOL_Crd1K and VLN_Content = 2

		.VOD_BClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00007"), {})
		.VOD_DClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00143"), {})
		.VOD_EClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00014"), {})
		.VOD_FClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00154"), {})
		.VOD_GClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00114"), {})
		.VOD_HClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00156"), {})
		.VOD_IClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00036"), {})
		.VOD_JClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00140"), {})
		.VOD_KClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00229"), {})
		.VOD_MClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00074"), {})
		.VOD_NClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00047"), {})
		.VOD_OClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00050"), {})
		.VOD_RClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00167"), {})
		.VOD_VClosingDate = nvl(.FOU_ChkSets("STARPAR-D-00270"), {})
		.VOL_RateioContabilObrigatorio = .FOU_ChkSets("STARPAR-L-00273")
		.VOL_RateioFinanceiroObrigatorio = .FOU_ChkSets("STARPAR-L-00274")
		.VOL_RateioEstoqueObrigatorio = .FOU_ChkSets("STARPAR-L-00275")
		.VOL_RateioFolhaObrigatorio = .FOU_ChkSets("STARPAR-L-00252")
		.VOL_JustResponsible = .FOU_ChkSets("STARPAR-C-00172")
		
		*- Indica se a conta analítica deverá ser gerada parar (Cliente, fornecedor ou ambos).
		.VON_Array445 = this.FOU_Chksets("STARPAR-N-00215")

		*- Ukey da conta sintética do cliente.
		.VOC_A03B11UkeySynthetic = this.FOU_Chksets("STARPAR-C-00216")

		*- Ukey da conta sintética do fornecedor.
		.VOC_A08B11UkeySynthetic = this.FOU_Chksets("STARPAR-C-00217")
		 
		*- Filtro do Registro de saida 
		.VON_RepoOutflowRecord = this.FOU_Chksets("STARPAR-N-00230")

		*- Filtro do Registro de Entrada 
		.VON_RepoInflowRecord = this.FOU_Chksets("STARPAR-N-00231")

		*- Indica se deve ser utilizado o valor contábil da capa para impressão dos relatórios em livros
		.VOL_AccountingCover = this.FOU_ChkSets("STARPAR-L-00256")

		if !this.VOL_ServiceMode
			VGO_ToolBar.FOL_ActionInToolBar("", "",  "setcaption", .t. )
			VGO_ToolBar.FOL_ActionInButton( "" , "", "tooltiptext", .t. )
		endif

		VLC_AdditionExFlds = .FOU_ChkSets("STARPAR-C-00255")
		if !empty(VLC_AdditionExFlds)
			VLC_AdditionExFlds = upper(alltrim(chrtran(VLC_AdditionExFlds, ",", " ")))
			.VOC_SQLExcludeFields = .VOC_SQLExcludeFields + VLC_AdditionExFlds + " "
		endif
		
		*- Inicia a configuração de cores dos documentos HTML
		VGO_WebPages.FOL_InitColors(this.FOU_Chksets("STARPAR-C-00182"))
		
		*- Função que monta a string de acesso do sistema
		this.FOL_MakeAccessString(this.VOC_UserCode)

		*- Seta a moeda padrão
		if .FOL_FindExpression("A36_A36_004_N", "A36_A36_004_N", 1)
			this.VOC_Currency = a36_a36_004_n.a36_code
		else
			this.VOC_Currency = ""
		endif
		use in A36_A36_004_N

		*- Conta administrativa de mensagens
		VLC_T77Ukey = .FOU_ChkSets("STARPAR-C-00276")
		if !empty(nvl(VLC_T77Ukey, ""))
			if .FOL_FindExpression("T77_UKEY", "T77", VLC_T77Ukey)
				.VOC_SystemMsgAccount = alltrim(t77.t77_001_C)
	  			.VOC_SystemT77Ukey = t77.ukey
			endif
		else
			.VOC_SystemMsgAccount = ""
  			.VOC_SystemT77Ukey = .null.
  		endif

		if !this.VOL_ServiceMode
			*- Seta os toolbar's
			VGO_ToolBar.FOL_AcessRefresh()
			VGO_ToolBar.FOL_ActionInButton("_ssm_ciai", "", "reload", .t.)
		endif

		if used("grp")
			use in grp
		endif
		if used("A24")
			use in A24
		endif
		if used("A23")
			use in A23
		endif
		if used("A22")
			use in A22
		endif
		if used("A36")
			use in A36
		endif
		if used("CIA")
			use in cia
		endif
		if used("USR")
			use in usr
		endif
	endwith

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ProgressBar("autoincrement", 1)
		VGO_ToolBar.FOL_ProgressBar("visible", .f.)
	endif
	
	*	Inicializo o HELP !!!
	*
	VLC_HelpFile = upper(this.VOC_DirHome+[components\references\]+this.VOC_USER_Language+[\starsoft.chm])
	if file(VLC_HelpFile)
		Set Help To (VLC_HelpFile)
		Set Help On
	else
		Set Help OFF
	endif
else
	this.VOC_CompanyCode    = ""
	this.VOC_UserCode   	= ""
endif

if !empty(this.VOC_TriggersFile)
	try
		release procedure (this.VOC_TriggersFile)
		catch to VLO_Error
	endtry
endif

if VLN_Return = 1 && Logado com sucesso
	this.VOL_LoggedOn = .t.
	this.FOL_InitSemaphore()
	this.FOL_StartMessenger(1)
	*- Verifico se existe o arquivo de resource

	if !empty( this.voc_usercode )
		if !this.FOL_WorkResource(6, "USR", this.VOC_UserCode)
			this.FOL_WorkResource(1, "USR", this.VOC_UserCode, "", .F.)
		endif

		this.FOL_WorkResource(2, "USR", this.VOC_UserCode)
	endif

	this.VOO_Security.FOL_InitCustomEnvironment()

	if !this.VOL_ServiceMode
		VGO_ToolBar.FOL_ActionInButton("_ssm_sysm", "", "reload", .t.)
	endif

else
	this.FOL_CloseSemaphoreHandle()
endif

return VLN_Return
ENDPROC

*-----------------------------------------------------------------------------------------------------------------------
*- Campo com a configuração para relatórios.
*- Retorno   : String de configuração de relatórios
*- Denis Carrizo em 12/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_RelatField
	do case
		case this.VON_DBType = 1 && Sql Server
			return y06.y06_006_m
		case this.VON_DBType = 2 && Oracle
			return y06.y06_013_m
		case this.VON_DBType = 3 && DB2
			return y06.y06_011_m
	endcase
endproc

*-----------------------------------------------------------------------------------------------------------------------
*- Campo com a configuração para pesquisas
*- Retorno   : String de configuração de pesquisas
*- Denis Carrizo em 12/09/2001
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_SearchField
	do case
		case this.VON_DBType = 1 && Sql Server
			return y34.y34_004_m
		case this.VON_DBType = 2 && Oracle
			return y34.y34_013_m
		case this.VON_DBType = 3 && DB2
			return y34.y34_011_m
	endcase
endproc

*/----------------------------------------------------------------------------------------
*/ Parametros  		: PLC_Alias : O nome do table a ser procurado 	  
*/ Descrição   		: Retorna o campo do Y02 que contém a string do select.
*/ Retorno     		: Conteúdo do campo
*/----------------------------------------------------------------------------------------
procedure FOC_SelectField

	do case
		case this.VON_DBType = 1 && Sql Server
			return y02.y02_010_m
		case this.VON_DBType = 2 && Oracle
			return y02.y02_013_m
		case this.VON_DBType = 3 && DB2
			return y02.y02_011_m	
	endcase
endproc


*/--------------------------------------------------------------------------------------/*
*/ FOC_GetPermission
*/ Parâmetros	: 	- PLC_PermissionUkey : Indica se a verificação é baseada nas opções de
*/					permissões configuradas na tela de acesso.
*/					- PLC_UserList : Lista de usuários que podem efetuar a permissão
*/ Descrição    : Permite a validação da senha de um usuário do sistema, baseado em uma
*/					lista de usuários
*/ Retorno      : Ukey do usuário que conseguiu validar, vazio se não houve validação
*/--------------------------------------------------------------------------------------/*
PROCEDURE FOC_GetPermission
	lParameters PLC_PermissionUkey, PLC_UserList

	VLC_UserList = "STAR_"
	VLC_UserUkey = ""

	if !empty(PLC_PermissionUkey)

		this.FOL_ExecuteCursor("Y14_PERMISSIONS", "tmppermission", 0)
		select tmppermission
		index on nvl(usr_ukey, space(5)) tag usr_ukey
		index on nvl(grp_ukey, space(5)) tag grp_ukey

		this.FOL_ExecuteCursor("GXU_ALL", "gxu", 0)
		select gxu
		index on nvl(usr_ukey, space(5)) tag usr_ukey
		index on nvl(grp_ukey, space(5)) tag grp_ukey

		this.FOL_ExecuteCursor("USR_ALL_1", "usr", 0)
		select usr

		VLC_OptionToSearch = ";" + PLC_PermissionUkey + ";"
		VLC_UserList = ""

		scan
			dimension VLA_AccessCfg[4]
			store "" to VLA_AccessCfg
			VLL_StopEval = .f.
			VLL_AddUser = .f.

			*- A prioridade é do acesso ao usuário
			set order to usr_ukey in tmppermission
			if seek(usr.ukey, "tmppermission", "usr_ukey")
				select tmppermission
				scan while tmppermission.usr_ukey == usr.ukey
					if tmppermission.y14_001_c = "SSACS_ACCESS_GRANT_VIEW"
						VLA_AccessCfg[1] = tmppermission.y14_006_m
					else
						*- tmppermission.y14_001_c = "SSACS_ACCESS_DENY_VIEW"
						VLA_AccessCfg[2] = tmppermission.y14_006_m
					endif
				endscan
				
				if at(VLC_OptionToSearch, VLA_AccessCfg[1]) > 0
					*- Acesso permitido
					VLL_StopEval = .t.
					VLL_AddUser = .t.
				else
					if at(VLC_OptionToSearch, VLA_AccessCfg[2]) > 0
						*- Acesso negado
						VLL_StopEval = .t.
					endif
				endif
			endif
			
			if !VLL_StopEval
				set order to grp_ukey in tmppermission
				if seek(usr.grp_ukey, "tmppermission", "grp_ukey")
					select tmppermission
					scan while tmppermission.grp_ukey == usr.grp_ukey and !VLL_StopEval
						if tmppermission.y14_001_c = "SSACS_ACCESS_GRANT_VIEW"
							if at(VLC_OptionToSearch, tmppermission.y14_006_m) > 0
								*- Acesso permitido
								VLL_StopEval = .t.
								VLL_AddUser = .t.
							endif
						endif
					endscan
				endif


				if !VLL_StopEval
					set order to usr_ukey in gxu
					if seek(usr.ukey, "gxu", "usr_ukey")
						select gxu
						scan while gxu.usr_ukey = usr.ukey and !VLL_StopEval

							set order to grp_ukey in tmppermission
							if seek(gxu.grp_ukey, "tmppermission", "grp_ukey")
								select tmppermission
								scan while tmppermission.grp_ukey == gxu.grp_ukey and !VLL_StopEval
									if tmppermission.y14_001_c = "SSACS_ACCESS_GRANT_VIEW"
										if at(VLC_OptionToSearch, tmppermission.y14_006_m) > 0
											*- Acesso permitido
											VLL_StopEval = .t.
											VLL_AddUser = .t.
										endif
									endif
								endscan
							endif

							select gxu
						endscan
					endif

				endif

			endif
			
			if VLL_AddUser
				VLC_UserList = VLC_UserList + ";" + usr.ukey
			endif

			select usr
		endscan

		use in usr
		use in tmppermission
		use in gxu
	endif

	if !empty(PLC_UserList)
		VLC_UserList = VLC_UserList + ";" + PLC_UserList
	endif

	if !empty(VLC_UserList)
		VLC_UserUkey = this.FOL_DoNormalForm("autorizacao", VLC_UserList)
	else
		this.FON_Msg(1170)
	endif

	return VLC_UserUkey
endproc

*----------------------------------------------------------------------------------------
* Função Objeto	: FOL_KeyInfo
* Parâmetros   	: PLC_KeyName	- Nome do índice, poser também uma FK
*                 PRA_KeyArray	- Matriz referência que irá armazenar os detalhes da chave
* Descrição 	: Preenche a matriz referência com as informações sobre a chave solicitada
* Retorno     	: .t. se conseguir, .f. c.c.
*----------------------------------------------------------------------------------------
procedure FOL_KeyInfo
lparameters PLC_KeyName, PRA_KeyArray

local VLN_KeyCount, VLC_Sys07, VLC_Sys08

VLC_Sys07 = this.FOC_TableIdent("SYS07") + "SYS07" && Identificador da tabela SYS07
VLC_Sys08 = this.FOC_TableIdent("SYS08") + "SYS08" && Identificador da tabela SYS08

*- Índices da tabela solicitada, incluindo suas Key´s
VLC_Select = "select sys07.tab_id, sys07.tab_loc, sys08.* from " + VLC_Sys08 + " sys08, " + VLC_Sys07 + " sys07 " +;
		"where sys08.sys07_ukey = sys07.ukey and sys08.idx_name = '" + PLC_KeyName + "'"

VLN_KeyCount = 0
VLL_SqlRet = this.FOL_SqlExec(VLC_Select, "srv_key")

if !VLL_SqlRet or reccount("srv_key") <> 1
	return .f.
endif

select srv_key
if srv_key.idx_type = 9 && Foreign Key
	VLC_RefTable 	= this.FOC_TagContent(srv_key.idx_exp, "REFTABLE", 1)
	VLC_Key 		= this.FOC_TagContent(srv_key.idx_exp, "FKEY", 1)
	VLC_RefExpr		= this.FOC_TagContent(srv_key.idx_exp, "REFEXPR", 1)
else
	VLC_RefTable 	= ""
	VLC_Key 		= ""
	VLC_RefExpr		= ""
endif

dimension PRA_KeyArray[7]
PRA_KeyArray[1] = alltrim(srv_key.idx_name)
PRA_KeyArray[2] = alltrim(srv_key.idx_exp)
PRA_KeyArray[3] = srv_key.idx_type
PRA_KeyArray[4] = VLC_RefTable
PRA_KeyArray[5] = VLC_Key
PRA_KeyArray[6] = VLC_RefExpr
PRA_KeyArray[7] = alltrim(srv_key.tab_id)
use in srv_key

return .t.
endproc

*/--------------------------------------------------------------------------------------------
*/ Parametros       : PLC_Code            : Código de centro de custo a ser avaliado  
*/                    PLC_Pict            : Picture do código 
*/                    PLC_Alias           : Alias a ser pesquisado nível anterior 
*/                    PLN_Seek1           : Indica que o Contábil está clicado  
*/                    PLN_Seek2           : Indica que o Estoque está clicado 
*/                    PLN_Seek3           : Indica que o RH está clicado  
*/                    PLN_Seek4           : Indica que o Finaceiro está clicado 
*/                    PLN_Seek5           : Indica que o Pcp está clicado 
*/                    PLC_TagSeek1   	  : Se for local Tag utilizada para Contabilidade 
*/									   		Se for SQL select do Y02  
*/                    PLC_TagSeek2        : Se for local Tag utilizada para Estoque 
*/									        Se for SQL select do Y02  
*/                    PLC_TagSeek3        : Se for local Tag utilizada para Rh  
*/									        Se for SQL select do Y02  
*/                    PLC_TagSeek4        : Se for local Tag utilizada para Financeiro  
*/									        Se for SQL select do Y02  
*/                    PLC_TagSeek5        : Se for local Tag utilizada para Pcp 
*/									        Se for SQL select do Y02  
*/                    PRC_Fault           : Retorna o módulo com nível anterior não 
*/ encontrado  
*/                    PRC_UkeyParentCont  : Retorna a ukey do nível anterior Contábil 
*/                    PRC_UkeyParentEstq  : Retorna a ukey do nível anterior Estoque  
*/                    PRC_UkeyParentFolha : Retorna a ukey do nível anterior RH 
*/                    PRC_UkeyParentFinan : Retorna a ukey do nível anterior Financeiro 
*/                    PRC_UkeyParentPcp   : Retorna a ukey do nível anterior Pcp  
*/                    PLL_IsLocal    	  : Indica que a tabela a ser pesquisa é local  
*/ Descrição        : Verifica se o centro de custo com nível anterior é sintético  
*/ Retorno          : Retorna .T. se for sintética ou .F. caso contrário  
*/--------------------------------------------------------------------------------------------
PROCEDURE FOL_LevelBackisSintetica
	LParameters PLC_Code, PLC_Pict, PLC_Alias, PLN_Seek1, PLN_Seek2, PLN_Seek3, PLN_Seek4, PLN_Seek5, ;
	PLC_TagSeek1, PLC_TagSeek2, PLC_TagSeek3, PLC_TagSeek4, PLC_TagSeek5, PRC_Fault, PRC_UkeyParentCont,;
	PRC_UkeyParentEstq, PRC_UkeyParentFolha, PRC_UkeyParentFinan, PRC_UkeyParentPcp, PLL_IsLocal

	LOCAL VLC_Code, VLC_ClearString, VLC_Parent, VLN_Pos, VLN_Counter, VLC_Caracter, VLL_Ok,;
	VLN_Level1, VLN_Level2, VLN_Level3, VLN_Level4, VLN_Level5

	store 1 to VLN_Level1, VLN_Level2, VLN_Level3, VLN_Level4, VLN_Level5
	store "" to PRC_UkeyParentCont, PRC_UkeyParentEstq, PRC_UkeyParentFolha, PRC_UkeyParentFinan, PRC_UkeyParentPcp

	this.FOL_PushSelect()

	VLC_Code = PLC_Code
	VLC_ClearString = ""

	*- Retira os caracteres de separação de níveis
	VLC_ClearString = this.FOC_ClearSeparator(VLC_Code)

	*- Coloca mascara que esta na Matriz VGA_PICTS
	VLC_Code = transform(VLC_ClearString, PLC_Pict)
	store "" to VLC_Parent, VLC_ClearString
	VLN_Pos = 0

	*- Varre toda String pesquisando cada nível do código
	for VLN_Counter = len(VLC_Code) to  1  step -1
		VLC_Caracter = substr(VLC_Code, VLN_Counter, 1)
	*--	Se o caracter pesquisado não for Número ou Letra é considerado separador
		if !isdigit(VLC_Caracter) and !isalpha(VLC_Caracter)
			VLC_Parent = substr( VLC_Code, 1, VLN_Counter )
			exit
		endif
	endfor

	*- Retira os caracteres de separação de níveis
	VLC_ClearString = this.FOC_ClearSeparator(VLC_Parent)

	select (PLC_Alias)

	if !empty(PLN_Seek1) or !empty(PLN_Seek2) or !empty(PLN_Seek3) or !empty(PLN_Seek4) or !empty(PLN_Seek5)
		VLC_Alias = strtran(upper(PLC_Alias), "T", "")
		if PLL_IsLocal
	*---	Contábil
			if PLN_Seek1 = 1
				if seek("1" + VLC_ClearString, PLC_Alias, PLC_TagSeek1)
					PRC_UkeyParentCont = ukey
					VLN_Level1 = array_017
				endif
			endif

	*---	Estoque
			if PLN_Seek2 = 1
				if seek("1" + VLC_ClearString, PLC_Alias, PLC_TagSeek2)
					PRC_UkeyParentEstq = ukey
					VLN_Level2 = array_017
				endif
			endif

	*---	Folha de Pagamento
			if PLN_Seek3 = 1
				if seek("1" + VLC_ClearString, PLC_Alias, PLC_TagSeek3)
					PRC_UkeyParentFolha	= ukey
					VLN_Level3 = array_017
				endif
			endif

	*---	Financeiro
			if PLN_Seek4 = 1
				if seek("1" + VLC_ClearString, PLC_Alias, PLC_TagSeek4)
					PRC_UkeyParentFinan = ukey
					VLN_Level4 = array_017
				endif
			endif

	*---	Pcp
			if PLN_Seek5 = 1
				if seek("1" + VLC_ClearString, PLC_Alias, PLC_TagSeek5)
					PRC_UkeyParentPcp = ukey
					VLN_Level5 = array_017
				endif
			endif

		else
	*---	Contábil
			if PLN_Seek1 = 1
				if  this.FOL_FindExpression(PLC_TagSeek1, "TMP_Select", VLC_ClearString)
					PRC_UkeyParentCont = TMP_Select.ukey
					VLN_Level1 = TMP_Select.array_017
					this.FOL_CloseTable("TMP_Select")
				endif
			endif
	*---	Estoque
			if PLN_Seek2 = 1
				if  this.FOL_FindExpression(PLC_TagSeek2, "TMP_Select", VLC_ClearString)
					PRC_UkeyParentEstq = TMP_Select.ukey
					VLN_Level2 = TMP_Select.array_017
					this.FOL_CloseTable("TMP_Select")
				endif
			endif
	*---	Folha de Pagamento
			if PLN_Seek3 = 1
				if  this.FOL_FindExpression(PLC_TagSeek3, "TMP_Select", VLC_ClearString)
					PRC_UkeyParentFolha = TMP_Select.ukey
					VLN_Level3 = TMP_Select.array_017
					this.FOL_CloseTable("TMP_Select")
				endif
			endif
	*---	Financeiro
			if PLN_Seek4 = 1
				if  this.FOL_FindExpression(PLC_TagSeek4, "TMP_Select",  VLC_ClearString)
					PRC_UkeyParentFinan = TMP_Select.ukey
					VLN_Level4 = TMP_Select.array_017
					this.FOL_CloseTable("TMP_Select")
				endif
			endif
	*---	Pcp
			if PLN_Seek5 = 1
				if  this.FOL_FindExpression(PLC_TagSeek5, "TMP_Select", VLC_ClearString)
					PRC_UkeyParentPcp = TMP_Select.ukey
					VLN_Level5 = TMP_Select.array_017
					this.FOL_CloseTable("TMP_Select")
				endif
			endif
		endif

		VLL_Ok = VLN_Level1 = 1 and VLN_Level2 = 1 and VLN_Level3 = 1 and VLN_Level4 = 1 and VLN_Level5 = 1
	endif

	if !VLL_Ok
		if VLN_Level1 = 2
			PRC_Fault = "Contabilidade"
		else
			if VLN_Level2 = 2
				PRC_Fault = "Estoque"
			else
				if VLN_Level3 = 2
					PRC_Fault = "Folha de Pagamento"
				else
					if VLN_Level4 = 2
						PRC_Fault = "Financeiro"
					else
						if VLN_Level5 = 2
							PRC_Fault = "Pcp"
						endif
					endif
				endif
			endif
		endif
	endif
	
	this.FOL_PopSelect()

	return VLL_Ok
ENDPROC

*----------------------------------------------------------------------------------------
* Função Objeto	: FOL_AlterUsrAccess
* Parâmetros   	: PLC_UsrUkey - Ukey do usuário
*				  PLC_OptionStr - Opção que será configurada (Form, Relatório, etc)
*				  PLC_CfgStr - String com as configurações ex.: "1221"
* Descrição 	: Cria ou altera o registro de acesso no Y14 para um determinado usuário
* Retorno     	: .t. se conseguir, .f. c.c.
*----------------------------------------------------------------------------------------
procedure FOL_AlterUsrAccess
lparameters PLC_UsrUkey, PLC_OptionStr, PLC_CfgStr

local VLL_Return, VLN_Counter

if PLC_UsrUkey = "STAR_"
	return .T.
endif

if len(PLC_CfgStr) <> 4
	return .F.
else
	this.FON_Msg(1713)
endif

VLC_Option = ";" + alltrim(PLC_OptionStr) + ";"
for VLN_Counter = 1 to 4
	if substr(PLC_CfgStr, VLN_Counter, 1) = "1"
		this.VOA_AccessCfg[VLN_Counter, 1] = this.VOA_AccessCfg[VLN_Counter, 1] + VLC_Option
	else
		this.VOA_AccessCfg[VLN_Counter, 2] = this.VOA_AccessCfg[VLN_Counter, 2] + VLC_Option
	endif
endfor

return .t.

endproc


*----------------------------------------------------------------------------------------
* Função Objeto	: FOU_PropertyValue
* Parâmetros   	: PLC_String - String onde está a propriedade eo seu valor
*				  PLC_PropertyName - Nome da propriedade
*				  PLU_Default - caso a propriedade não seja encontrada, esse é o retorno
* Descrição 	: Retorna o valor de uma propriedade no padrão #NomedaPropriedade = Valor
* Retorno     	: Valor da propriedade (eval)
*----------------------------------------------------------------------------------------
procedure FOU_PropertyValue
lparameters PLC_String, PLC_PropertyName, PLU_Default

if empty(PLU_Default)
	VLU_Return = ""
else
	VLU_Return = PLU_Default
endif

VLN_StartProp = at(PLC_PropertyName, PLC_String)
VLC_Content = ""

if VLN_StartProp > 0
	VLN_Len = len(PLC_PropertyName)
	VLC_StrToSearch = substr(PLC_String, VLN_StartProp + VLN_Len)
	VLN_EqualPos = at("=", VLC_StrToSearch)
	if VLN_EqualPos > 0
		VLC_Content = substr(VLC_StrToSearch, VLN_EqualPos + 1)
		VLN_CRLF = at(chr(13), VLC_Content)

		if VLN_CRLF > 0
			VLC_Content = substr(VLC_Content, 1, VLN_CRLF - 1)
		endif

		VLC_Content = alltrim(VLC_Content)
		VLL_Eval = .t.
		if len(VLC_Content) > 240
			if !inlist(left(alltrim(VLC_Content), 1), ["['])
				VLU_Return = VLC_Content
				VLL_Eval = .f.
			endif
		endif
		
		if VLL_Eval
			VLU_Return = evaluate(VLC_Content)
		endif
	endif
endif

return VLU_Return
endproc


*----------------------------------------------------------------------------------------
* Função Objeto	: FOC_QuickSQL
* Parâmetros   	: PLC_SourceFld - Campo que será trazido (sem alias)
*				  PLC_SourceTable - Tabela de onde deve ser trazido o conteúdo
*				  PLC_CondField - Campo que deve ser comparado
*				  PLC_CondExpr - Valor que deve ser comparado
*				  PLL_SqlString - Indica se o que deve ser retornado é a string SQL
* Descrição 	: Retorna o conteúdo do campo ou a string SQL para posterior execução
* Rodrigo       : 24/11/2004
*----------------------------------------------------------------------------------------
procedure FOC_QuickSQL
lparameters PLC_SourceFld, PLC_SourceTable, PLC_CondField, PLC_CondExpr, PLL_SqlString

local VLC_Return, VLC_TableFrom, VLC_SqlCmd
VLC_Return = ""

VLC_TableFrom = "STAR_DATA@" + PLC_SourceTable + " " + PLC_SourceTable
if this.VON_DBType = 1 && Sql Server
	VLC_TableFrom = VLC_TableFrom + " (NOLOCK)"
endif
VLC_SqlCmd = "SELECT " + PLC_SourceFld + " FROM " + VLC_TableFrom + " WHERE " + PLC_CondField + " = " + PLC_CondExpr

if empty(PLL_SqlString)
	if this.FOL_SqlExec(VLC_SqlCmd, "TMPQUICK")
		select TMPQUICK
		if !eof("TMPQUICK")
			VLC_Return = eval(PLC_SourceFld)
		endif
		use in TMPQUICK
	endif
else
	VLC_Return = VLC_SqlCmd
endif

return VLC_Return
endproc

*----------------------------------------------------------------------------------------
* Função Objeto	: FOL_VerifyClose
* Parâmetros   	: PLD_Date - Data para verificação do fechamento
*				  PLC_Module - Módulo à verificar se está fechado
*				  PLL_MessageDefault - Indica se executa a mensagem padrão	
* Descrição 	: Verifica se o módulo está fechado, Retorna .T. se fechado e .F. caso contrário
* Gustavo Haidar Viotto 11/01/2002
*----------------------------------------------------------------------------------------
procedure FOL_VerifyClose
lparameters PLD_Date, PLC_Module, PLL_MessageDefault

local VLL_Verify
*- Verifica se a data não está vazia
if !empty(PLD_Date)
	VLL_Verify = PLD_Date <= eval("this.VOD_" + alltrim(PLC_Module) + "ClosingDate")
	if VLL_Verify and PLL_MessageDefault
		this.FON_Msg(1553, dtoc(PLD_Date))
	endif
endif
return VLL_Verify
endproc



*/----------------------------------------------------------------------------------/*
*/	Função				: FOL_TreatError											/*
*/	Descrição			: Método que efetua o tratamento de erro de uma manipulacao /*
*/						de registros												/*
*/	Parâmetros			: PLN_Error -> Numero do Erro								/*
*/						  PLO_Form  -> Form que esta executando o tratamento		/*
*/	Tipo de Retorno 	: .T. Executou a rotina com sucesso, .F. c.c				/*
*/	Grupo				: Form														/*
*/	Procura por			: Alias, Error   											/*
*/ 	Alterado por		: Denis Carrizo												/*
*/	Data da Alteração	: 22/01/2002    											/*
*/	Versão				: 5															/*
*/----------------------------------------------------------------------------------/*
procedure FOL_TreatError
lparameters PLN_Error, PLO_CallerForm, PLL_NotMessage
	LOCAL VLO_ToFile, VLN_I, VLC_FieldName, VLO_Object, VLL_SRVMode

	if PLL_NotMessage
		VLL_SRVMode = this.VOL_ServiceMode
		this.VOL_ServiceMode = .T.
	endif

	if !vartype(PLO_CallerForm) == "O"
		PLO_CallerForm = .null.
	endif

	do case
		case this.VON_SQLError = -9 && Violação de FK
			dimension VLA_FKArray[1]

			if this.FOL_KeyInfo(this.VOC_ErrorPar, @VLA_FKArray)
				VLC_Table1 = alltrim(VLA_FKArray[7])
				VLC_Table2 = alltrim(VLA_FKArray[4])
				VLC_Field = alltrim(VLA_FKArray[5])
				VLC_Table1Desc = this.FOC_Caption("table_" + lower(VLC_Table1))
				VLC_Table2Desc = this.FOC_Caption("table_" + lower(VLC_Table2))
			else
				VLC_Table1 = ""
				VLC_Table2 = ""		
				VLC_Field = ""
				VLC_Table1Desc = ""
				VLC_Table2Desc = ""
			endif

			this.FON_Msg(478, VLC_Table2Desc, VLC_Table1Desc + " (" + VLC_Table1 + "." + VLC_Field + ")")

		case this.VON_SQLError = -7 && Chave única violada
			VLO_Object = .null.
			VLC_FieldName = ""		
			if !empty(this.VOC_ErrorPar)
				*- Procura maiores informações sobre a tag
				dimension VLA_KeyArray[1]
				if this.FOL_KeyInfo(this.VOC_ErrorPar, @VLA_KeyArray)
					VLN_FldCount = occurs(",", VLA_KeyArray[2]) + 1
					VLC_FldExpr	= ""
					VLN_Initial = 1
					VLC_SumStr = ""

					for VLN_Counter = 1 to VLN_FldCount
						VLN_EndPos = at(",", VLA_KeyArray[2], VLN_Counter)
						if VLN_EndPos <= 0 && Não achou, logo é a última posição
							VLN_EndPos = len(VLA_KeyArray[2]) + 1
						endif
						VLC_FldName = lower(alltrim(substr(VLA_KeyArray[2], VLN_Initial, VLN_EndPos - VLN_Initial)))
						if !("_par" $ VLC_FldName or "_ukeyp" $ VLC_FldName)
							VLC_FldExpr = VLC_FldExpr + VLC_SumStr + this.FOC_Caption(VLC_FldName)
							VLC_SumStr = ", "
						endif
						VLN_Initial = VLN_EndPos + 1
					endfor
					
					VLC_TableDesc = this.FOC_Caption("table_" + lower(VLA_KeyArray[7]))
					this.FON_Msg(397, VLC_TableDesc, VLC_FldExpr)
				else
					this.FON_Msg(397)
				endif
			else
				this.FON_Msg(397)
			endif

			if !isnull(PLO_CallerForm)
				PLO_CallerForm.FOL_Posfirstfield()
			endif

		case this.VON_SQLError = -8 && DeadLock
			this.FON_Msg(477)

		case this.VON_SQLError = -6 && Excedido o tempo de espera (normalmente lock de registro)
			this.FON_Msg(308)


		case this.VON_SQLError = -5 && Primary Key violada
			VLC_TableStr = " (" + this.VOC_ErrorPar + ") " + this.FOC_Caption("table_" + lower(this.VOC_ErrorPar))
			this.FON_Msg(397, VLC_TableStr)


		case this.VON_SQLError = -1 && Registro modificado (timestamp diferente)
			if !isnull(PLO_CallerForm) and pemstatus(PLO_CallerForm, "VOC_AliasDefault", 5) and ;
				pemstatus(PLO_CallerForm, "VOC_AliasADO", 5)

				select (PLO_CallerForm.VOC_AliasDefault)
				if this.FOL_FindExpression(PLO_CallerForm.VOC_AliasADO, PLO_CallerForm.VOC_AliasADO, ukey)
					if this.FON_Msg(305) = 6
						select (PLO_CallerForm.VOC_AliasADO)
						scatter memvar memo
						select (PLO_CallerForm.VOC_AliasDefault)
						gather memvar memo
						this.FOL_TableUpdate(PLO_CallerForm.VOC_AliasDefault, .F.)
						PLO_CallerForm.VOC_LastExecuteCursor = ""
					endif
					use in (PLO_CallerForm.VOC_AliasADO)
				else
					if !PLO_CallerForm.VOL_Adding
				*--		Registro não existe no SQL
						this.FON_Msg(1306)
					endif
				endif
			else
				*--		Registro não existe no SQL
				this.FON_Msg(1306)
			endif

		case this.VON_SQLError = -10 && Falta de espaço no banco de dados
			this.FON_Msg(1227)
			on shutdown
			quit

		otherwise
			if !this.VOL_ServiceMode
				messagebox("DB Server Error #" + transform(this.VON_SQLError) + ".", 48)
			endif

	endcase

	if PLL_NotMessage
		this.VOL_ServiceMode = VLL_SRVMode
	endif

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Grava uma string no registro
*/	PLN_HKey   : valor da chave primaria HKEY
*/  PLC_Subkey : O valor da sub-chave de registro
*/  PLC_Entry  : A chave para se gravar o valor
*/  PLU_Value  : Valor para gravar ou .NULL. para apagar a chave
*/  PLL_Create : Cria a chave se não existir
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_WriteRegistryString
lparameters PLN_HKey, PLC_Subkey, PLC_Entry, PLU_Value, PLL_Create, PLN_RegType
local VLN_RegHandle, VLN_Result, VLN_Size, VLC_DataBuffer, VLN_Type

	PLN_HKey = iif(!empty(PLN_HKey), PLN_HKey, HKEY_LOCAL_MACHINE)

	VLN_RegHandle = 0

	VLN_Result = RegOpenKey(PLN_HKey, PLC_Subkey, @VLN_RegHandle)
	if VLN_Result # ERROR_SUCCESS
	   if !PLL_Create
	      return .f.
	   else
	      VLN_Result = RegCreateKey(PLN_HKey, PLC_Subkey, @VLN_RegHandle)
	      if VLN_Result # ERROR_SUCCESS
	         return .f.
	      endif
	   endif
	endif

	*- Verifica se for .NULL. para apagar a chave
	if !isnull(PLU_Value)
		VLN_Size = 0
			
		do case
		  	case empty(PLN_RegType)
		  		PLN_RegType = REG_SZ
		  		VLN_Size = len(PLU_Value)

			  	if VLN_Size = 0
			     	PLC_Value = chr(0)
			  	endif

		  	case PLN_RegType = 1
		  		PLN_RegType = REG_DWORD
				VLN_Size = 4
				PLU_Value = this.FOC_WordToChar(PLU_Value)

		  	case PLN_RegType = 2
		  		PLN_RegType = REG_BINARY
		endcase

		VLN_Result = RegSetValueEx(VLN_RegHandle, PLC_Entry, 0, PLN_RegType, PLU_Value, VLN_Size)
	else
	  *- Apaga a chave
	  VLN_Result = RegDeleteValue(VLN_RegHandle, PLC_Entry)
	endif

	RegCloseKey(VLN_RegHandle)
	                        
	if VLN_Result # ERROR_SUCCESS
	   return .F.
	endif

	return .t.
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Adiciona uma tarefa no arquivo job
*/	PLN_JobType			: Tipo de job de acordo com array_394
*/	PLN_JobProperties	: Parâmetros para a execução do job no formato XML
*/	PLC_UsrUkey			: Usuário que solicitou a tarefa
*/	PLC_CiaUkey			: Empresa do contexto da tarefa
*/	PLN_Priority		: Prioridade da tarefa
*/	PLT_Schedule		: Data/Hora para execução da tarefa (no caso de agendamento)
*/	PLC_Par				: Referência do arquivo que originou a tarefa
*/	PLC_Ukeyp			: Referência da ukey do arquivo PLC_Par
*/  PLC_RecRule			: Regra de agendamento ou caso processamento periódico
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_CreateJob
lparameters PLN_JobType, PLC_JobProperties, PLC_UsrUkey, PLC_CiaUkey, PLN_Priority, PLT_Schedule, PLC_Par, PLC_Ukeyp, PLC_RecRule

	local VLL_Return, VLC_RetUkey, VLC_JobXML

	VLC_RetUkey = ""

	if !empty(PLC_UsrUkey) and !empty(PLC_CiaUkey)
		VLL_Return = .t.
	endif

	if !empty(PLC_RecRule)
		PLT_Schedule = this.FOT_NextExecDateTime(PLC_RecRule, .null.)
	else
		if empty(nvl(PLT_Schedule, ""))
			PLT_Schedule = datetime()
		endif
	endif

	if VLL_Return
		*- Traz a estrutura do arquivo de tarefas
		with this
			VLC_JobXML = "<job_properties>" + PLC_JobProperties + "</job_properties>"
			if !empty(PLC_RecRule)
				VLC_JobXML = VLC_JobXML + "<job_recurring>" + PLC_RecRule + "</job_recurring>"
			endif
		
			.FOL_ExecuteCursor("JOB_FRM", "JOB", 0)
			.FOL_NewRecordToForm("JOB")
			replace job.array_394	with PLN_JobType,;
					job.job_001_m 	with VLC_JobXML,;
					job.usr_ukey 	with PLC_UsrUkey,;
					job.cia_ukey 	with PLC_CiaUkey,;
					job.job_002_n 	with PLN_Priority,;
					job.job_003_m 	with .null.,;
					job.array_395 	with 1,;
					job.job_004_t 	with PLT_Schedule,;
					job.job_par 	with PLC_Par,;
					job.job_ukeyp 	with PLC_Ukeyp;
			in job
				
			.FOL_CreateSQLString("JOB", "JOB", .t., .t.)
			.FOL_TableUpdate("JOB", .F.)

			if this.FOL_SaveRecordSql("JOB", "JOB", "UKEY")
				VLC_RetUkey = job.ukey
			endif
			use in job
		endwith
	endif	

	return VLC_RetUkey
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Apaga uma tarefa no arquivo job
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_DropJob
lparameters PLC_JobUkey

	local VLL_Return

	with this
		.FOL_SetParameter(1, PLC_JobUkey)
		if .FOL_ExecuteCursor("JOB_UKEY", "JOB", 1)
			VLL_Return = .FOL_DeleteRecordSql("JOB", "JOB", "UKEY")
			use in job
		endif
	endwith

	return VLL_Return
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Imprime um relatório sem necessidade interface
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_DirectPrintReport
lparameters PLC_ReportUkey, PLC_Parameters, PLN_PrintType, PLL_Reprocess, PLL_KeepFileResult

	local VLC_ReportName, VLN_RepoCounter, VLN_Counter, VLC_Separator, VLC_ReportStr
	VLC_ReportName = ""
	VLC_Separator = ""
	VLC_ReportStr = ""
	local array VLA_RetArray[1]

	this.FOL_ResetReportResults()

	VLN_RepoCounter = this.FON_GetReports(1, PLC_ReportUkey, @VLA_RetArray)
	
	for VLN_Counter = VLN_RepoCounter to 1 step -1
		if VLA_RetArray[VLN_Counter, 1] == PLC_ReportUkey
			VLC_ReportStr = VLA_RetArray[VLN_Counter, 3]
			VLC_ReportName = VLA_RetArray[VLN_Counter, 11]
		endif
	endfor

	if empty(VLC_ReportStr)
		return .f.
	endif

	VLO_ReportPrinter = createobject("CGO_ReportPrinter")
	with VLO_ReportPrinter
		.VOC_ReportTitle = VLC_ReportName
		.VOC_ReportUkey = PLC_ReportUkey
		this.VOC_ReportName = .VOC_ReportTitle
		.FOL_LoadConfig(VLC_ReportStr)
	endwith

	for VLN_AskCounter = 1 to VLO_ReportPrinter.VON_AskCounter
		VLC_FilterCfg = this.FOC_TagContent(PLC_Parameters, "filter", VLN_AskCounter)
			
		if !empty(VLC_FilterCfg)
			VLN_FilterId = val(this.FOC_TagContent(VLC_FilterCfg, "id", 1))
			VLN_FilterOp = val(this.FOC_TagContent(VLC_FilterCfg, "op", 1))
			VLC_FilterExp1 = this.FOC_TagContent(VLC_FilterCfg, "exp1", 1)
			VLC_FilterExp2 = this.FOC_TagContent(VLC_FilterCfg, "exp2", 1)
			VLC_Neg = this.FOC_TagContent(VLC_FilterCfg, "neg", 1)
    		VLC_EvalExpr1 = this.FOC_TagContent(VLC_FilterCfg, "evalexp1", 1)
    		VLC_EvalExpr2 = this.FOC_TagContent(VLC_FilterCfg, "evalexp2", 1)

    		if !empty(VLC_EvalExpr1)
	    		VLO_ErrorOnExpr = null
	    		VLL_ErrorOnExpr = .f.
	    		VLC_TempFilterExp = ""

    			try
    				VLC_TempFilterExp = execscript(VLC_EvalExpr1)
    				
    				catch to VLO_ErrorOnExpr
    					VLL_ErrorOnExpr = .t.
    			endtry
    			
    			if !VLL_ErrorOnExpr
    				VLC_FilterExp1 = transform(VLC_TempFilterExp)
    			endif
    		endif

    		if !empty(VLC_EvalExpr2)
	    		VLO_ErrorOnExpr = null
	    		VLL_ErrorOnExpr = .f.
	    		VLC_TempFilterExp = ""

    			try
    				VLC_TempFilterExp = execscript(VLC_EvalExpr2)
    				
    				catch to VLO_ErrorOnExpr
    					VLL_ErrorOnExpr = .t.
    			endtry
    			
    			if !VLL_ErrorOnExpr
    				VLC_FilterExp2 = transform(VLC_TempFilterExp)
    			endif
    		endif

		else
			VLN_FilterId = VLN_AskCounter
			VLN_FilterOp = 1
			VLC_FilterExp1 = ""
			VLC_FilterExp2 = ""
			VLC_Neg = "0"
			VLC_EvalExpr1 = ""
			VLC_EvalExpr2 = ""
		endif

		with VLO_ReportPrinter.VOA_AskObject[VLN_FilterId, 2]

			if VLN_FilterOp = 1 and .VOC_FieldType $ "D/T"
				.VOC_Operator = this.FOC_Operator(10)
			else
				.VOC_Operator = this.FOC_Operator(VLN_FilterOp)
			endif
				
			.VOL_Negative = lower(VLC_Neg) $ "1/on/yes/.t."
			.FOL_RedimValues(1)

			do case
				case alltrim(.VOC_FieldType) $ "C/M/Y"
					do case
						case VLN_FilterOp = 10 && Between
							.FOL_RedimValues(2)
							.VOA_Values[1] = alltrim(VLC_FilterExp1)
							.VOA_Values[2] = alltrim(VLC_FilterExp2)

						case VLN_FilterOp = 9 && In
							VLC_InExpression = strtran(alltrim(VLC_FilterExp1), ",", chr(13))
							VLC_InExpression = strtran(alltrim(VLC_InExpression), ";", chr(13))
							VLC_InExpression = strtran(alltrim(VLC_InExpression), "|", chr(13))
							aline(VLA_InArray, VLC_InExpression, .T.)
							VLN_InCount = alen(VLA_InArray, 1)
							.FOL_RedimValues(VLN_InCount)
							for VLN_Counter = 1 to VLN_InCount
								.VOA_Values[VLN_Counter] = VLA_InArray[VLN_Counter]
							endfor

						case VLN_FilterOp = 8 && %Like%
							.VOA_Values[1] = "%"+alltrim(VLC_FilterExp1)+"%"

						case VLN_FilterOp = 7 && Like%
							.VOA_Values[1] = "%"+alltrim(VLC_FilterExp1)

						case VLN_FilterOp = 6 && %Like
							.VOA_Values[1] = alltrim(VLC_FilterExp1)+"%"

						otherwise
							.VOA_Values[1] = alltrim(VLC_FilterExp1)
					endcase

				case .VOC_FieldType $ "B/N"
					if VLN_FilterOp = 10 && Between
						.FOL_RedimValues(2)
						.VOA_Values[1] = val(VLC_FilterExp1)
						.VOA_Values[2] = val(VLC_FilterExp2)
					else
						.VOA_Values[1] = val(VLC_FilterExp1)
					endif

				case .VOC_FieldType = "L"
					.VOA_Values[1] = (lower(alltrim(VLC_FilterExp1)) $ "1/on/yes/.t.")
				
				case .VOC_FieldType = "D"

					if "/" $ VLC_FilterExp1 or "\" $ VLC_FilterExp1 or "-" $ VLC_FilterExp1
						VLL_DateFormat = 1
					else
						VLL_DateFormat = 2
					endif

					do case 
						case VLN_FilterOp = 10 && Between
							.FOL_RedimValues(2)
							if VLL_DateFormat = 1
								.VOA_Values[1] = ctod(VLC_FilterExp1)
								.VOA_Values[2] = dtot(ctod(VLC_FilterExp2)+1)-60
							else
								.VOA_Values[1] = this.FOD_StoD(VLC_FilterExp1)
								.VOA_Values[2] = dtot(this.FOD_StoD(VLC_FilterExp2)+1)-60
							endif

						case VLN_FilterOp = 1 && Igual
							.FOL_RedimValues(2)
							if VLL_DateFormat = 1
								.VOA_Values[1] = ctod(VLC_FilterExp1)
								.VOA_Values[2] = dtot(ctod(VLC_FilterExp1)+1)-60
							else
								.VOA_Values[1] = this.FOD_StoD(VLC_FilterExp1)
								.VOA_Values[2] = dtot(this.FOD_StoD(VLC_FilterExp1)+1)-60
							endif

						otherwise
							if empty(VLC_FilterExp1) 
								.VOA_Values[1] = ctod("01/01/1900")
							else
								if VLL_DateFormat = 1
									.VOA_Values[1] = ctod(VLC_FilterExp1)
								else
									.VOA_Values[1] = this.FOD_StoD(VLC_FilterExp1)
								endif
							endif
					endcase

				case .VOC_FieldType = "T"
					if VLN_FilterOp = 10 && Between
						.FOL_RedimValues(2)
						.VOA_Values[1] = ctot(VLC_FilterExp1)
						.VOA_Values[2] = ctot(VLC_FilterExp2)
					else
						.VOA_Values[1] = ctot(VLC_FilterExp1)
					endif

				case .VOC_FieldType = "A"
					.VOA_Values[1] = val(VLC_FilterExp1)
			endcase

		endwith

	endfor

	return VLO_ReportPrinter.FOL_Print(PLN_PrintType, PLL_Reprocess, PLL_KeepFileResult)
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Armazena valores para que possam ser utilizados nos filtros de relatórios
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_SetReportFilter
lparameters PLC_QuestionId, PLC_Operator, PLU_Value, PLU_Value2

with this
	if !empty(PLC_QuestionId)
		VLN_ArrayId = ascan(.VOA_ReportFilter, PLC_QuestionId, -1, -1, 1)
		if VLN_ArrayId <= 0
			.VON_ReportFilterCount = .VON_ReportFilterCount + 1
			VLN_ArrayId = .VON_ReportFilterCount
			dimension .VOA_ReportFilter[VLN_ArrayId, 4]
			.VOA_ReportFilter[VLN_ArrayId, 1] = PLC_QuestionId
		else
			VLN_ArrayId = asubscript(.VOA_ReportFilter, VLN_ArrayId, 1)
		endif

		.VOA_ReportFilter[VLN_ArrayId, 2] = PLC_Operator
		.VOA_ReportFilter[VLN_ArrayId, 3] = PLU_Value
		.VOA_ReportFilter[VLN_ArrayId, 4] = PLU_Value2
	endif
endwith

return .t.

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna um valor armazenado na pilha de armazenadores de relatórios, .null. quando não encontra
*/ PLN_ReturnType: 0-Existe o identificador, 1-Operador, 2-Valor#1, 3-Valor#2
*/------------------------------------------------------------------------------------------------------------------------
procedure FOU_GetReportFilter
lparameters PLC_QuestionId, PLN_ReturnType

local VLU_Return

VLU_Return = .null.

with this
	if !empty(PLC_QuestionId)
		VLN_ArrayId = ascan(.VOA_ReportFilter, PLC_QuestionId, -1, -1, 1)
		if empty(PLN_ReturnType)
			VLU_Return = VLN_ArrayId > 0
		else
			if VLN_ArrayId > 0
				VLN_ArrayId = asubscript(.VOA_ReportFilter, VLN_ArrayId, 1)
				VLU_Return = .VOA_ReportFilter[VLN_ArrayId, PLN_ReturnType + 1]
			endif
		endif
	endif
endwith

return VLU_Return

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Apaga todos os valores armazenados nos filtros de relatórios
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_ResetReportFilter

with this
	dimension .VOA_ReportFilter[1]
	.VOA_ReportFilter = .f.
	.VON_ReportFilterCount = 0
endwith
return .t.

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Verifica o status de um JOB, trazendo seu retorno se ja foi processado
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_GetJobStatus
lparameters PLC_JobUkey, PRC_Return

	local VLN_Return

	VLN_Return = 0

	if !empty(nvl(PLC_JobUkey, ""))
		with this
			.FOL_SetParameter(1, PLC_JobUkey)
			.FOL_ExecuteCursor("JOB_STATUS", "JOB", 1)

			if reccount("JOB") = 1
				VLN_Return = job.array_395
				PRC_Return = job.job_003_m
			endif
			use in job
		endwith
	endif	

	return VLN_Return
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Procura uma String exatamente igual em outra String separada por caracteres que não letras, dígitos ou _
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_SearchIn

lparameters PLC_SearchExpression, PLC_ExpressionSearched

local VLL_Ret, VLN_Pos, VLN_Occurs, VLC_Before, VLC_After, VLN_Cont

VLN_Occurs = occurs(PLC_SearchExpression, PLC_ExpressionSearched)

for VLN_Cont = 1 to VLN_Occurs
	VLN_Pos = atcc(PLC_SearchExpression, PLC_ExpressionSearched, VLN_Cont)
	VLC_Before = substr(PLC_ExpressionSearched, VLN_Pos-1, 1)

	if empty(VLC_Before)
		VLC_Before = '\'
	endif

	VLC_After = substr(PLC_ExpressionSearched, VLN_Pos+len(PLC_SearchExpression), 1)

	if empty(VLC_After)
		VLC_After = '\'
	endif

	if !isdigit(VLC_Before) and !isalpha(VLC_Before) and asc(VLC_Before) <> 95 and !isdigit(VLC_After) and !isalpha(VLC_After) and asc(VLC_After) <> 95
		VLL_Ret = .T.
		exit
	endif

endfor

return VLL_Ret

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Prepara uma string padrão com os filtros de um relatório para impressão com a FOL_DirectPrintReport
*/	PLC_Y06Ukey	: Ukey do relatório (Y06) ou ID de um relatório específico
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_CreateReportFilterStr
lparameters PLC_Y06Ukey

	local VLC_ReportName, VLN_RepoCounter, VLN_Counter, VLC_Separator, VLC_ReportStr, VLC_Properties
	VLC_ReportName = ""
	VLC_Separator = ""
	VLC_ReportStr = ""
	local array VLA_RetArray[1]

	VLN_RepoCounter = this.FON_GetReports(1, PLC_Y06Ukey, @VLA_RetArray)
	
	for VLN_Counter = 1 to VLN_RepoCounter
		VLC_ReportName = VLC_ReportName + VLC_Separator + alltrim(VLA_RetArray[VLN_Counter, 2])
		VLC_Separator = " - "
		if VLA_RetArray[VLN_Counter, 1] == PLC_Y06Ukey
			VLC_ReportStr = VLA_RetArray[VLN_Counter, 3]
		endif
	endfor

	if empty(VLC_ReportStr)
		return ""
	endif

	VLO_ReportPrinter = createobject("CGO_ReportPrinter")
	with VLO_ReportPrinter
		.VOC_ReportTitle = VLC_ReportName
		.FOL_LoadConfig(VLC_ReportStr)
	endwith

	VLC_Properties = ""

	for VLN_Counter = 1 to VLO_ReportPrinter.VON_AskCounter
		with VLO_ReportPrinter.VOA_AskObject[VLN_Counter, 2]
			VLC_Type = .VOC_FieldType
			VLC_ArrayName = .VOC_SourceArray
			VLC_AdditionalOption = .VOC_AdditionalOption
	    	VLN_SelectedOption = this.FON_Operator(.VOC_Operator)
			VLC_QuestionId = .VOC_QuestionId
			VLL_Hidden = .VOL_Hidden

			VLU_InitValue1 = ""
			VLU_InitValue2 = ""

			if VLC_Type = "A"
				VLC_SourceArray = .VOC_SourceArray
				VLN_ArrayLen = alen(&VLC_SourceArray)
				if empty(.VON_InitialValue)
					if !empty(.VOC_AdditionalOption)
						VLU_InitValue1 = transform(VLN_ArrayLen + 1)
					else
						VLU_InitValue1 = "1"
					endif
				else
					VLN_InitialValue = transform(.VON_InitialValue)
				endif
		
			endif
		endwith		

		VLN_Default = 0
		VLC_AskNo = transform(VLN_Counter, "@L 999")
	    VLC_FilterType = VLC_Type
	    VLC_Value = ""
		VLC_OrderNo = transform(VLN_Counter)
	    VLC_OperatorObj = "Operator" + VLC_OrderNo
	    VLC_TextObj = "ReportFilter" + VLC_OrderNo

		*- Inicialização dos filtros padrões
		if !empty(VLC_QuestionId)
			if this.FOU_GetReportFilter(VLC_QuestionId)
				with this
					VLN_SelectedOption	= .FON_Operator(.FOU_GetReportFilter(VLC_QuestionId, 1))
					VLU_InitValue1 		= .FOU_GetReportFilter(VLC_QuestionId, 2)
					VLU_InitValue2 		= .FOU_GetReportFilter(VLC_QuestionId, 3)
				endwith
			endif
		endif

	   VLC_Properties = VLC_Properties + "<filter><id>" + VLC_OrderNo + "</id>" +;
	    			"<op>" + transform(VLN_SelectedOption, "@L 99") + "</op>" +;
	    			"<exp1>" + transform(VLU_InitValue1) + "</exp1>" +;
	    			"<exp2>" + transform(VLU_InitValue2) + "</exp2>" +;
	    			"<evalexp1></evalexp1>" +;
	    			"<evalexp2></evalexp2>" +;
	    			"</filter>"
				
	endfor

	return VLC_Properties

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Conversão para o tipo WORD (usado em DLL´s)
*/	PLN_Number	: Valor numérico para conversão
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_WordToChar
lparameters PLN_Number

	return chr(bitand(255, PLN_Number)) + chr(bitand(65280, PLN_Number)%255) + chr(bitand(16711680, PLN_Number)%255)+; 
		chr(bitand(4278190080, PLN_Number)%255)
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Conversão para o tipo Caracter (usado em DLL´s)
*/	PLC_Buffer	: Valor caracter para conversão
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_CharToWord
lparameters PLC_Buffer


	return asc(substr(PLC_Buffer, 1, 1)) + asc(substr(PLC_Buffer, 2, 1))*256 + asc(substr(PLC_Buffer, 3, 1))*65536 + ;
		asc(substr(PLC_Buffer, 4, 1))*16777216 

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna o primeiro dia da semana solicitada (sempre é um domingo)
*/	PLN_WeekNo	: Número da semana (1-53)
*/  PLN_Year	: Ano da semana
*/------------------------------------------------------------------------------------------------------------------------
procedure FOD_FirstDayOfWeek
lparameters PLN_WeekNo, PLN_Year

	local VLD_Return
	VLD_Return = {}

	if between(PLN_WeekNo, 1, 53) and between(PLN_Year, 100, 9999)
		VLD_1stYearDay = date(PLN_Year, 1, 1)
		VLN_1stDOW = dow(VLD_1stYearDay)
		VLD_Return = VLD_1stYearDay - (VLN_1stDOW - 1) + ((PLN_WeekNo - 1) * 7)
	endif

	return VLD_Return

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna o número de semanas existentes em um ano especificado
*/  PLN_Year	: Ano para avaliação
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_WeeksInYear
lparameters PLN_Year

	local VLN_Return
	VLN_Return = 0

	if between(PLN_Year, 100, 9999)
		VLD_LastYearDay = date(PLN_Year, 12, 31)
		do while VLN_Return <= 1
			VLN_Return = week(VLD_LastYearDay)
			VLD_LastYearDay = VLD_LastYearDay - 1
		enddo
	endif

	return VLN_Return

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Retorno idêntico ao da função week() do fox, porém respeita o ano de referência (retorna 52 e não 1 na data 31/12/2001)
*/  PLN_Year	: Ano para avaliação
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_Week
lparameters PLD_Data

	local VLN_Return
	VLN_Return = 0

	VLN_Return = week(PLD_Data)
	if VLN_Return = 1 and month(PLD_Data) = 12
		VLN_Return = this.FON_WeeksInYear(year(PLD_Data))
	endif

	return VLN_Return

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna o conteúdo de uma propriedade hidden
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_GetMaintenance
	return this.VOO_Security.FON_GetMaintenance()
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna o número de pesquisas de acordo com o tipo e chave, preenchendo uma array (referência) com suas configurações
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_GetSearchs
lparameters PLN_SearchType, PLC_SearchKey, PRA_RetSearchArray, PLC_SearchId, PLL_ChkKey

	local VLN_RetSearchCount, VLN_ClientSearchCount, VLN_Counter, VLL_SearchsForAccess, VLL_AllSearchs

	VLN_RetSearchCount = 0
	PLL_ChkKey = .t.
	VLL_SearchsForAccess = .f.
	VLL_AllSearchs = (PLC_SearchKey == "_ALL_")

	this.FOL_UseInPackage("y34", "y34")
	
	do case
		*- Pesquisa por chave de procura (ponto de chamada)
		case empty(PLC_SearchId) and !empty(PLN_SearchType) and !empty(PLC_SearchKey)
			select y34
			set order to main in y34

			if VLL_AllSearchs
				PLC_SearchKey = ""
			endif

			if seek(str(PLN_SearchType) + PLC_SearchKey, "Y34", "main")
				scan while y34.y34_002_n = PLN_SearchType and (alltrim(PLC_SearchKey) == alltrim(y34.y34_003_c) or;
					VLL_AllSearchs)
					
					if y34.Y34_003_C = "CIA" or !this.FOL_AcessMenu(y34.ukey)
						VLN_RetSearchCount = VLN_RetSearchCount + 1	
						dimension PRA_RetSearchArray[VLN_RetSearchCount, 9]
						PRA_RetSearchArray[VLN_RetSearchCount, 1] = this.FOC_Caption(alltrim(y34.y34_001_c))
						PRA_RetSearchArray[VLN_RetSearchCount, 2] = this.FOC_SearchField()
						PRA_RetSearchArray[VLN_RetSearchCount, 3] = ""
						PRA_RetSearchArray[VLN_RetSearchCount, 4] = 0
						PRA_RetSearchArray[VLN_RetSearchCount, 5] = ""
						PRA_RetSearchArray[VLN_RetSearchCount, 6] = ""
						PRA_RetSearchArray[VLN_RetSearchCount, 7] = y34.y34_002_n
						PRA_RetSearchArray[VLN_RetSearchCount, 8] = y34.ukey
						PRA_RetSearchArray[VLN_RetSearchCount, 9] = y34.y34_003_c
					endif
				endscan
			endif

		*- Todas as pesquisas para controle de acesso
		case empty(PLC_SearchId) and empty(PLN_SearchType) and empty(PLC_SearchKey)
			VLL_SearchsForAccess = .t.
			PLC_SearchKey = ""
			PLC_SearchId = ""

			select y34
			VLN_FldCount = afields(VLA_Y34Struct, "y34") + 1
			dimension VLA_Y34Struct[VLN_FldCount, AFIELDS2NDCOL]
			VLA_Y34Struct(VLN_FldCount, 1) = "esp"
			VLA_Y34Struct(VLN_FldCount, 2) = "l"
			VLA_Y34Struct(VLN_FldCount, 3) = 1
			VLA_Y34Struct(VLN_FldCount, 4) = 0
			store "" to VLA_Y34Struct(VLN_FldCount, 7), VLA_Y34Struct(VLN_FldCount, 8),  VLA_Y34Struct(VLN_FldCount, 9), ;
						VLA_Y34Struct(VLN_FldCount, 10), VLA_Y34Struct(VLN_FldCount, 11), VLA_Y34Struct(VLN_FldCount, 12), ;
						VLA_Y34Struct(VLN_FldCount, 13), VLA_Y34Struct(VLN_FldCount, 14), VLA_Y34Struct(VLN_FldCount, 15), ;
						VLA_Y34Struct(VLN_FldCount, 16)
			
			VLN_FldCount = VLN_FldCount + 1
			dimension VLA_Y34Struct[VLN_FldCount, AFIELDS2NDCOL]
			VLA_Y34Struct(VLN_FldCount, 1) = "descr"
			VLA_Y34Struct(VLN_FldCount, 2) = "c"
			VLA_Y34Struct(VLN_FldCount, 3) = 220
			VLA_Y34Struct(VLN_FldCount, 4) = 0
			store "" to VLA_Y34Struct(VLN_FldCount, 7), VLA_Y34Struct(VLN_FldCount, 8),  VLA_Y34Struct(VLN_FldCount, 9), ;
						VLA_Y34Struct(VLN_FldCount, 10), VLA_Y34Struct(VLN_FldCount, 11), VLA_Y34Struct(VLN_FldCount, 12), ;
						VLA_Y34Struct(VLN_FldCount, 13), VLA_Y34Struct(VLN_FldCount, 14), VLA_Y34Struct(VLN_FldCount, 15), ;
						VLA_Y34Struct(VLN_FldCount, 16)

			create cursor tmp_y34 from array VLA_Y34Struct
			select tmp_y34
			index on str(y34_002_n) + y34_003_c + y34_009_c tag main
			
			select y34
			set order to main in y34
			scan
				VLN_RetSearchCount = VLN_RetSearchCount + 1	
				scatter memvar
				append blank in tmp_y34
				m.y34_001_c = ""
				m.descr = this.FOC_Caption(alltrim(y34.y34_001_c))
				m.esp = .f.
				select tmp_y34
				gather memvar

				select y34
			endscan

		*- Pesquisa por ukey de pesquisa
		case !empty(PLC_SearchId)
			select y34
			set order to ukey in y34
			if seek(PLC_SearchId, "Y34", "ukey")
				if y34.Y34_003_C = "CIA" or !this.FOL_AcessMenu(y34.ukey)
					VLN_RetSearchCount = VLN_RetSearchCount + 1	
					dimension PRA_RetSearchArray[VLN_RetSearchCount, 9]
					PRA_RetSearchArray[VLN_RetSearchCount, 1] = this.FOC_Caption(alltrim(y34.y34_001_c))
					PRA_RetSearchArray[VLN_RetSearchCount, 2] = this.FOC_SearchField()
					PRA_RetSearchArray[VLN_RetSearchCount, 3] = ""
					PRA_RetSearchArray[VLN_RetSearchCount, 4] = 0
					PRA_RetSearchArray[VLN_RetSearchCount, 5] = ""
					PRA_RetSearchArray[VLN_RetSearchCount, 6] = ""
					PRA_RetSearchArray[VLN_RetSearchCount, 7] = y34.y34_002_n
					PRA_RetSearchArray[VLN_RetSearchCount, 8] = y34.ukey
					PRA_RetSearchArray[VLN_RetSearchCount, 9] = y34.y34_003_c
				endif
			endif
	endcase

	if used("y34")
		use in y34
	endif

	local array VLA_SearchFile[1]

	*- Pesquisas específicas do cliente
	VLC_ObjName = alltrim(PLC_SearchKey)
	VLN_ClientSearchCount = adir(VLA_SearchFile, this.VOC_ClientSearchDir + "*.sss")

	for VLN_Counter = 1 to VLN_ClientSearchCount
		VLL_AddSearch = .f.
		VLC_ReportCfg = ""
		VLC_FileName = this.VOC_ClientSearchDir + VLA_SearchFile[VLN_Counter, 1]
		VLC_FileContent = filetostr(VLC_FileName)
		VLC_HeaderContent = strextract(VLC_FileContent, "\\BEGIN SEARCHHEADER", "\\END SEARCHHEADER", 1, 1)
		if !empty(VLC_HeaderContent)
			VLN_HeaderLines = alines(VLA_Header, VLC_HeaderContent, .t.)
			
			VLC_AuxFields = ""
			VLC_ObjList = ""
			VLC_SearchId = ""
			VLN_SearchType = 0
			VLC_Description = ""
			VLN_ShowOrder = 0
			VLC_AuxFields = ""
			VLN_FilledProperties = 0
			VLC_ValidKey = ""

			for VLN_HCounter = 1 to VLN_HeaderLines
				VLC_Line = VLA_Header[VLN_HCounter]

				if !empty(VLC_Line)
					
					do case
						case "#VOC_OBJFILTER" $ VLC_Line
							VLC_ObjList = this.FOU_PropertyValue(VLC_Line, "#VOC_OBJFILTER")
						case "#VOC_SEARCHID" $ VLC_Line
							VLC_SearchId = this.FOU_PropertyValue(VLC_Line, "#VOC_SEARCHID")
						case "#VON_SEARCHTYPE" $ VLC_Line
							VLN_SearchType = this.FOU_PropertyValue(VLC_Line, "#VON_SEARCHTYPE", 0)
						case "#VOC_DESCRIPTION" $ VLC_Line
							VLC_Description = this.FOU_PropertyValue(VLC_Line, "#VOC_DESCRIPTION")
						case "#VON_SHOWORDER" $ VLC_Line
							VLN_ShowOrder = this.FOU_PropertyValue(VLC_Line, "#VON_SHOWORDER", 0)
						case "#VOC_AUXFIELDS" $ VLC_Line
						    *- Nessa propriedade podem ser colocados campos auxiliares, que não serão salvos na tela
						    VLC_AuxFields = this.FOU_PropertyValue(VLC_Line, "#VOC_AUXFIELDS")
						case "#VOC_VALIDATIONKEY" $ VLC_Line
							VLC_ValidKey = this.FOU_PropertyValue(VLC_Line, "#VOC_VALIDATIONKEY")
					endcase
				endif
			endfor

			if !empty(VLC_ObjList) and !empty(VLN_SearchType) and !empty(VLC_SearchId)
				if !VLL_SearchsForAccess
					if empty(PLC_SearchId)
						VLL_AddSearch = (VLL_AllSearchs or this.FOL_SearchIn(upper(VLC_ObjName), upper(VLC_ObjList))) and;
									VLN_SearchType = PLN_SearchType
					else
						VLL_AddSearch = (PLC_SearchId == VLC_SearchId)
					endif

					VLL_AddSearch = VLL_AddSearch and !this.FOL_AcessMenu(VLC_SearchId)
				else
					VLL_AddSearch = .t.
				endif
			endif
			
			VLC_AuxDesc = " (Esp)"
			VLC_SearchCfg = ""

			if VLL_AddSearch
				VLC_SearchCfg = this.FOC_SearchProperties(VLC_SearchId, VLC_FileName)
				VLL_ValidKey = .t.
				if !empty(PLL_ChkKey)
					VLL_ValidKey = .t.
					*- Verifico se a chave é válida
					if .f. && !empty(VLC_ValidKey) and !VLL_SearchsForAccess
						VLC_Info = ""
						VLN_Action = this.VOO_Security.FON_ValidateKey(1, VLC_ValidKey, VLC_SearchCfg, @VLC_Info)
						do case
							case VLN_Action = 1 && Chave provisória válida
								VLL_ValidKey = .t.
								VLC_Expires = this.FOC_TagContent(VLC_Info, "EXP_DATE", 1)

							case VLN_Action = 2 and !empty(VLC_Info) && Chave provisória precisa ser armazenada
								VLC_FContent = filetostr(VLC_FileName)
								VLC_KeyStr = "#VOC_VALIDATIONKEY"
								*- Início da substituição da chave existente
								VLN_StartCount = at(VLC_KeyStr, VLC_FContent, 1) + len(VLC_KeyStr)
								VLN_FileLen = len(VLC_FContent)
								VLN_CharsReplaced = 0
								VLL_CharFound = .f.
								*- Procuro o final da cadeia para substituir
								for VLN_CharCounter = VLN_StartCount to VLN_FileLen
									VLC_Char = substr(VLC_FContent, VLN_CharCounter, 1)
									if inlist(VLC_Char, chr(13), chr(10))
										VLL_CharFound = .t.
										exit
									else
										VLN_CharsReplaced = VLN_CharsReplaced + 1
									endif
								endfor
								
								if VLL_CharFound
									VLC_TempKey = " = " + this.FOC_TagContent(VLC_Info, "KEY_STR", 1)
									VLC_FContent = stuff(VLC_FContent, VLN_StartCount, VLN_CharsReplaced, VLC_TempKey)
									if strtofile(VLC_FContent, VLC_FileName, 0) > 0
										VLL_ValidKey = .t.
										VLC_Expires = this.FOC_TagContent(VLC_Info, "EXP_DATE", 1)
									endif
								endif
						endcase
					endif

					VLN_RetSearchCount = VLN_RetSearchCount + 1

					if !VLL_SearchsForAccess
						if VLL_ValidKey
							*- Arquivo válido temporariamente
							VLC_AuxDesc = VLC_AuxDesc && + " (" + this.FOC_Caption("valido ate") + " " + dtoc(this.FOD_StoD(VLC_Expires)) + ")"
						else
							VLC_AuxDesc = VLC_AuxDesc + " (" + this.FOC_Caption("nao identificado") + ")"
							*- Estará bloqueado a pártir da próxima versão (4.02)
							*- VLC_SearchCfg = ""
							*- --------------------------------------------------
						endif
					
						dimension PRA_RetSearchArray[VLN_RetSearchCount, 9]
						PRA_RetSearchArray[VLN_RetSearchCount, 1] = alltrim(VLC_Description) + VLC_AuxDesc
						PRA_RetSearchArray[VLN_RetSearchCount, 2] = VLC_SearchCfg
						PRA_RetSearchArray[VLN_RetSearchCount, 3] = VLC_AuxFields
						PRA_RetSearchArray[VLN_RetSearchCount, 4] = 0
						PRA_RetSearchArray[VLN_RetSearchCount, 5] = ""
						PRA_RetSearchArray[VLN_RetSearchCount, 6] = ""
						PRA_RetSearchArray[VLN_RetSearchCount, 7] = VLN_SearchType
						PRA_RetSearchArray[VLN_RetSearchCount, 8] = VLC_SearchId
						PRA_RetSearchArray[VLN_RetSearchCount, 9] = VLC_ObjList
					else
						append blank in tmp_y34
						m.y34_001_c = ""
						m.descr = VLC_Description
						m.y34_002_n = VLN_SearchType
						m.y34_003_c = VLC_ObjList
						m.y34_009_c = transform(VLN_ShowOrder, "@L 9999")
						m.esp = .t.
						m.ukey = VLC_SearchId
						select tmp_y34
						gather memvar
					endif
				endif


				if !empty(PLC_SearchId)
					exit
				endif
			endif
		endif
	endfor

	return VLN_RetSearchCount
endproc

*/----------------------------------------------------------------------------------------
*/ Retorna as propriedades de uma pesquisa pela sua identificação
*/----------------------------------------------------------------------------------------
PROCEDURE FOC_SearchProperties
	LPARAMETERS PLC_SearchUkey, PLC_FileName

	LOCAL VLC_Return, VLL_ClientReport, VLN_CliSCount

	VLC_Return = ""
	
	if !empty(PLC_FileName) and file(PLC_FileName) && Arquivo já passado para facilitar o trabalho
		VLL_ClientReport = .t.
	else
		this.FOL_UseInPackage("y34", "y34")
	endif

	VLC_UkeyToFind = alltrim(PLC_SearchUkey)
	if !VLL_ClientReport and seek(VLC_UkeyToFind, "Y34", "UKEY")
		VLC_Return = this.FOC_SearchField()
	else

		local array VLA_SearchFile[1]
		if !VLL_ClientReport
			VLN_CliSCount = adir(VLA_SearchFile, this.VOC_ClientSearchDir + "*.sss")
		else
			VLN_CliSCount = 1
			VLA_SearchFile[1] = justfname(PLC_FileName)
		endif

		for VLN_CSCounter = 1 to VLN_CliSCount
			VLC_SearchCfg = ""
			VLC_FileName = this.VOC_ClientSearchDir + VLA_SearchFile[VLN_CSCounter, 1]
			VLC_FileContent = filetostr(VLC_FileName)
			VLC_HeaderContent = strextract(VLC_FileContent, "\\BEGIN SEARCHHEADER", "\\END SEARCHHEADER", 1, 1)
			VLL_IdFound = .f.
			if !empty(VLC_HeaderContent)
				VLN_HeaderLines = alines(VLA_Header, VLC_HeaderContent, .t.)
				VLC_SearchId = ""

				for VLN_HCounter = 1 to VLN_HeaderLines
					VLC_Line = VLA_Header[VLN_HCounter]

					if !empty(VLC_Line) and "#VOC_SEARCHID" $ VLC_Line
						VLC_SearchId = this.FOU_PropertyValue(VLC_Line, "#VOC_SEARCHID")
						VLL_IdFound = VLC_SearchId == VLC_UkeyToFind
					endif
				endfor

				if VLL_IdFound
					VLN_SStructCount = 8
					dimension VLA_SearchStruct[VLN_SStructCount, 2]
					VLA_SearchStruct[1, 1] = "\\BEGIN FIELD"
					VLA_SearchStruct[1, 2] = "\\END FIELD"
					VLA_SearchStruct[2, 1] = "\\BEGIN ORDER"
					VLA_SearchStruct[2, 2] = "\\END ORDER"
					VLA_SearchStruct[3, 1] = "\\BEGIN SQL_STRUCTURE"
					VLA_SearchStruct[3, 2] = "\\END SQL_STRUCTURE"
					VLA_SearchStruct[4, 1] = "\\BEGIN SHOWDATA"
					VLA_SearchStruct[4, 2] = "\\END SHOWDATA"
					VLA_SearchStruct[5, 1] = "\\BEGIN DATA_TREATMENT"
					VLA_SearchStruct[5, 2] = "\\END DATA_TREATMENT"
					VLA_SearchStruct[6, 1] = "\\BEGIN OLAP_CONFIG"
					VLA_SearchStruct[6, 2] = "\\END OLAP_CONFIG"
					VLA_SearchStruct[7, 1] = "\\BEGIN POINTERCFG"
					VLA_SearchStruct[7, 2] = "\\END POINTERCFG"
					VLA_SearchStruct[8, 1] = "\\BEGIN FILEINFO"
					VLA_SearchStruct[8, 2] = "\\END FILEINFO"

					if VLC_SearchId == PLC_SearchUkey
						VLC_FileContent = filetostr(VLC_FileName)

						VLC_SearchCfg = ""
						VLC_ChrSeparator = ""
						for VLN_SSCounter = 1 to VLN_SStructCount
							VLC_StrContent = strextract(VLC_FileContent, VLA_SearchStruct[VLN_SSCounter, 1],;
										VLA_SearchStruct[VLN_SSCounter, 2], 1, 1)
							if !empty(VLC_StrContent)
								VLC_SearchCfg = VLC_SearchCfg + VLC_ChrSeparator + VLA_SearchStruct[VLN_SSCounter, 1] +;
									VLC_StrContent + VLA_SearchStruct[VLN_SSCounter, 2]
								VLC_ChrSeparator = chr(13) + chr(10)
							endif
						endfor
						
						VLC_Return = VLC_SearchCfg
					endif
				endif
			endif
		endfor

	endif

	if used("y34")
		use in y34
	endif

	return VLC_Return
ENDPROC

*/----------------------------------------------------------------------------------------
*/ Retorna um objeto com as propriedades de definição de um cubo
*/----------------------------------------------------------------------------------------
PROCEDURE FOO_CubeDefinitions
	LPARAMETERS PLC_CubeCfgStr

	local VLO_Return

	VLN_StartOCCfg = at("\\BEGIN OLAP_CONFIG", PLC_CubeCfgStr)
	VLN_FinishOCCfg = at("\\END OLAP_CONFIG", PLC_CubeCfgStr) + 17

	VLC_OCCfg = ""
	if VLN_FinishOCCfg > 17 and VLN_StartOCCfg > 0 and VLN_FinishOCCfg > VLN_StartOCCfg
		VLC_OCCfg = substr(PLC_CubeCfgStr, VLN_StartOCCfg, VLN_FinishOCCfg - VLN_StartOCCfg)
	endif

	VLO_Return = .null.

	if !empty(VLC_OCCfg)
		VLO_Return = createobject("CGO_Custom")
		VLN_CfgSize = alines(VLA_CubeCfg, VLC_OCCfg, .T.)
		with VLO_Return
			.addproperty("VON_DimCount", 0)
			.addproperty("VOA_Dimensions[1]", 0)
			.addproperty("VON_LineCount", 0)
			.addproperty("VOA_Lines[1]", .null.)
			.addproperty("VON_ColumnCount", 0)
			.addproperty("VOA_Columns[1]", .null.)
			.addproperty("VON_PageCount", 0)
			.addproperty("VOA_Pages[1]", .null.)
			.addproperty("VON_DataCount", 0)
			.addproperty("VOA_Data[1]", .null.)
		endwith
		
		*- Crio os objetos e inicializo as propriedades
		for VLN_LineCounter = 1 to VLN_CfgSize
			VLN_StartName = at("\\DEFINE_DIMENSION", VLA_CubeCfg[VLN_LineCounter], 1) + 18
			if VLN_StartName > 18
				VLC_DimName = alltrim(substr(VLA_CubeCfg[VLN_LineCounter], VLN_StartName))
				with VLO_Return
					if .addobject((VLC_DimName), "CGO_Custom")
						.VON_DimCount = .VON_DimCount + 1

						*- Armazeno a referência do objeto para usar depois
						VLO_DimObj = .&VLC_DimName.
						with VLO_DimObj
							.addproperty("VOC_Expression", "")
							.addproperty("VOC_ExpressionType", "")
							.addproperty("VOC_Description", "")
							.addproperty("VON_DimType", 0)
							.addproperty("VON_GroupFooterType", 0)
							.addproperty("VON_AggregateFunc", 0)
							.addproperty("VOC_GroupFooterCaption", "")
							.addproperty("VOC_InputMask", "")
							.addproperty("VOC_Format", "")
							.addproperty("VOL_CalcOnFly", .f.)
							.addproperty("VOC_SummaryExpr", "")
						endwith

						VLN_LineCounter = VLN_LineCounter + 1
						do while !("\\END_DEFINE" $ VLA_CubeCfg[VLN_LineCounter])
							VLN_StartProperty = at("#", VLA_CubeCfg[VLN_LineCounter], 1) + 1
							VLN_EqualPosition = at("=", VLA_CubeCfg[VLN_LineCounter], 1)
							VLN_FinishProperty = VLN_EqualPosition - VLN_StartProperty
							VLC_PropertieName = alltrim(substr(VLA_CubeCfg[VLN_LineCounter], VLN_StartProperty, VLN_FinishProperty))
							VLC_Content = alltrim(substr(VLA_CubeCfg[VLN_LineCounter], VLN_EqualPosition + 1))
							VLO_DimObj.&VLC_PropertieName = eval(VLC_Content)
							VLN_LineCounter = VLN_LineCounter + 1
						enddo
						
						*- Trato algumas propriedades da dimensão para acertar a ordem de carregamento dos dados
						do case
							case VLO_DimObj.VON_DimType = 2 && Linhas
								.VON_LineCount = .VON_LineCount + 1
								dimension .VOA_Lines[.VON_LineCount]
								.VOA_Lines[.VON_LineCount] = VLO_DimObj

							case VLO_DimObj.VON_DimType = 1 && Colunas
								.VON_ColumnCount = .VON_ColumnCount + 1
								dimension .VOA_Columns[.VON_ColumnCount]
								.VOA_Columns[.VON_ColumnCount] = VLO_DimObj

							case VLO_DimObj.VON_DimType = 4 && Page
								.VON_PageCount = .VON_PageCount + 1
								dimension .VOA_Pages[.VON_PageCount]
								.VOA_Pages[.VON_PageCount] = VLO_DimObj

							case VLO_DimObj.VON_DimType = 3 && Dados
								.VON_DataCount = .VON_DataCount + 1
								dimension .VOA_Data[.VON_DataCount]
								.VOA_Data[.VON_DataCount] = VLO_DimObj

						endcase

					endif
				endwith
			endif
		endfor

		*- Carrego uma array de referências na ordem de carregamento dos objetos
		with VLO_Return
			dimension .VOA_Dimensions[.VON_DimCount]
			VLN_GenCounter = 0
			
			for VLN_OrderCounter = 1 to .VON_LineCount
				VLN_GenCounter = VLN_GenCounter + 1
				.VOA_Dimensions[VLN_GenCounter] = .VOA_Lines[VLN_OrderCounter]
			endfor

			for VLN_OrderCounter = 1 to .VON_ColumnCount
				VLN_GenCounter = VLN_GenCounter + 1
				.VOA_Dimensions[VLN_GenCounter] = .VOA_Columns[VLN_OrderCounter]
			endfor

			for VLN_OrderCounter = 1 to .VON_PageCount
				VLN_GenCounter = VLN_GenCounter + 1
				.VOA_Dimensions[VLN_GenCounter] = .VOA_Pages[VLN_OrderCounter]
			endfor

			for VLN_OrderCounter = 1 to .VON_DataCount
				VLN_GenCounter = VLN_GenCounter + 1
				.VOA_Dimensions[VLN_GenCounter] = .VOA_Data[VLN_OrderCounter]
			endfor

		endwith		
	endif
	
	return VLO_Return
endproc


*/----------------------------------------------------------------------------------------
*/ Resume um cursor convertendo os valores para a moeda solicitada, utilizado em cubos
*/----------------------------------------------------------------------------------------
procedure FOL_VMCursor
	lparameters PLC_Cursor, PLC_FieldValue, PLC_CurField, PLC_A36UkeyD, PLC_SumFields

	select (PLC_Cursor)
	VLN_FldCount = afields(VLA_Struct, PLC_Cursor)

	PLC_FieldValue = upper(alltrim(PLC_FieldValue))
	PLC_CurField = upper(alltrim(PLC_CurField))
	PLC_SumFields = upper(alltrim(PLC_SumFields))

	VLC_Comma1Sign = ""
	VLC_Comma2Sign = ""
	VLC_FldExpr = ""
	VLC_GroupExpr = ""

	for VLN_FldCounter = 1 to VLN_FldCount
		VLL_SumFld = VLA_Struct[VLN_FldCounter, 1] $ PLC_SumFields or VLA_Struct[VLN_FldCounter, 1] == PLC_FieldValue
		if VLL_SumFld
			VLC_FldExpr = VLC_FldExpr + VLC_Comma1Sign + "sum(" + VLA_Struct[VLN_FldCounter, 1] + ") as " +;
						VLA_Struct[VLN_FldCounter, 1]
			VLC_Comma1Sign = ", "
		else
			if !VLA_Struct[VLN_FldCounter, 1] == PLC_CurField
				select (PLC_Cursor)
				index on &VLA_Struct[VLN_FldCounter, 1] tag &VLA_Struct[VLN_FldCounter, 1]
				VLC_FldExpr = VLC_FldExpr + VLC_Comma1Sign + VLA_Struct[VLN_FldCounter, 1]
				VLC_GroupExpr = VLC_GroupExpr + VLC_Comma2Sign + VLA_Struct[VLN_FldCounter, 1]
				VLC_Comma1Sign = ", "
				VLC_Comma2Sign = ", "
			endif
		endif
	endfor

	select (PLC_Cursor)
	index on deleted() tag del
	index on &PLC_CurField tag currency
	set order to currency ascending
	go top in (PLC_Cursor)
	*- Preciso trabalhar com do while para otimizar a conversão de moedas
	do while !eof(PLC_Cursor)
		VLC_FldCurrency = evaluate(PLC_CurField)
		*- Se for igual não há a necessidade de conversão
		if substr(VLC_FldCurrency, 1, 5) == PLC_A36UkeyD
			select (PLC_Cursor)
			set order to currency descending
			*- Acho o último registro da mesma moeda, pulando assim todo o intervalo que não necessita de conversão
			=seek(PLC_A36UkeyD, PLC_Cursor)
			set order to currency ascending
		else
			VLN_Value = evaluate(PLC_FieldValue)
			VLD_Date = this.FOD_StoD(substr(VLC_FldCurrency, 6))
			VLN_Value = this.FON_VM(VLN_Value, VLC_FldCurrency, PLC_A36UkeyD, VLD_Date)
			replace (PLC_FieldValue) with VLN_Value in (PLC_Cursor)
		endif
		
		skip in (PLC_Cursor)
	enddo

	if !empty(VLC_GroupExpr)
		select &VLC_FldExpr from (PLC_Cursor) group by &VLC_GroupExpr into cursor tmpcconv
	else
		select &VLC_FldExpr from (PLC_Cursor) into cursor tmpcconv
		
		
	endif

	select (PLC_Cursor)
	zap in (PLC_Cursor)
	append from dbf("tmpcconv")
	VLC_TodayCurrency = substr(PLC_A36UkeyD, 1, 5) + dtos(date())
	replace all (PLC_CurField) with VLC_TodayCurrency in (PLC_Cursor)
	use in tmpcconv
endproc


*/----------------------------------------------------------------------------------------
*/ Faz a tradução de expressões contidas em um objeto e seus respectivos filhos
*/----------------------------------------------------------------------------------------
procedure FOL_TranslateObj
	lparameters PLO_ParObj
	
	local VLL_Return, VLN_MemberCount, VLN_MCounter
	local array VLA_MemberArray[1]
	
	if vartype(PLO_ParObj) == "O"
		if lower(PLO_ParObj.baseclass) = "form"
			PLO_ParObj.caption = this.FOC_Caption("form_" + lower(PLO_ParObj.name))
			if !this.VOL_LastSeek
				activate screen
				? "form_" + lower(PLO_ParObj.name)
			endif
		endif

		VLN_MemberCount = amembers(VLA_MemberArray, PLO_ParObj, 2)
		for VLN_MCounter = 1 to VLN_MemberCount
			VLO_ActualMember = PLO_ParObj.&VLA_MemberArray[VLN_MCounter]
			VLC_BaseClassName = lower(VLO_ActualMember.baseclass)

			if VLC_BaseClassName $ "|optionbutton|checkbox|label|commandbutton|page"
				*- Tradução do caption
				if !empty(VLO_ActualMember.caption)
					VLO_ActualMember.caption = this.FOC_Caption(lower(VLO_ActualMember.caption))
				endif
				*- Tradução do tooltip
				if !empty(VLO_ActualMember.tooltiptext)
					VLO_ActualMember.tooltiptext = this.FOC_Caption(lower(VLO_ActualMember.tooltiptext))
				endif

			endif

			if VLC_BaseClassName $ "|container|pageframe|optiongroup|page"
				VLL_Return = this.FOL_TranslateObj(VLO_ActualMember)
			endif
		endfor
	endif

	
	return VLL_Return
endproc


*/----------------------------------------------------------------------------------------
*/ Verifica o acesso a funcionalidades disponíveis através do pagamento de manutenção
*/----------------------------------------------------------------------------------------
function FOL_AccessMntFunction
	lparameters PLN_FunctionId, PLL_HideMsg
	local VLL_Return

	this.FOL_PushSelect()

	this.FOL_UseInPackage("y03", "y03")

	select y03
	set order to y03_001_n
	if seek(PLN_FunctionId, "Y03", "Y03_001_N")
		VLL_Return = y03.y03_006_n <= this.FON_GetMaintenance()
		if VLL_Return
			VLL_Return = val(y03.y03_005_c) = val(sys(2007, y03.y03_002_m + y03.y03_003_m + chr((y03.y03_006_n + y03.y03_001_n)%255)))
		endif
	endif

	VLC_MsgContent = y03.y03_003_m

	use in "y03"

	if !PLL_HideMsg
		if !VLL_Return and !empty(VLC_MsgContent) and !this.VOL_ServiceMode
			this.FOL_DoNormalForm("htmlviewer", VLC_MsgContent, 0)
		endif
	endif

	this.FOL_PopSelect()

	return VLL_Return
endproc

*/----------------------------------------------------------------------------------------
*/ Efetua a conversão de uma expressão caractér para DWORD (4 bytes numéricos)
*/----------------------------------------------------------------------------------------
function FON_Buf2DWord
	lparameters PLC_Buffer

	return asc(substr(PLC_Buffer, 1, 1)) + asc(substr(PLC_Buffer, 2, 1)) * 256 +;
		   asc(substr(PLC_Buffer, 3,1)) * 65536 + asc(substr(PLC_Buffer, 4,1)) * 16777216

endfunc

*/----------------------------------------------------------------------------------------
*/ Extrai uma string de Global Memory Block
*/----------------------------------------------------------------------------------------
function FOC_GetGlobalMemStr
	lparameters PLN_Addr, PLN_BaseAddr, PLC_Buffer)

    local VLN_Index, VLC_Result
    
    VLC_Result = ""

    if !PLN_Addr < PLN_BaseAddr
	    VLN_Index = PLN_Addr - PLN_BaseAddr + 1 

	    do while substr(PLC_Buffer, VLN_Index, 1) <> chr(0)
	        VLC_Result = VLC_Result + substr(PLC_Buffer, VLN_Index, 1) 
	        VLN_Index = VLN_Index + 1
	    enddo
	endif

	return VLC_Result
endfunc


*/----------------------------------------------------------------------------------------
*/ Efetua a conversão de uma expressão numérica para WORD (2 bytes char)
*/----------------------------------------------------------------------------------------
function FOC_Num2Word
	lparameters PLN_Value

	return chr(mod(PLN_Value, 256)) + chr(int(PLN_Value/256))
endfunc

*/----------------------------------------------------------------------------------------
*/ Efetua a conversão de uma expressão numérica para DWORD (4 bytes char)
*/----------------------------------------------------------------------------------------
function FOC_Num2DWord
	lparameters PLN_Value

    local VLN_0, VLN_1, VLN_2, VLN_3
    VLN_3 = int(PLN_Value/16777216)
    VLN_2 = int((PLN_Value - VLN_3*16777216)/65536)
    VLN_1 = int((PLN_Value - VLN_3*16777216 - VLN_2*65536)/256)
    VLN_0 = mod(PLN_Value, 256)

	return chr(VLN_0) + chr(VLN_1) + chr(VLN_2) + chr(VLN_3)
endfunc

*/----------------------------------------------------------------------------------------
*/ Efetua a conversão de uma expressão caractér para WORD (2 bytes numéricos)
*/----------------------------------------------------------------------------------------
function FON_Buf2Word
	lparameters PLC_Buffer

	return asc(substr(PLC_Buffer, 1,1)) + asc(substr(PLC_Buffer, 2,1)) * 256
endfunc

*/----------------------------------------------------------------------------------------
*/ Retorna um buffer de memória baseado no seu endereço
*/----------------------------------------------------------------------------------------
function FON_GetMemBuf
	lparameters PLN_Addr, PLN_Bufsize

    declare RtlMoveMemory in kernel32 as Heap2Str string @dest, integer src, integer nlength

    local VLC_Buffer
    VLC_Buffer = replicate(chr(0), PLN_Bufsize)
    =Heap2Str(@VLC_Buffer, PLN_Addr, PLN_Bufsize)

	clear dlls Heap2Str

	return VLC_Buffer
endfunc


*/----------------------------------------------------------------------------------------
*/ Retorna informações sobre o posicionamento de uma janela
*/----------------------------------------------------------------------------------------
procedure FOL_GetWindowRect
	lparameters PLN_HWnd, PRN_Width, PRN_Height, PRN_Top, PRN_Left, PRN_Bottom, PRN_Right

    declare integer GetWindowRect IN user32 INTEGER hwnd, STRING @lpRect

    LOCAL VLC_Rect, VLL_Return

    VLC_Rect = replicate(chr(0), 16)
    VLL_Return = !empty(GetWindowRect(PLN_HWnd, @VLC_Rect))

	if VLL_Return	
	    PRN_Right  = this.FON_Buf2Word(substr(VLC_Rect,  9, 4))
	    PRN_Bottom = this.FON_Buf2Word(substr(VLC_Rect, 13, 4))
	    PRN_Top    = this.FON_Buf2Word(substr(VLC_Rect,  5, 4))
	    PRN_Left   = this.FON_Buf2Word(substr(VLC_Rect,  1, 4))

	    if PRN_Left > PRN_Right
	        PRN_Left = PRN_Left - maxDword
	    endif

	    if PRN_Top > PRN_Bottom
	        PRN_Top = PRN_Top - maxDword
	    endif

	    PRN_Width = PRN_Right - PRN_Left
	    PRN_Height = PRN_Bottom - PRN_Top
	endif
	
	clear dlls GetWindowRect
	return VLL_Return
endproc

*/----------------------------------------------------------------------------------------
*/ Retorna um handle (DC) a um job de impressão
*/----------------------------------------------------------------------------------------
function FON_GetPrnDC
	lparameters PLL_DefaultPrn

	#DEFINE PD_RETURNDC         256
	#DEFINE PD_RETURNDEFAULT   1024

	declare integer PrintDlg IN comdlg32 string @lppd

    LOCAL VLC_Struct, VLN_Flags, VLN_Return
    VLN_Return = 0
	
	if !empty(PLL_DefaultPrn)
		VLN_Flags = PD_RETURNDEFAULT + PD_RETURNDC
	else
		VLN_Flags = PD_RETURNDC
	endif

    *- Preenche a estrutura PRINTDLG
    VLC_Struct = this.FOC_Num2DWord(66) + replicate(chr(0), 16) + this.FOC_Num2DWord(VLN_Flags) + replicate(chr(0), 42)
    if PrintDlg(@VLC_Struct) > 0
        VLN_Return = this.FON_Buf2DWord(substr(VLC_Struct, 17, 4))
    endif
    
    clear dlls PrintDlg
    
	return VLN_Return
endfunc

procedure FON_InitBMPBitsArray
	lparameters PLN_Width, PLN_Height, PLN_BitsPerPixel

	#DEFINE GMEM_FIXED   0
	
	declare integer GlobalAlloc in kernel32 integer wFlags, integer dwBytes
	declare RtlZeroMemory in kernel32 integer dest, integer numBytes
	
    local VLN_Ptr, VLN_AllocSize, VLN_BytesPerScan

    *- Certeza que o valor é DWORD-alinhado
    VLN_BytesPerScan = int((PLN_Width * PLN_BitsPerPixel)/8)
    if mod(VLN_BytesPerScan, 4) > 0
        VLN_BytesPerScan = VLN_BytesPerScan + 4 - mod(VLN_BytesPerScan, 4)
    endif

    VLN_AllocSize = PLN_Height * VLN_BytesPerScan
    VLN_Ptr = GlobalAlloc(GMEM_FIXED, VLN_AllocSize)
    RtlZeroMemory(VLN_Ptr, VLN_AllocSize)
	
    clear dlls GlobalAlloc, RtlZeroMemory
	
	return VLN_Ptr
endproc

function FOC_InitBitmapInfo
	lparameters PLN_TargetDC, PLN_Width, PLN_Height, PLN_BitsPerPixel

	#DEFINE BI_RGB         0
	#DEFINE RGBQUAD_SIZE   4  && RGBQUAD
	#DEFINE BHDR_SIZE     40  && BITMAPINFOHEADER

    local VLN_RgbQuadSize, VLC_RgbQuad, VLC_BIHdr

    *- Inicializando a estrutura BitmapInfoHeader
    VLC_BIHdr = this.FOC_Num2DWord(BHDR_SIZE) + this.FOC_Num2DWord(PLN_Width) + this.FOC_Num2DWord(PLN_Height) +;
        this.FOC_Num2Word(1) + this.FOC_Num2Word(PLN_BitsPerPixel) + this.FOC_Num2DWord(BI_RGB) + repli(chr(0), 20)

    *- Criando Buffer para a tabela de cores
    if PLN_BitsPerPixel = 8
        VLN_RgbQuadSize = (2^PLN_BitsPerPixel) * RGBQUAD_SIZE
        VLC_RgbQuad = replicate(chr(0), VLN_RgbQuadSize)
    else
        VLC_RgbQuad = ""
    endif
	return VLC_BIHdr + VLC_RgbQuad
endfunc


*/------------------------------------------------------------------------------------------------------------------------
*/ Lê o conteúdo de uma chave do registro
*/	PLN_HKey   : Valor da chave primaria HKEY
*/  PLC_Subkey : O valor da sub-chave de registro
*/  PLC_Entry  : A chave para se ler o valor
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_ReadRegistryKey
lparameters PLN_HKey, PLC_Subkey, PLC_Entry

	local VLN_Reserved, VLN_Type, VLC_Data, VLN_LenData
	store 0 to VLN_Reserved, VLN_Type
	store space(256) to VLC_Data
	store len(VLC_Data) to VLN_LenData

	VLN_RegHandle = 0

	VLN_Result = RegOpenKey(PLN_HKey, PLC_Subkey, @VLN_RegHandle)
	if VLN_Result # ERROR_SUCCESS
		return ""
	endif

	VLN_Result = RegQueryValueEx(VLN_RegHandle, PLC_Entry, VLN_Reserved, @VLN_Type, @VLC_Data, @VLN_LenData)

	* Check for error 
	if VLN_Result # ERROR_SUCCESS
		return ""
	endif

	* Make sure we have a data string data type
	if VLN_Type # REG_SZ
		return ""
	endif

	RegCloseKey(VLN_RegHandle)

	return left(VLC_Data, VLN_LenData-1)
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Converte datahora no formato WWW, DD MMM YYYY HH:MM:SS -FFFF para datahora no formato VFP
*/ Ex: Thu, 30 Jan 2003 15:29:32 -0300 ==> {30/01/2003 15:29:32}
*/------------------------------------------------------------------------------------------------------------------------
procedure FOT_ConvertLongToDate
	lparameters PLC_LongDate

	local VLT_Return, VLC_CleanStr

	if substr(PLC_LongDate, 4, 1) = ","
		*- Caso o dia da semana tenha sido passado, retiro da string
		PLC_LongDate = substr(PLC_LongDate, 5)
	endif

	VLC_CleanStr = chrtran(PLC_LongDate, ",-; ", "::::")
	VLC_CleanStr = strtran(VLC_CleanStr, "::", ":")
	VLN_LineCount = alines(VLA_DateContent, VLC_CleanStr, .t., ":")

	VLT_Return = {:}

	if VLN_LineCount >= 7
		VLC_MonthStr = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
		VLN_Day = val(VLA_DateContent[1])
		VLC_Month = upper(VLA_DateContent[2])
		VLN_Month = int(at(VLC_Month, VLC_MonthStr)/3)+1
		VLN_Year = val(VLA_DateContent[3])
		VLN_Hour = val(VLA_DateContent[4])
		VLN_Min = val(VLA_DateContent[5])
		VLN_Sec = val(VLA_DateContent[6])

		VLT_Return = datetime(VLN_Year, VLN_Month, VLN_Day, VLN_Hour, VLN_Min, VLN_Sec)
	endif

	return VLT_Return
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retira um arquivo do executável, colocando-o em um diretório especificado
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_ExtractFileFromExe
	lparameters PLC_FileToExtract, PLC_FullPath, PLC_DestinFile

	local VLN_Return
	VLN_Return = 0

	VLC_FileName = justfname(PLC_FileToExtract)

	if empty(PLC_FullPath)
		PLC_FullPath = addbs(this.VOC_UserTmp)
	endif

	if empty(PLC_DestinFile)
		PLC_DestinFile = PLC_FileToExtract
	endif

	if !empty(PLC_FileToExtract) and file(PLC_FileToExtract)
		VLC_FileContent = filetostr(PLC_FileToExtract)
		if !empty(VLC_FileContent)
			VLC_NewFileLoc = (PLC_FullPath + PLC_DestinFile)
			VLN_Return = strtofile(VLC_FileContent, VLC_NewFileLoc)
		endif
	endif

	return VLN_Return
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Limpa o diretório de arquivos temporários
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_CleanTmpDir
	lparameters PLC_SubDirectory

	local VLN_TmpFileCount, VLN_Counter, VLN_FHandle, VLC_TmpDir, VLL_RemoveMainDir
	local array VLA_TmpFiles[1]

	if empty(PLC_SubDirectory)
		if empty(this.VOC_DirSettings)
			return
		endif

		VLC_TmpDir = addbs(this.VOC_DirSettings)
		VLL_RemoveMainDir = .t.

			
		if (this.VON_DbType = 2 and !(this.FOC_StartPar("RESOURCE") == 'DBSERVER'))  
			VLL_RemoveMainDir = .F.
		else
			VLL_RemoveMainDir = .t.
		endif


		try
			delete file (addbs(this.VOC_UserTmp) + '*.*')
			
			if VLL_RemoveMainDir 
				delete file (VLC_TmpDir + '*.*')
			endif
			
			catch to VLO_Error
		endtry

	else
		VLC_TmpDir = addbs(PLC_SubDirectory)
	endif

	if !empty(VLC_TmpDir)
	
		VLN_TmpFileCount = adir(VLA_TmpFiles, VLC_TmpDir + "*.*", "SHD")

		for VLN_Counter = 1 to VLN_TmpFileCount
			if !(VLA_TmpFiles[VLN_Counter, 1] == "." or VLA_TmpFiles[VLN_Counter, 1] == "..")
				this.FOL_SetFileattributes(VLC_TmpDir + VLA_TmpFiles[VLN_Counter, 1], "-R")
				if at("A", VLA_TmpFiles[VLN_Counter, 5]) > 0
					try
						if (VLC_TmpDir = addbs(this.VOC_DirSettings)) and ;
							(this.VON_DbType = 2 and !(this.FOC_StartPar("RESOURCE") == 'DBSERVER')) 
						else
							delete file (VLC_TmpDir + VLA_TmpFiles[VLN_Counter, 1])
						endif
						
						catch to VLO_Error
					endtry
				else
					if at("D", VLA_TmpFiles[VLN_Counter, 5]) > 0
						this.FOL_CleanTmpDir(VLC_TmpDir + VLA_TmpFiles[VLN_Counter, 1])
						try
							rmdir (VLC_TmpDir + VLA_TmpFiles[VLN_Counter, 1])
							catch to VLO_Error
						endtry
					endif
				endif
			endif
		endfor
	endif


	if VLL_RemoveMainDir
		try
			rmdir (this.VOC_DirSettings)
			catch to VLO_Error
		endtry
	endif
	
	return
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna uma cor do sistema (Windows)
*/ COLOR_SCROLLBAR         0					&& Scroll bar gray area.
*/ COLOR_BACKGROUND        1					&& Desktop.
*/ COLOR_ACTIVECAPTION     2					&& Active window title bar. Windows 98/Me, Windows 2000 or later: Specifies the left side color in the color gradient of an active window's title bar if the gradient effect is enabled.
*/ COLOR_INACTIVECAPTION   3					&& Inactive window caption. Windows 98/Me, Windows 2000 or later: Specifies the left side color in the color gradient of an inactive window's title bar if the gradient effect is enabled.
*/ COLOR_MENU              4					&& Menu background.
*/ COLOR_WINDOW            5					&& Window background.
*/ COLOR_WINDOWFRAME       6					&& Window frame.
*/ COLOR_MENUTEXT          7					&& Text in menus.
*/ COLOR_WINDOWTEXT        8					&& Text in windows.
*/ COLOR_CAPTIONTEXT       9					&& Text in caption, size box, and scroll bar arrow box.
*/ COLOR_ACTIVEBORDER      10					&& Active window border.
*/ COLOR_INACTIVEBORDER    11					&& Inactive window border.
*/ COLOR_APPWORKSPACE      12					&& Background color of multiple document interface (MDI) applications.
*/ COLOR_HIGHLIGHT         13					&& Item(s) selected in a control.
*/ COLOR_HIGHLIGHTTEXT     14					&& Text of item(s) selected in a control.
*/ COLOR_BTNFACE           15					&& Face color for three-dimensional display elements and for dialog box backgrounds.
*/ COLOR_BTNSHADOW         16					&& Shadow color for three-dimensional display elements (for edges facing away from the light source).
*/ COLOR_GRAYTEXT          17					&& Grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color.
*/ COLOR_BTNTEXT           18					&& Text on push buttons.
*/ COLOR_INACTIVECAPTIONTEXT 19				&& Color of text in an inactive caption.
*/ COLOR_BTNHIGHLIGHT      20					&& Highlight color for three-dimensional display elements (for edges facing the light source.)
*/ COLOR_3DDKSHADOW        21					&& Dark shadow for three-dimensional display elements.
*/ COLOR_3DLIGHT           22					&& Light color for three-dimensional display elements (for edges facing the light source.)
*/ COLOR_INFOTEXT          23					&& Text color for tooltip controls.
*/ COLOR_INFOBK            24					&& Background color for tooltip controls.
*/ COLOR_HOTLIGHT                  26			&& Windows 98/Me, Windows 2000 or later: Color for a hot-tracked item. Single clicking a hot-tracked item executes the item.
*/ COLOR_GRADIENTACTIVECAPTION     27			&& Windows 98/Me, Windows 2000 or later: Right side color in the color gradient of an active window's title bar. COLOR_ACTIVECAPTION specifies the left side color. Use SPI_GETGRADIENTCAPTIONS with the SystemParametersInfo function to determine whether the gradient effect is enabled.
*/ COLOR_GRADIENTINACTIVECAPTION   28			&& Windows 98/Me, Windows 2000 or later: Right side color in the color gradient of an inactive window's title bar. COLOR_INACTIVECAPTION specifies the left side color. 
*/ COLOR_MENUHILIGHT				29			&& Whistler: The color used to highlight menu items when the menu appears as a flat menu (see SystemParametersInfo). The highlighted menu item is outlined with COLOR_HIGHLIGHT.
*/ COLOR_MENUBAR					30			&& Whistler: The background color for the menu bar when menus appear as flat menus (see SystemParametersInfo). However, COLOR_MENU continues to specify the background color of the menu popup.
*/ COLOR_DESKTOP           COLOR_BACKGROUND
*/ COLOR_3DFACE            COLOR_BTNFACE
*/ COLOR_3DSHADOW          COLOR_BTNSHADOW
*/ COLOR_3DHIGHLIGHT       COLOR_BTNHIGHLIGHT
*/ COLOR_3DHILIGHT         COLOR_BTNHIGHLIGHT
*/ COLOR_BTNHILIGHT        COLOR_BTNHIGHLIGHT
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_GetSystemColor
	lparameters PLN_ObjIndex

	return this.VOA_SystemColors[PLN_ObjIndex + 1]
endproc

*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOU_getHproperties		  
*/ Parâmetros  		: PLC_PropertieName - Nome da propriedade 
*/					  PLU_MoreInfo - Parametro adicoinal para utilizacao no CASE alternativo																	  
*/ Descrição   		: Permite a leitura das propriedade Hidden desta classe, para isso estas propriedadespropriedades
*/					  nao devem constar na propriedade voc_NeverGetThisProperties
*/ Última alteração : 15/04/2003                                                
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
procedure FOU_GetHproperties
lparameters PLC_PropertieName, PLU_MoreInfo

Local VLN_ProgLevel 

	*	Nao permite que esta funcao seja chamada de Ajustes !
	*
	VLN_ProgLevel = 0
	Do While !Empty(PROGRAM(VLN_ProgLevel ))
		If Lower( PROGRAM(VLN_ProgLevel )) = "fol_triggers"
			Return -2
			Exit
		Endif
		VLN_ProgLevel = VLN_ProgLevel + 1

	Enddo
	
	If Vartype(PLC_PropertieName) = "C" And Atc(PLC_PropertieName, this.voc_NeverGetThisProperties) = 0
		If Pemstatus(this,PLC_PropertieName,5)
			Return Getpem( this, PLC_PropertieName)
		Else
			PLC_PropertieName = Lower(PLC_PropertieName)
			
			Do case
				Case PLC_PropertieName = "voa_licence[]" And !Empty(PLU_MoreInfo)
					Return this.VOA_Licence[PLU_MoreInfo] 
				Otherwise
					Return -1
			endcase
		Endif
	Endif
	
	Return -1 
	
Endproc

*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOU_setHproperties		  
*/ Parâmetros  		: PLC_PropertieName - Nome da propriedade 
*/					  PLN_ChkSum  - ChkSUm de verificacao
*/					  PLU_Value - Novo Valor a ser atribuido
*/					  PLU_MoreInfo - Parametro adicoinal para utilizacao no CASE alternativo																	  
*/ Descrição   		: Permite a alteraca das propriedade Hidden desta classe, para isso estas propriedades
*/					  nao devem constar na propriedade voc_NeverGetThisProperties.
*/					  Antes de passar os valores desejados para esta funcao, e necessario calcular o CHKsum
*/					  da propriedade, conforme descrito no codigo.
*/ Última alteração : 15/04/2003                                                
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
procedure FOU_SetHproperties
lparameters PLC_PropertieName, PLN_ChkSum, PLU_Value, PLU_MoreInfo

Local VLN_return, VLN_ProgLevel 
VLN_return = -1

	*	Nao permite que esta funcao seja chamada de Ajustes !
	*
	VLN_ProgLevel = 0
	Do While !Empty(PROGRAM(VLN_ProgLevel ))
		If Lower( PROGRAM(VLN_ProgLevel )) = "fol_triggers"
			Return -2
			Exit
		Endif
		VLN_ProgLevel = VLN_ProgLevel + 1

	Enddo

	If 	Vartype(PLC_PropertieName) = "C" And Atc(PLC_PropertieName, this.voc_NeverGetThisProperties) = 0 And ;
		Vartype(PLN_ChkSum		 ) = "C" And ;
		PLN_ChkSum = Sys(2007,PLC_PropertieName + "StarSoft-Internal Control")
		If Pemstatus(this,PLC_PropertieName,5)
			this.&PLC_PropertieName = PLU_Value
			VLN_return = 1

		Else
			PLC_PropertieName = Lower(PLC_PropertieName)
			
			Do case
				Case PLC_PropertieName = "voa_licence[]" And !Empty(PLU_MoreInfo)
					this.VOA_Licence[PLU_MoreInfo] = PLU_Value
					VLN_return = 1
					
				Otherwise
			endcase
		Endif
	Endif
	
	Return VLN_return 
	
Endproc



*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_DirInfo		  
*/ Parâmetros  		: PLC_Dir - Rais do Diretorio 
*/					  PRN_Folders -  Numero de Folders encontrados
*/					  PLL_SubFolders - Processa recursivamente todos os subs direotrios
*/					  PRN_Files -  Numero de arquivos encontrados
*/					  PRN_Bytes -  Total de bytes nos arquivos encontrados
*/					  PRA_allFolders - Array contendo a lista de todos os DIretorios encontrados
*/					  PRA_allFiles - Array contendo informacoes sobre todos os arquivos encontrados
*/					  PLC_NotTHisType - Tipos de arquivos NAO permitidos
*/					  
*/ Descrição   		: Le diversas informacoes sobre um determinado diretori, util para analisar o 
*/					  conteudo do mesmo antes de compactar os arquivos e guardar no sistema.
*/					  
*/ Última alteração : 25/04/2003                                                
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
Procedure FOL_DirInfo
Lparameters PLC_Dir, PLL_SubFolders, PRN_Folders, PRN_Files, PRN_Bytes, PRA_allFolders, PRA_allFiles, PLC_NotTHisType

Local VLN_Files, Vln_i, VLC_Name, VLC_Attrib, VLN_TotalBytes
Local Array VLA_Dir(1)

	If Empty(PLC_NotTHisType)
		PLC_NotTHisType= "D"
	endif

	VLN_Files = Adir( VLA_Dir, PLC_Dir + "\*.*","D" )	&& "D" e obrigatorio
	VLN_TotalBytes = 0

	For Vln_i = 1 To VLN_Files

		VLC_Name   = VLA_Dir(Vln_i ,1)
		VLC_Size   = VLA_Dir(Vln_i, 2)
		VLC_Attrib = VLA_Dir(Vln_i ,5)


		If 	!(("A" $ VLC_Attrib And ("A" $ PLC_NotTHisType)) or ;
			("H" $ VLC_Attrib And ("H" $ PLC_NotTHisType)) or ;
			("R" $ VLC_Attrib And ("R" $ PLC_NotTHisType)) or ;
			("D" $ VLC_Attrib And ("D" $ PLC_NotTHisType)) or ;
			("S" $ VLC_Attrib And ("S" $ PLC_NotTHisType)) )
		
			PRN_Bytes = PRN_Bytes + VLC_Size

			PRN_Files = PRN_Files + 1
			Dimension PRA_allFiles(PRN_Files ,5 )
			PRA_allFiles(PRN_Files ,1 ) = PLC_Dir+"\"+VLA_Dir(Vln_i ,1)
			PRA_allFiles(PRN_Files ,2 ) = VLA_Dir(Vln_i ,2)
			PRA_allFiles(PRN_Files ,3 ) = VLA_Dir(Vln_i ,3)
			PRA_allFiles(PRN_Files ,4 ) = VLA_Dir(Vln_i ,4)
			PRA_allFiles(PRN_Files ,5 ) = VLA_Dir(Vln_i ,5)
			
		Else

			If "D" $ VLC_Attrib And VLC_Name != "." And PLL_SubFolders = .t.
			
				PRN_Folders = PRN_Folders+ 1
				Dimension PRA_allFolders(PRN_Folders,1)
				PRA_allFolders(PRN_Folders,1) = PLC_Dir+"\"+VLC_Name
				
				this.FOL_DirInfo(PLC_Dir+"\"+VLC_Name, PLL_SubFolders, @PRN_Folders, @PRN_Files,;
									 @PRN_Bytes, @PRA_allFolders, @PRA_allFiles, PLC_NotTHisType)
				
			endif

		endif
	Next
	
	Return VLN_TotalBytes
	
Endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Traz a data e hora do servidor
*/------------------------------------------------------------------------------------------------------------------------
procedure FOD_SrvDateTime
	lparameters PLN_Type

	local VLT_Return, VLL_Ret
	
	VLT_Return = datetime()
	
	do case
		case this.VON_DBType = 1 && Sql Server
			VLL_Ret = this.FOL_SqlExec("select getdate() as datahora", "tmpdt")
			if VLL_Ret
				VLT_DT = tmpdt.datahora
			endif

		case this.VON_DBType = 2 && Oracle
			VLL_Ret = this.FOL_SqlExec("select to_char(sysdate, 'YYYYMMDDHH24MISS') as datahora from dual", "tmpdt")
			if VLL_Ret
				VLC_CharDT = alltrim(tmpdt.datahora)
				VLT_DT = datetime(val(substr(VLC_CharDT, 1, 4)), val(substr(VLC_CharDT, 5, 2)),;
						 val(substr(VLC_CharDT, 7, 2)), val(substr(VLC_CharDT, 9, 2)), val(substr(VLC_CharDT, 11, 2)),;
						 val(substr(VLC_CharDT, 13, 2)))
			endif
			
		case this.VON_DBType = 3 && DB2
			VLL_Ret = this.FOL_SqlExec("select current timestamp as datahora from sysibm.sysdummy1", "tmpdt")
			if VLL_Ret
				VLT_DT = tmpdt.datahora
			endif
			
	endcase
	
	if VLL_Ret
		do case
			case empty(PLN_Type)
				VLT_Return = datetime(year(VLT_DT), month(VLT_DT), day(VLT_DT), hour(VLT_DT), minute(VLT_DT), sec(VLT_DT))
			case PLN_Type = 1
				VLT_Return = date(year(VLT_DT), month(VLT_DT), day(VLT_DT))
		endcase
	endif
	
	if used("tmpdt")
		use in tmpdt
	endif

	return VLT_Return
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Apaga todos os valores armazenados no array VOA_ReportResults
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_ResetReportResults

with this
	dimension .VOA_ReportResults[1]
	.VOA_ReportResults = .f.
	.VON_ReportResultsCount = 0
endwith
return .t.

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Adiciona um elemento no array VOA_ReportResults
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_AddReportResult
lparameters PLC_ReportResult

local VLL_Empty

VLL_Empty = empty(PLC_ReportResult)

if !VLL_Empty
	with this
		.VON_ReportResultsCount = .VON_ReportResultsCount + 1
		dimension .VOA_ReportResults[.VON_ReportResultsCount]
		.VOA_ReportResults[.VON_ReportResultsCount] = PLC_ReportResult
	endwith
endif

return !VLL_Empty

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna a próxima data e hora baseado na string de configuração de recorrência
*/------------------------------------------------------------------------------------------------------------------------
procedure FOT_NextExecDateTime
	lparameters PLC_CfgStr, PLT_LastExecution

	VLT_Return = { :}

	VLC_StartDate = this.FOC_TagContent(PLC_CfgStr, "start_date", 1)
	VLD_StartDate = this.FOD_StoD(VLC_StartDate)
	VLC_FinishType = this.FOC_TagContent(PLC_CfgStr, "finish_type", 1)
	VLC_ExecTime = this.FOC_TagContent(PLC_CfgStr, "exec_time", 1)
	VLN_ExecHour = val(substr(VLC_ExecTime, 1, 2))
	VLN_ExecMin = val(substr(VLC_ExecTime, 4, 2))
	VLC_IntervalType = this.FOC_TagContent(PLC_CfgStr, "interval_type", 1)
	VLC_FinishDate = this.FOC_TagContent(PLC_CfgStr, "finish_date", 1)

	if empty(nvl(PLT_LastExecution, { :}))
		VLL_1stTime = .t.
		VLD_BaseDate = date(year(VLD_StartDate), month(VLD_StartDate), day(VLD_StartDate))
	else
		VLL_1stTime = .f.
		VLD_BaseDate = ttod(PLT_LastExecution)
	endif

	do case
		*- Diário
		case VLC_IntervalType = "1"
			VLC_DayInterval = this.FOC_TagContent(PLC_CfgStr, "day_interval", 1)
			VLC_DayType = this.FOC_TagContent(PLC_CfgStr, "day_type", 1)
			if VLC_DayType = "1"
				*- Intervalo de dias
				if !VLL_1stTime
					VLD_BaseDate = VLD_BaseDate + val(VLC_DayInterval)
				endif
			else
				*- Dias de semana
				if !VLL_1stTime
					VLD_BaseDate = VLD_BaseDate + 1
				endif
				do while inlist(dow(VLD_BaseDate), 1, 7)
					VLD_BaseDate = VLD_BaseDate + 1
				enddo
			endif

		*- Semanal
		case VLC_IntervalType = "2"
			VLN_WeekInterval = val(this.FOC_TagContent(PLC_CfgStr, "week_interval", 1))
			VLC_WeekDays = this.FOC_TagContent(PLC_CfgStr, "week_days", 1)
			VLL_Found = .f.

			*- Proteção para evitar loop infinito
			if "1" $ VLC_WeekDays
				VLN_DayCount = occurs("1", VLC_WeekDays)
				VLN_LastDay = at("1", VLC_WeekDays, VLN_DayCount)
				VLL_LastJob = (dow(VLD_BaseDate) = VLN_LastDay)

				if !VLL_1stTime and !VLL_LastJob
					VLD_BaseDate = VLD_BaseDate + 1
				endif

				VLN_Dow = dow(VLD_BaseDate)

				if VLL_LastJob
					*- Avanço até o fim da semana
					VLD_BaseDate = VLD_BaseDate + (7 - VLN_Dow) + 1
				endif

				if VLN_WeekInterval > 1 and VLL_LastJob
					*- Somo as semanas
					VLD_BaseDate = VLD_BaseDate + ((VLN_WeekInterval - 1) * 7)
				endif

				do while !VLL_Found
					if substr(VLC_WeekDays, dow(VLD_BaseDate), 1) == "1"
						VLL_Found = .t.
						exit
					endif
					VLD_BaseDate = VLD_BaseDate + 1
				enddo
			endif

		*- Mensal		
		case VLC_IntervalType = "3"
			VLN_MonthInterval = val(this.FOC_TagContent(PLC_CfgStr, "month_interval", 1))
			VLN_MonthDay = val(this.FOC_TagContent(PLC_CfgStr, "month_day", 1))
			VLD_BaseDate = date(year(VLD_BaseDate), month(VLD_BaseDate), VLN_MonthDay)

			if !VLL_1stTime
				do while ttod(PLT_LastExecution) >= VLD_BaseDate
					VLD_BaseDate = gomonth(VLD_BaseDate, VLN_MonthInterval)
				enddo
			else
				do while VLD_StartDate > VLD_BaseDate
					VLD_BaseDate = gomonth(VLD_BaseDate, VLN_MonthInterval)
				enddo
			endif

		*- Anual
		case VLC_IntervalType = "4"
			VLN_YearMonth = val(this.FOC_TagContent(PLC_CfgStr, "year_month", 1))
			VLN_YearDay = val(this.FOC_TagContent(PLC_CfgStr, "year_day", 1))
			VLD_BaseDate = date(year(VLD_BaseDate), VLN_YearMonth, VLN_YearDay)

			if !VLL_1stTime
				do while ttod(PLT_LastExecution) >= VLD_BaseDate
					VLD_BaseDate = gomonth(VLD_BaseDate, 12)
				enddo
			else
				do while VLD_StartDate > VLD_BaseDate
					VLD_BaseDate = gomonth(VLD_BaseDate, 12)
				enddo
			endif

		*- Execução única
		case VLC_IntervalType = "5"
			VLC_DateTime = this.FOC_TagContent(PLC_CfgStr, "exec_once", 1)

			if VLL_1stTime
				VLT_Return = ctot(trans(VLC_DateTime, "@R 9999-99-99T99:99:99"))
			else
				VLT_Return = { :}
			endif
			
	endcase

	if VLC_IntervalType <> "5"
		VLT_Return = datetime(year(VLD_BaseDate), month(VLD_BaseDate), day(VLD_BaseDate), VLN_ExecHour, VLN_ExecMin, 0)
		if VLC_FinishType = "2" and !empty(VLC_FinishDate)
			if ttod(VLT_Return) > this.FOD_StoD(VLC_FinishDate)
				VLT_Return = { :}
			endif
		endif
	endif

	return VLT_Return
endproc
	
*/------------------------------------------------------------------------------------------------------------------------
*/ Armazena um arquivo qualquer no T37
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_StoreFile
	lparameters PLC_FullFileName, PLC_Par, PLC_Ukeyp, PLN_Type

	local VLC_Return, VLC_FContent
	
	VLC_Return = ""

	if !empty(PLC_FullFileName) and file(PLC_FullFileName)
		VLC_FContent = filetostr(PLC_FullFileName)

		with this
			.FOL_SetParameter(1, space(3))
			.FOL_SetParameter(2, space(20))
			if .FOL_EditCursor("T37_PARUKEYP", "T37_PARUKEYP", "T37T", "Z", 0, 2)
				select t37t
				.FOL_NewRecordToForm("t37t")
			
				replace t37t.t37_001_m	with VLC_FContent,;
						t37t.t37_002_n	with iif(empty(PLN_Type), 0, PLN_Type),;
						t37t.t37_003_m	with PLC_FullFileName,;
						t37t.t37_004_c 	with justext(PLC_FullFileName),;
						t37t.t37_par	with PLC_Par,;
						t37t.t37_ukeyp	with PLC_Ukeyp,;
						t37t.mycontrol 	with "1",;
						t37t.usr_ukey 	with .VOC_UserCode;
				in t37t

				.FOL_CreateSQLString("T37T", "T37")

				if .FOL_SaveRecordSql("T37T", "T37", "UKEY")
					VLC_Return = t37t.ukey
				endif
				use in t37t
			endif
		endwith
	endif

	return VLC_Return
	
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Converte uma picture numérica do VFP para o padrão do VB
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_FormatVFP2VB
	lparameters PLC_VFPPict

	local VLC_Return
	
	VLC_Return = alltrim(PLC_VFPPict)
	VLN_PointLoc = at(".", VLC_Return)
	VLN_PictLen = len(VLC_Return)
	
	if VLN_PointLoc > 0
		VLC_AfterPoint = substr(VLC_Return, VLN_PointLoc)
		VLC_AfterPoint = chrtran(VLC_AfterPoint, "9", "0")

		if VLN_PictLen+2 > len(VLC_AfterPoint)
			VLC_BeforePoint = left(VLC_Return, VLN_PointLoc-2) + "0"
		else
			VLC_BeforePoint = "0"
		endif

	else
		VLC_AfterPoint = ""
		if VLN_PictLen > 1
			VLC_BeforePoint = left(VLC_Return, VLN_PictLen - 1) + "0"
		else
			VLC_BeforePoint = "0"
		endif
	endif

	VLC_BeforePoint = chrtran(VLC_BeforePoint, "9", "#")
	VLC_Return = VLC_BeforePoint + VLC_AfterPoint

	return VLC_Return
	
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ 
*/------------------------------------------------------------------------------------------------------------------------
*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_ListInLIst		  
*/ Parâmetros  		: PLC_List1- Lista de itens Origem 
*/					  PLC_List2- Lista de itens Destino
*/					  PLC_Separator- Separador dos itens nas listas
*/					  PLL_PreserveCase-  Ignora Maiusculo e Minusculo
*/					  
*/ Descrição   		: Verifica se algum dos elementos da primeira lista esta contida na Segunda Lista
*/					  
*/ Última alteração : 25/06/2003                                                
*/ Alterado por 	: WLC				                						
*/ Versão 			: 1 														
*/----------------------------------------------------------------------------------------/*
procedure FOL_ListInLIst
	Lparameters PLC_List1, PLC_List2, PLC_Separator, PLL_PreserveCase
	
	If Empty(PLC_List1 ) Or Empty(PLC_List1 )
		Return .f.
	endif
	If Empty(PLC_Separator)
		PLC_Separator = ","
	Endif

	If !PLL_PreserveCase
		PLC_List1 = Lower(PLC_List1)
		PLC_List2 = Lower(PLC_List2)
	endif

	VLN_Occurs = Occurs(PLC_Separator, PLC_List1)+1
	For VLN_i = 1 To VLN_Occurs

		VLC_Item = Strextract( PLC_Separator + PLC_List1 + PLC_Separator, PLC_Separator,PLC_Separator, VLN_i, 1 )
		
		If (PLC_Separator + VLC_Item + PLC_Separator) $ (PLC_Separator + PLC_List2 + PLC_Separator)
			Return .t.
		endif

	Next

	Return .F.

Endproc 

*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna codigo para utilização nos Servicepack
*/------------------------------------------------------------------------------------------------------------------------
procedure FOC_SS
	return this.VOC_SecurityCodeSP
endproc



*/----------------------------------------------------------------------------------------/*
*/ Função Objeto    : FOC_OperateNumber                                                   /*
*/ Parâmetros  		: PLC_Operand   : Operando                                            /*
*/               	  PLC_Operator  : Operador                                            /*
*/               	  PLC_What   	: Operação a ser executada (*,/,+,-)                  /*
*/                    PLN_Decimals  : Número de casas decimais                            /*  
*/ Descrição   		: Executa uma operação de números gigantes                            /*
*/ Retorno 			: String contendo o resultado                                         /*
*/ Última alteração : 26/08/03                                                            /*
*/ Alterado por		: Denis Carrizo													      /*
*/----------------------------------------------------------------------------------------/*
procedure FOC_OperateNumber
lparameters PLC_Operand, PLC_Operator, PLC_What, PLN_Decimals

local VLC_Ret, VLN_Sign1, VLN_Sign2, VLN_Dec1, VLN_Dec2, VLN_i, VLN_Len1, VLN_Len2, VLN_Rest, ;
		VLC_Zeros, VLN_c1, VLN_c2, VLC_Dividend, VLN_MaxLen

VLC_Ret = ""
PLC_Operand = this.FOC_RetSignif(alltrim(PLC_Operand), .T., .T.)
PLC_Operator = this.FOC_RetSignif(alltrim(PLC_Operator), .T., .T.)

if empty(PLN_Decimals)
	PLN_Decimals = 0
endif	

*- Guarda os sinais do operador e do operando
VLN_Sign1 = 1
if left(PLC_Operand, 1) = "-"
	VLN_Sign1 = -1
	PLC_Operand = substr(PLC_Operand, 2)
endif	
if PLC_Operand == "0"
	VLN_Sign1 = 0
endif
	
VLN_Sign2 = 1
if left(PLC_Operator, 1) = "-"
	VLN_Sign2 = -1
	PLC_Operator = substr(PLC_Operator, 2)
endif
if PLC_Operator == "0"
	VLN_Sign2 = 0
endif

*- Verifica o número de casas decimais do operando
VLN_Dec1 = 0
VLN_i = at(".", PLC_Operand)
if VLN_i > 0
	VLN_Dec1 = len(substr(PLC_Operand, VLN_i + 1))
endif		

*- Verifica o número de casas decimais do operador
VLN_Dec2 = 0
VLN_i = at(".", PLC_Operator)
if VLN_i > 0
	VLN_Dec2 = len(substr(PLC_Operator, VLN_i + 1))
endif		

*- Ajusta as casas decimais para os dois números
VLN_i = max(VLN_Dec1, VLN_Dec2)
PLC_Operand = strtran(padr(PLC_Operand, len(PLC_Operand) + VLN_i - VLN_Dec1, "0"), ".", "")
PLC_Operand = this.FOC_RetSignif(PLC_Operand, .T., .F.)
PLC_Operator = strtran(padr(PLC_Operator, len(PLC_Operator) + VLN_i - VLN_Dec2, "0"), ".", "")
PLC_Operator = this.FOC_RetSignif(PLC_Operator, .T., .F.)
VLN_Dec1 = VLN_i

VLN_len1 = len(PLC_Operand)
VLN_len2 = len(PLC_Operator)
VLN_maxlen = max(VLN_len1, VLN_len2)

do case
	*-	Multiplicação
	case PLC_What = "*"
		VLC_Zeros = ""
		if VLN_Sign1 <> 0 and VLN_Sign2 <> 0
			for VLN_c2 = VLN_len2 to 1 step -1
				VLN_n2 = val(subs(PLC_Operator, VLN_c2, 1))
				if VLN_n2 > 0
					VLC_Aux = ""
					VLN_Rest = 0
					for VLN_c1 = VLN_len1 to 1 step -1
						VLN_n1 = val(subs(PLC_Operand, VLN_c1, 1)) * VLN_n2 + VLN_Rest
						if VLN_n1 > 9
							VLN_Rest = int(VLN_n1 / 10)
							VLN_n1 = VLN_n1 - (VLN_Rest * 10)
						else
							VLN_Rest = 0	
						endif	
						VLC_Aux = str(VLN_n1,1) + VLC_Aux
					endfor
					if VLN_Rest > 0
						VLC_Aux = str(VLN_Rest,1) + VLC_Aux
					endif
					VLC_Ret = this.FOC_SumSubNumber(VLC_Ret, VLC_Aux+VLC_Zeros, "+")
				endif
				VLC_Zeros = VLC_Zeros + "0"
			endfor
		else
			VLC_Ret = "0"
		endif
*---	ajusta o ponto decimal
		VLC_Ret = padl("", 2 * VLN_Dec1 + 1, "0") + VLC_Ret
		VLN_i = len(VLC_Ret)
		if VLN_Dec1 > 0	
			VLN_i = VLN_i - (2 * VLN_Dec1)
			VLC_Ret = substr(VLC_Ret, 1, VLN_i) + "." + substr(VLC_Ret, VLN_i + 1)
		endif	
		if (VLN_Sign1 * VLN_Sign2) < 0
			VLC_Ret = "-"+VLC_Ret
		endif	

*--	divisão 
	case PLC_What = "/"
		PLC_Operand = PLC_Operand + padl("", PLN_Decimals + 5, "0")
		VLN_len1 = len(PLC_Operand)
		VLC_Dividend = substr(PLC_Operand, 1, VLN_len2)
		if VLN_Sign1 <> 0 and VLN_Sign2 <> 0
			for VLN_c1 = len(VLC_Dividend) + 1 to VLN_len1
				if len(VLC_Dividend) >= VLN_len2
					VLC_Aux	= VLC_Dividend
					for VLN_n1 = 0 to 9
						VLC_Aux = this.FOC_SumSubNumber(VLC_Aux, PLC_Operator, "-")
						if VLC_Aux = "-"
							exit
						endif
						VLC_Dividend = VLC_Aux	
					endfor		
				else
					VLN_n1 = 0	
				endif
				VLC_Ret = VLC_Ret + str(VLN_n1, 1)
				VLC_Dividend = VLC_Dividend + substr(PLC_Operand, VLN_c1, 1)
				if VLC_Dividend = "0"
					VLC_Dividend = substr(VLC_Dividend, 2)
				endif	
			endfor	
		else
			VLC_Ret = "0"
		endif
*---	Ajusta o ponto decimal
		VLC_Ret = padl("", PLN_Decimals + 5, "0") + VLC_Ret
		VLN_i = len(VLC_Ret) - (PLN_Decimals + 4)
		VLC_Ret = substr(VLC_Ret, 1, VLN_i) + "." + substr(VLC_Ret, VLN_i + 1)

		if (VLN_Sign1 * VLN_Sign2) < 0
			VLC_Ret = "-"+VLC_Ret
		endif	
				
*--	Sum
	case PLC_What = "+"
		if VLN_Sign1 = 0
			VLC_Ret = iif(VLN_Sign2 < 0, "-" + PLC_Operator, PLC_Operator)
		else	
			if VLN_Sign2 = 0
				VLC_Ret = iif(VLN_Sign1 < 0, "-" + PLC_Operand, PLC_Operand)
			else
				if VLN_Sign1 = -1 
					if VLN_Sign2 = 1
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operator, PLC_Operand, "-")
					else	
						VLC_Ret = "-" + this.FOC_SumSubNumber(PLC_Operand, PLC_Operator, "+")
					endif
				else		
					if VLN_Sign2 = 1
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operand, PLC_Operator, "+")
					else	
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operator, PLC_Operand, "-")
					endif
				endif	
			endif
		endif		

*---	Ajusta o ponto decimal
		VLC_Ret = padl("", VLN_Dec1 + 1, "0") + VLC_Ret
		VLN_i = len(VLC_Ret)
		if VLN_Dec1 > 0	
			VLN_i = VLN_i - VLN_Dec1
			VLC_Ret = substr(VLC_Ret, 1, VLN_i) + "." + substr(VLC_Ret, VLN_i + 1)
		endif	
		
*--	Subtraction
	case PLC_What = "-"
		if VLN_Sign1 = 0
			VLC_Ret = iif(VLN_Sign2 < 0, "-" + PLC_Operator, PLC_Operator)
		else	
			if VLN_Sign2 = 0
				VLC_Ret = iif(VLN_Sign1 < 0, "-" + PLC_Operand, PLC_Operand)
			else
				if VLN_Sign1 = -1 
					if VLN_Sign2 = 1
						VLC_Ret = "-" + this.FOC_SumSubNumber(PLC_Operand, PLC_Operator, "+")
					else	
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operator, PLC_Operand, "-")
					endif
				else		
					if VLN_Sign2 = 1
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operand, PLC_Operator, "-")
					else	
						VLC_Ret = this.FOC_SumSubNumber(PLC_Operator, PLC_Operand, "+")
					endif
				endif	
			endif	
		endif		
*---	Ajusta o ponto decimal
		VLC_Ret = padl("", VLN_Dec1+1, "0") + VLC_Ret
		VLN_i = len(VLC_Ret)
		if VLN_Dec1 > 0	
			VLN_i = VLN_i - VLN_Dec1
			VLC_Ret = substr(VLC_Ret, 1, VLN_i) + "." + substr(VLC_Ret, VLN_i + 1)
		endif	
endcase

VLC_Ret = this.FOC_RetSignif(VLC_Ret, .T., .T.)

*- adjust the Decimals when required
VLL_Negative = (VLC_Ret = "-")
if VLL_Negative
	VLC_Ret = substr(VLC_Ret, 2)
endif

VLN_i = at(".", VLC_Ret)
if VLN_i > 0
	VLN_Dec1 = len(substr(VLC_Ret, VLN_i + 1))
	if VLN_Dec1 <= PLN_Decimals
		VLC_Ret = VLC_Ret + padl("", PLN_Decimals-VLN_Dec1, "0")
	else
		VLC_Ret = substr(VLC_Ret, 1, len(VLC_Ret) - (VLN_Dec1 - PLN_Decimals) + 1)
		if right(VLC_Ret, 1) <= "5"
			VLC_Ret = substr(VLC_Ret, 1, len(VLC_Ret) - 1)
		else
			VLC_Ret = substr(VLC_Ret, 1, len(VLC_Ret) - 1)
			VLC_Ret = this.FOC_RetSignif(strtran(VLC_Ret, ".", ""), .T., .F.)
			VLC_Ret = padl("", PLN_Decimals + 1, "0") + this.FOC_SumSubNumber(VLC_Ret, "1", "+")
			VLN_i = len(VLC_Ret)
			if PLN_Decimals > 0	
				VLN_i = VLN_i - PLN_Decimals
				VLC_Ret = substr(VLC_Ret, 1, VLN_i) + "." + substr(VLC_Ret, VLN_i + 1)
			endif	
			VLC_Ret = this.FOC_RetSignif(VLC_Ret, .T., .F.)
		endif
	endif	
else
	VLC_Ret = VLC_Ret + "." + padl("", PLN_Decimals, "0")
endif

if VLC_Ret = "."
	VLC_Ret = "0" + VLC_Ret
endif
if VLL_Negative
	VLC_Ret = "-" + VLC_Ret
endif
if PLN_Decimals = 0 and "." $ VLC_Ret
	VLC_Ret = strtran(VLC_Ret, ".", "")
endif
return VLC_Ret
endproc



*/----------------------------------------------------------------------------------------/*
*/ Função Objeto    : FOC_RetSignif                                                       /*
*/ Parâmetros  		: PLC_Number : Número a ser operado                                   /*
*/               	  PLL_Left   : Se retira não significativos do começo                 /*
*/               	  PLL_Right  : Se retira não significativos do final                  /*
*/ Descrição   		: Retira os dígitos não significativos do número gigante              /*
*/ Retorno 			: String contendo o resultado                                         /*
*/ Última alteração : 26/08/03                                                            /*
*/ Alterado por		: Denis Carrizo													      /*
*/----------------------------------------------------------------------------------------/*
procedure FOC_RetSignif
LParameters PLC_Number, PLL_Left, PLL_Right

LOCAL VLL_Negative, VLC_Aux

*- Retira os zeros não significativos à esquerda
if PLL_Left
	VLL_Negative = (left(PLC_Number, 1) == "-")
	if VLL_Negative
		PLC_Number = substr(PLC_Number, 2)
	endif
	do while .T.
		VLC_Aux = left(PLC_Number, 1)
		if VLC_Aux = "0"
			PLC_Number = substr(PLC_Number, 2)
		else
			exit
		endif
	enddo
	if VLL_Negative
		PLC_Number = "-"+PLC_Number
	endif
endif

*- Retira os zeros não significativos à direita
if PLL_Right and at(".", PLC_Number) > 0
	do while .T.
		VLC_Aux = right(PLC_Number, 1)
		if VLC_Aux = "0"
			PLC_Number = left(PLC_Number, len(PLC_Number) - 1)
		else
			if VLC_Aux = "."
				PLC_Number = left(PLC_Number, len(PLC_Number) - 1)
			endif	
			exit
		endif
	enddo
endif

if empty(PLC_Number)
	PLC_Number = "0"
endif	
return PLC_Number

endproc



*/----------------------------------------------------------------------------------------/*
*/ Função Objeto    : FOC_SumSubNumber                                                    /*
*/ Parâmetros  		: PLC_Operand   : Operando somente com significativos sem decimais    /*
*/               	  PLC_Operator  : Operador somente com significativos sem decimais    /*
*/               	  PLC_What   	: Operação a ser executada (+,-)                      /*
*/ Descrição   		: Executa soma ou subtração de inteiros gigantes.                     /*
*/ Retorno 			: string contendo o resultado                                         /*
*/ Última alteração : 26/08/03                                                            /*
*/ Alterado por		: Denis Carrizo													      /*
*/----------------------------------------------------------------------------------------/*
procedure FOC_SumSubNumber
LParameters PLC_Operand, PLC_Operator, PLC_What

LOCAL VLN_i, VLN_Len1, VLN_Len2, VLC_Ret, VLL_Negative, VLN_Rest, VLN_n1, VLN_n2, VLC_Aux

VLN_Len1 = len(PLC_Operand)
VLN_Len2 = len(PLC_Operator)

VLL_Negative = .F.
VLC_Aux = this.FOC_CompareNumber(PLC_Operand, PLC_Operator) 


if VLC_Aux = "<"
	VLC_Aux = PLC_Operand
	PLC_Operand = PLC_Operator
	PLC_Operator = VLC_Aux
	VLN_len1 = len(PLC_Operand)
	VLN_len2 = len(PLC_Operator)
	VLL_Negative = .T.
endif

VLC_Ret = ""
VLN_Rest = 0
do case
*--	Soma
	case PLC_What = "+"
		for VLN_i = 1 to VLN_Len1
			VLN_n1 = iif(VLN_i > VLN_Len1, 0, val(substr(PLC_Operand, VLN_Len1 - VLN_i + 1, 1)))
			VLN_n2 = iif(VLN_i > VLN_Len2, 0, val(substr(PLC_Operator, VLN_Len2 - VLN_i + 1, 1)))
			VLN_n1 = VLN_n1 + VLN_n2 + VLN_Rest
			if VLN_n1 > 9
				VLN_Rest = int(VLN_n1 / 10)
				VLN_n1 = VLN_n1 - (VLN_Rest * 10)
			else
				VLN_Rest = 0	
			endif	
			VLC_Ret = str(VLN_n1, 1) + VLC_Ret
		endfor
		if VLN_Rest > 0
			VLC_Ret = str(VLN_Rest, 1) + VLC_Ret
		endif
		
*--	Subtração
	case PLC_What = "-"
		for VLN_i = 1 to VLN_Len1
			VLN_n1 = iif(VLN_i > VLN_Len1, 0, val(substr(PLC_Operand, VLN_Len1 - VLN_i + 1, 1))) - ;
						VLN_Rest
			VLN_n2 = iif(VLN_i > VLN_Len2, 0, val(substr(PLC_Operator, VLN_Len2 - VLN_i + 1, 1)))
			if VLN_n1 < VLN_n2
				if VLN_i = VLN_Len1
					VLC_Ret = "-"+substr(PLC_Operator, 1, VLN_len2 - VLN_i) + ;
								str(VLN_n2 - VLN_n1) + VLC_Ret
					exit			
				else
					VLN_n1 = 10 + VLN_n1 - VLN_n2
					VLN_Rest = 1
				endif
			else
				VLN_n1 = VLN_n1 - VLN_n2
				VLN_Rest = 0
			endif
			VLC_Ret = str(VLN_n1,1) + VLC_Ret
		endfor
		if VLL_Negative
			if left(VLC_Ret,1) = "-"
				VLC_Ret = substr(VLC_Ret, 2)
			else
				VLC_Ret = "-" + VLC_Ret
			endif
		endif
endcase

return this.FOC_RetSignif(VLC_Ret, .T., .F.)

endproc

*/----------------------------------------------------------------------------------------/*
*/ Função Objeto    : FOC_CompareNumber                                                   /*
*/ Parâmetros  		: PLC_Operand   : Operando somente com significativos sem decimais    /*
*/               	  PLC_Operator  : Operador somente com significativos sem decimais    /*
*/ Descrição   		: Compara dois números gigantes                                       /*
*/ Retorno 			: ">" : Se operando > operador, "<" : Se operando < operador, "=" c.c./*
*/ Última alteração : 26/08/03                                                            /*
*/ Alterado por		: Denis Carrizo													      /*
*/----------------------------------------------------------------------------------------/*
procedure FOC_CompareNumber
LParameters PLC_Operand, PLC_Operator

LOCAL VLN_i, VLN_Len1, VLN_Len2, VLN_n1, VLN_n2

VLN_Len1 = len(PLC_Operand)
VLN_Len2 = len(PLC_Operator)

if VLN_Len1 < VLN_Len2
	return "<"
endif	
if VLN_Len1 > VLN_Len2
	return ">"
endif	

for VLN_i = 1 to VLN_Len1
	VLN_n1 = substr(PLC_Operand, VLN_i, 1)
	VLN_n2 = substr(PLC_Operator, VLN_i, 1)
	
	if VLN_n1 > VLN_n2
		return ">"
	endif	
	if VLN_n1 < VLN_n2
		return "<"
	endif	
endfor

return "="

endproc

*/-----------------------------------------------------------------------------------------/*
*/ Função Objeto    : FON_RoundingMask                                                     /*
*/ Parâmetros  		: PLN_PictNumber : Número da picture                                   /*
*/ Descrição   		: Verifica o arredondamento da picture                                 /*
*/ Retorno 			: Quantidade de casas decimais da picture                              /*
*/ Última alteração : 04/12/2003                                                           /*
*/ Alterado por		: Rodrigo													           /*
*/-----------------------------------------------------------------------------------------/*
procedure FON_RoundingMask

LParameters PLN_PictNumber

LOCAL VLN_Len, VLN_Counter, VLN_Return

VLN_Return = 0

if !empty(PLN_PictNumber) and between(PLN_PictNumber, 1, alen(this.VOA_Picts, 1))
	VLN_Len = len(this.VOA_Picts[PLN_PictNumber,2])

	for VLN_Counter = VLN_Len to 1 step -1
		if !isdigit(substr(this.VOA_Picts[PLN_PictNumber,2], VLN_Counter, 1))
			VLN_Return = VLN_Len - VLN_Counter
			exit
		endif
	endfor
endif

return VLN_Return

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Inicializa um semaforo que indica que existe uma instância do sistema rodando
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_InitSemaphore

	DECLARE INTEGER CreateSemaphore IN kernel32; 
    	    INTEGER lcSmAttr, INTEGER lInitialCount,; 
        	INTEGER lMaximumCount, STRING lpName 

	this.VON_SemaphoreHandle = CreateSemaphore(0, 1, 1, SS_SEMAPHORE_NAME)

	clear dlls CreateSemaphore

    return this.VON_SemaphoreHandle > 0

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Verifica se existe uma instância do sistema rodando (Não deve ser chamada quando o sistema estiver instanciado)
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_VerifyInstance
	lparameters PLC_SemaphoreName

	DECLARE INTEGER CloseHandle IN kernel32 INTEGER hObject 
	DECLARE INTEGER OpenSemaphore IN kernel32; 
	        INTEGER dwDesiredAccess,; 
	        INTEGER bInheritHandle, STRING lpName 

	local VLN_Handle

	VLN_Handle = OpenSemaphore(STANDARD_RIGHTS_REQUIRED, 0, PLC_SemaphoreName)
	CloseHandle(VLN_Handle)

	clear dlls CloseHandle, OpenSemaphore
	
	return VLN_Handle > 0

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Fecha a referência à instância aberta
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_CloseSemaphoreHandle

	DECLARE INTEGER CloseHandle IN kernel32 INTEGER hObject 

	if !empty(this.VON_SemaphoreHandle)
		CloseHandle(this.VON_SemaphoreHandle)
		this.VON_SemaphoreHandle = 0
	endif

	clear dlls CloseHandle

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Inicia um processo, retornado seu handle
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_CreateProcess
	lparameters PLC_App, PLC_CmdLine, PLC_CurDir

	declare INTEGER GetLastError in kernel32
    declare INTEGER CreateProcess in kernel32;
        STRING lpApplicationName, STRING lpCommandLine, INTEGER lpProcessAttr, INTEGER lpThreadAttr,; 
        INTEGER bInheritHandles, INTEGER dwCreationFlags, INTEGER lpEnvironment, STRING lpCurrentDirectory,; 
        STRING @lpStartupInfo, STRING @lpProcessInformation 


	local VLC_ProcInfo, VLC_StartInfo, VLN_Flags, VLN_Result, VLN_Process, VLN_Thread 

	VLN_Process = 0
	VLC_ProcInfo = replicate(chr(0), 16)
    VLC_StartInfo = padr(chr(START_INFO_SIZE), START_INFO_SIZE, chr(0))
    VLN_Flags = 0 
    VLC_App = alltrim(PLC_App)
    VLC_CmdLine = " " + alltrim(PLC_CmdLine)

    VLN_Result = CreateProcess(VLC_App, VLC_CmdLine, 0, 0, 0, VLN_Flags, 0, PLC_CurDir, @VLC_StartInfo, @VLC_ProcInfo)

    if VLN_Result = 0 
	    *-  2 = ERROR_FILE_NOT_FOUND 
	    *-  3 = ERROR_PATH_NOT_FOUND 
	    *-  5 = ERROR_ACCESS_DENIED 
	    *- 87 = ERROR_INVALID_PARAMETER 
    else
	    *- process and thread handles returned in ProcInfo structure 
	    VLN_Process = this.FON_Buf2DWord(substr(VLC_ProcInfo, 1,4))
	    VLN_Thread = this.FON_Buf2DWord(substr(VLC_ProcInfo, 5,4))
		VLN_Result = VLN_Process
    endif

	clear dlls GetLastError, CreateProcess

	return VLN_Process

endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Encerra um processo
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_TerminateProcess
	lparameters PLN_ProcNo

	declare INTEGER TerminateProcess IN kernel32 INTEGER hProcess, INTEGER uExitCode
	=TerminateProcess(PLN_ProcNo, 0)
	clear dlls TerminateProcess
	
	return
endproc


*/------------------------------------------------------------------------------------------------------------------------
*/ Inicia o Star Soft Messenger
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_StartMessenger
	lparameters PLL_Minimized

	local VLC_InitSettings, VLO_ClipBoard, VLL_Return, VLC_FullPath, VLC_UserName, VLC_LanStr, VLC_MachineName
	
	VLC_FullPath = fullpath("utilities\messenger\messenger.exe")
	VLL_Return = !this.FOL_VerifyInstance(MESSENGER_SEMAPHORE_NAME) and file(VLC_FullPath, 1)

	if VLL_Return
		VLC_LanStr = "<machine>" + strtran(lower(sys(0)), "#", "</machine><user>") + "</user>"
		VLC_LanUserName = alltrim(this.FOC_TagContent(VLC_LanStr, "user", 1))
		VLC_MachineName = alltrim(this.FOC_TagContent(VLC_LanStr, "machine", 1))
		*- Armazena as informações para a intância do executável

		VLC_InitSettings = "<user.language>" + this.VOC_User_Language + "</user.language>" +;
							"<user.name>" + VLC_LanUserName + "</user.name>" +;
							"<user.dirsettings>" + this.VOC_DirSettings + "</user.dirsettings>" +;
							"<user.filesettings>" + this.VOC_FileSettings + "</user.filesettings>" +;
							"<user.tmp>" + this.VOC_UserTmp + "</user.tmp>" +;
							"<isclosable>" + transform(this.VON_MessengerMode, "@L 9") + "</isclosable>" +;
							"<start.minimized>" + iif(empty(PLL_Minimized), "0", "1") + "</start.minimized>" +;
							"<dirhome>" + this.VOC_DirHome + "</dirhome>"

		VLO_ClipBoard = createobject("cgo_libmain12")
		VLO_ClipBoard.FOL_CopyToClip(VLC_InitSettings)

		this.VON_MessengerProcessHnd = this.FON_CreateProcess(VLC_FullPath, "", SYS(5)+SYS(2003))
		VLL_Return = this.VON_MessengerProcessHnd > 0
	endif

	return VLL_Return
endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Inicia o Star Soft Messenger
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_TerminateMessenger
	if this.FOL_VerifyInstance(MESSENGER_SEMAPHORE_NAME) and !empty(this.VON_MessengerProcessHnd)
		this.FOL_TerminateProcess(this.VON_MessengerProcessHnd)
		this.VON_MessengerProcessHnd = 0
	endif
endproc



*/------------------------------------------------------------------------------------------------------------------------
*/ Retorna o número de relatórios de acordo com o filtro, preenchendo uma array (referência) com suas configurações
*/------------------------------------------------------------------------------------------------------------------------
procedure FON_GetReports
lparameters PLN_FilterType, PLC_Filter, PRA_RetReportArray, PLN_EspStd

	local VLN_RetReportCount, VLC_RepCfg, VLC_ReportLevel, VLN_FCount, VLN_ClientRepoCount, VLL_AddReport, VLN_Counter
	local VLC_FileName

	if empty(PLN_EspStd)
		PLN_EspStd = 0
	endif

	this.FOL_UseInPackage("y06", "y06")
	this.FOL_UseInPackage("y33", "y33")

	VLN_FCount = afields(VLA_Y06Struct, "Y06")
	*- Adiciona um campo para identificar relatórios específicos
	VLN_FCount = VLN_FCount + 1
	dimension VLA_Y06Struct[VLN_FCount, AFIELDS2NDCOL]
	VLA_Y06Struct(VLN_FCount, 1) = "esp"
	VLA_Y06Struct(VLN_FCount, 2) = "l"
	VLA_Y06Struct(VLN_FCount, 3) = 1
	VLA_Y06Struct(VLN_FCount, 4) = 0
	store "" to VLA_Y06Struct(VLN_FCount, 7), VLA_Y06Struct(VLN_FCount, 8),  VLA_Y06Struct(VLN_FCount, 9), ;
				VLA_Y06Struct(VLN_FCount, 10), VLA_Y06Struct(VLN_FCount, 11), VLA_Y06Struct(VLN_FCount, 12), ;
				VLA_Y06Struct(VLN_FCount, 13), VLA_Y06Struct(VLN_FCount, 14), VLA_Y06Struct(VLN_FCount, 15), ;
				VLA_Y06Struct(VLN_FCount, 16)

	*- Adiciona um campo para ordenar os relatórios
	VLN_FCount = VLN_FCount + 1
	dimension VLA_Y06Struct[VLN_FCount, AFIELDS2NDCOL]
	VLA_Y06Struct(VLN_FCount, 1) = "nodeord"
	VLA_Y06Struct(VLN_FCount, 2) = "c"
	VLA_Y06Struct(VLN_FCount, 3) = 200
	VLA_Y06Struct(VLN_FCount, 4) = 0
	store "" to VLA_Y06Struct(VLN_FCount, 7), VLA_Y06Struct(VLN_FCount, 8),  VLA_Y06Struct(VLN_FCount, 9), ;
				VLA_Y06Struct(VLN_FCount, 10), VLA_Y06Struct(VLN_FCount, 11), VLA_Y06Struct(VLN_FCount, 12), ;
				VLA_Y06Struct(VLN_FCount, 13), VLA_Y06Struct(VLN_FCount, 14), VLA_Y06Struct(VLN_FCount, 15), ;
				VLA_Y06Struct(VLN_FCount, 16)

	*- Adiciona um campo para armazenar o título completo dos relatórios
	VLN_FCount = VLN_FCount + 1
	dimension VLA_Y06Struct[VLN_FCount, AFIELDS2NDCOL]
	VLA_Y06Struct(VLN_FCount, 1) = "title"
	VLA_Y06Struct(VLN_FCount, 2) = "c"
	VLA_Y06Struct(VLN_FCount, 3) = 240
	VLA_Y06Struct(VLN_FCount, 4) = 0
	store "" to VLA_Y06Struct(VLN_FCount, 7), VLA_Y06Struct(VLN_FCount, 8),  VLA_Y06Struct(VLN_FCount, 9), ;
				VLA_Y06Struct(VLN_FCount, 10), VLA_Y06Struct(VLN_FCount, 11), VLA_Y06Struct(VLN_FCount, 12), ;
				VLA_Y06Struct(VLN_FCount, 13), VLA_Y06Struct(VLN_FCount, 14), VLA_Y06Struct(VLN_FCount, 15), ;
				VLA_Y06Struct(VLN_FCount, 16)


	*- Adiciona um campo para armazenar o form de filtro
	VLN_FCount = VLN_FCount + 1
	dimension VLA_Y06Struct[VLN_FCount, AFIELDS2NDCOL]
	VLA_Y06Struct(VLN_FCount, 1) = "form"
	VLA_Y06Struct(VLN_FCount, 2) = "c"
	VLA_Y06Struct(VLN_FCount, 3) = 30
	VLA_Y06Struct(VLN_FCount, 4) = 0
	store "" to VLA_Y06Struct(VLN_FCount, 7), VLA_Y06Struct(VLN_FCount, 8),  VLA_Y06Struct(VLN_FCount, 9), ;
				VLA_Y06Struct(VLN_FCount, 10), VLA_Y06Struct(VLN_FCount, 11), VLA_Y06Struct(VLN_FCount, 12), ;
				VLA_Y06Struct(VLN_FCount, 13), VLA_Y06Struct(VLN_FCount, 14), VLA_Y06Struct(VLN_FCount, 15), ;
				VLA_Y06Struct(VLN_FCount, 16)

	*- Adiciona um campo para armazenar a data de validade de um relatório específico
	VLN_FCount = VLN_FCount + 1
	dimension VLA_Y06Struct[VLN_FCount, AFIELDS2NDCOL]
	VLA_Y06Struct(VLN_FCount, 1) = "valid"
	VLA_Y06Struct(VLN_FCount, 2) = "d"
	VLA_Y06Struct(VLN_FCount, 3) = 8
	VLA_Y06Struct(VLN_FCount, 4) = 0
	store "" to VLA_Y06Struct(VLN_FCount, 7), VLA_Y06Struct(VLN_FCount, 8),  VLA_Y06Struct(VLN_FCount, 9), ;
				VLA_Y06Struct(VLN_FCount, 10), VLA_Y06Struct(VLN_FCount, 11), VLA_Y06Struct(VLN_FCount, 12), ;
				VLA_Y06Struct(VLN_FCount, 13), VLA_Y06Struct(VLN_FCount, 14), VLA_Y06Struct(VLN_FCount, 15), ;
				VLA_Y06Struct(VLN_FCount, 16)
	
	create cursor tmp_y06 from array VLA_Y06Struct
	select tmp_y06
	index on ukey tag ukey
	index on y06_ukeyp tag y06_ukeyp
	index on nodeord tag nodeord
	index on form tag form
	index on form + ukey tag form_ukey
	index on form + nodeord tag form_ord

	*- Relatórios padrões ou todos
	if inlist(PLN_EspStd, 0, 1)
		do case 
			case empty(PLN_FilterType)
				*- Filtro pelo código do form
				select y33
				set order to y33_001_c in y33
				=seek(PLC_Filter, "y33", "y33_001_c")
				scan while alltrim(y33.y33_001_c) == alltrim(PLC_Filter)
					if !empty(nvl(y33.y06_ukey, "")) and seek(y33.y06_ukey, "y06", "ukey")
						select y06
						*- Armazena os campos do registro no array
						if !empty(y06.y06_005_n) and !this.FOL_AcessMenu(y06.ukey)
							VLC_RepCfg = this.FOC_RelatField()
							if empty(VLC_RepCfg)
								loop
							else
								select y06
								scatter memvar memo
								m.y06_006_m = VLC_RepCfg
								append blank in tmp_y06
								select tmp_y06
								gather memvar memo
								replace tmp_y06.y06_002_c with this.FOC_Caption(alltrim(tmp_y06.y06_002_c)) in tmp_y06
							endif

						endif
					endif
					select y33
				endscan

			case PLN_FilterType = 1
				*- Filtro pela ukey do relatório
				select y06
				set order to ukey in y06
				if seek(PLC_Filter, "y06", "ukey")
					select y06
					*- Armazena os campos do registro no array
					if !empty(y06.y06_005_n) and !this.FOL_AcessMenu(y06.ukey)
						VLC_RepCfg = this.FOC_RelatField()
						if !empty(VLC_RepCfg)
							select y06
							scatter memvar memo
							m.y06_006_m = VLC_RepCfg
							append blank in tmp_y06
							select tmp_y06
							gather memvar memo
							replace tmp_y06.y06_002_c with this.FOC_Caption(alltrim(tmp_y06.y06_002_c)) in tmp_y06
						endif
					endif
				endif

			case PLN_FilterType = 2
				*- Sem filtro, todos os relatórios para carregamento da tela de acesso
				select y33
				set order to y33_001_c in y33
				scan for !empty(nvl(y33.y06_ukey, "")) and seek(y33.y06_ukey, "y06", "ukey")
					select y06
					scatter memvar memo
					append blank in tmp_y06
					select tmp_y06
					gather memvar memo
					replace tmp_y06.y06_002_c 	with this.FOC_Caption(alltrim(tmp_y06.y06_002_c)),;
							tmp_y06.form 		with y33.y33_001_c;
					in tmp_y06

					select y33
				endscan

		endcase
	endif

	*- Relatórios específicos
	if inlist(PLN_EspStd, 0, 2)
		local array VLA_ReportFile[1]
		VLN_ClientRepoCount = adir(VLA_ReportFile, this.VOC_ClientRepoDir + "*.ssr")
		for VLN_Counter = 1 to VLN_ClientRepoCount
			VLL_AddReport = .f.
			VLC_ReportCfg = ""
			VLC_FileName = this.VOC_ClientRepoDir + VLA_ReportFile[VLN_Counter, 1]
			VLC_FileContent = filetostr(VLC_FileName)
			VLC_HeaderContent = strextract(VLC_FileContent, "\\BEGIN REPORTHEADER", "\\END REPORTHEADER", 1, 1)
			if !empty(VLC_HeaderContent)
				VLN_HeaderLines = alines(VLA_Header, VLC_HeaderContent, .t.)
				VLC_FormList = ""
				VLC_ReportId = ""
				VLC_ValidKey = ""

				for VLN_HCounter = 1 to VLN_HeaderLines
					VLC_Line = VLA_Header[VLN_HCounter]

					if "#VOC_FORMFILTER" $ VLC_Line
						VLC_FormList = this.FOU_PropertyValue(VLC_Line, "#VOC_FORMFILTER")
					endif

					if "#VOC_REPORTID" $ VLC_Line
						VLC_ReportId = this.FOU_PropertyValue(VLC_Line, "#VOC_REPORTID")
					endif

					if "#VOC_VALIDATIONKEY" $ VLC_Line
						VLC_ValidKey = this.FOU_PropertyValue(VLC_Line, "#VOC_VALIDATIONKEY")
					endif

				endfor

				if !empty(VLC_FormList) and !empty(VLC_ReportId)
					if PLN_FilterType = 2
						VLL_AddReport = .t.
					else
						VLL_AddReport = (this.FOL_SearchIn(upper(alltrim(PLC_Filter)), upper(VLC_FormList)) and;
								 empty(PLN_FilterType)) or (!empty(PLN_FilterType) and PLC_Filter == VLC_ReportId)

						if VLL_AddReport
							VLL_AddReport = !this.FOL_AcessMenu(VLC_ReportId)
						endif
					endif
				endif


				if VLL_AddReport
					VLC_ReportCfg = strextract(VLC_FileContent, "\\BEGIN REPORTCFG", "\\END REPORTCFG", 1, 1)
					VLC_QuestionCfg = strextract(VLC_FileContent, "\\BEGIN QUESTION", "\\END QUESTION", 1, 1)
			
					if !empty(VLC_HeaderContent) and !empty(VLC_ReportCfg) and !empty(VLC_QuestionCfg)
						VLC_Modules = this.FOU_PropertyValue(VLC_HeaderContent, "#VOC_MODULES")
						VLC_Description = this.FOU_PropertyValue(VLC_HeaderContent, "#VOC_DESCRIPTION")
						VLC_ShowOrder = trans(val(this.FOU_PropertyValue(VLC_HeaderContent, "#VON_SHOWORDER")), "@L 9999")
						VLC_ParNodeUkey = this.FOU_PropertyValue(VLC_HeaderContent, "#VOC_PARNODEUKEY")
						
						VLC_RepCfg = "\\BEGIN REPORTCFG" + VLC_ReportCfg + CRLF + "\\END REPORTCFG" + CRLF +;
									"\\BEGIN QUESTION" + CRLF + VLC_QuestionCfg + "\\END QUESTION"

						VLL_ValidKey = .t.
						*- Verifico se a chave é válida
						if .f. && !empty(VLC_ValidKey) and PLN_FilterType <> 2
							VLC_Info = ""
							VLN_Action = this.VOO_Security.FON_ValidateKey(2, VLC_ValidKey, VLC_RepCfg, @VLC_Info)
							do case
								case VLN_Action = 1 && Chave provisória válida
									VLL_ValidKey = .t.
									VLC_Expires = this.FOC_TagContent(VLC_Info, "EXP_DATE", 1)

								case VLN_Action = 2 and !empty(VLC_Info) && Chave provisória precisa ser armazenada
									VLC_FContent = filetostr(VLC_FileName)
									VLC_KeyStr = "#VOC_VALIDATIONKEY"
									*- Início da substituição da chave existente
									VLN_StartCount = at(VLC_KeyStr, VLC_FContent, 1) + len(VLC_KeyStr)
									VLN_FileLen = len(VLC_FContent)
									VLN_CharsReplaced = 0
									VLL_CharFound = .f.
									*- Procuro o final da cadeia para substituir
									for VLN_CharCounter = VLN_StartCount to VLN_FileLen
										VLC_Char = substr(VLC_FContent, VLN_CharCounter, 1)
										if inlist(VLC_Char, chr(13), chr(10))
											VLL_CharFound = .t.
											exit
										else
											VLN_CharsReplaced = VLN_CharsReplaced + 1
										endif
									endfor
									
									if VLL_CharFound
										VLC_TempKey = " = " + this.FOC_TagContent(VLC_Info, "KEY_STR", 1)
										VLC_FContent = stuff(VLC_FContent, VLN_StartCount, VLN_CharsReplaced, VLC_TempKey)
										if strtofile(VLC_FContent, VLC_FileName, 0) > 0
											VLL_ValidKey = .t.
											VLC_Expires = this.FOC_TagContent(VLC_Info, "EXP_DATE", 1)
										endif
									endif
							endcase
						endif
						
						if PLN_FilterType <> 2
							*- Arquivo válido temporariamente
							if VLL_ValidKey
								VLD_ValidDate = {} && this.FOD_StoD(VLC_Expires)
							else
								VLD_ValidDate = {}
							endif
						
						
							*- Estará bloqueado a pártir da próxima versão (4.02)
							VLL_ValidKey = .t.
							*---------------------------------------------
							append blank in tmp_y06
							replace tmp_y06.ukey 		with VLC_ReportId,;
									tmp_y06.y06_ukeyp	with VLC_ParNodeUkey,;
									tmp_y06.y06_005_n	with 1,;
									tmp_y06.y06_002_c	with VLC_Description,;
									tmp_y06.y06_004_c	with VLC_ShowOrder,;
									tmp_y06.y06_006_m	with iif(VLL_ValidKey, VLC_RepCfg, ""),;
									tmp_y06.valid		with VLD_ValidDate,;
									tmp_y06.esp			with .t.;
							in tmp_y06
						else
							VLC_CurrentForm = ""
							VLC_FormList = VLC_FormList + "//"
							VLN_CharCount = len(VLC_FormList)
							for VLN_Cont = 1 to VLN_CharCount
								VLC_Char = substr(VLC_FormList, VLN_Cont, 1)
								if isdigit(VLC_Char) or isalpha(VLC_Char) or asc(VLC_Char) = 95
									VLC_CurrentForm = VLC_CurrentForm + VLC_Char
								else
									VLC_CurrentForm = alltrim(VLC_CurrentForm)
									if !empty(VLC_CurrentForm)
										append blank in tmp_y06
										replace tmp_y06.ukey 		with VLC_ReportId,;
												tmp_y06.y06_ukeyp	with VLC_ParNodeUkey,;
												tmp_y06.y06_005_n	with 1,;
												tmp_y06.y06_002_c	with VLC_Description,;
												tmp_y06.y06_004_c	with VLC_ShowOrder,;
												tmp_y06.y06_006_m	with "",;
												tmp_y06.form		with VLC_CurrentForm,;
												tmp_y06.valid		with {},;
												tmp_y06.esp			with .t.;
										in tmp_y06
									endif
									VLC_CurrentForm = ""
								endif
							endfor
						endif

					endif

				endif
			endif
		endfor
	endif

	*- Carrega os nós dos pais dos relatórios selecionados e refaz a ordem de carregamento da árvore
	select tmp_y06
	set order to 0
	scan
		VLC_ParNode = tmp_y06.y06_ukeyp
		VLN_Recno = recno("tmp_y06")
		VLC_CurOrder = alltrim(tmp_y06.y06_004_c)
		VLC_CurTitle = alltrim(tmp_y06.y06_002_c)
		VLC_Form = tmp_y06.form

		do while !empty(VLC_ParNode)
			if !seek(VLC_Form + VLC_ParNode, "tmp_y06", "form_ukey")
				if seek(VLC_ParNode, "y06", "ukey")
					select y06
					scatter memvar memo
					append blank in tmp_y06
					select tmp_y06
					m.form = VLC_Form
					m.y06_002_c = this.FOC_Caption(alltrim(m.y06_002_c))
					gather memvar memo
				endif
			endif

			if tmp_y06.ukey == VLC_ParNode
				VLC_CurOrder = alltrim(tmp_y06.y06_004_c) + "." + VLC_CurOrder
				VLC_CurTitle = alltrim(tmp_y06.y06_002_c) + " - " + VLC_CurTitle
				VLC_ParNode = tmp_y06.y06_ukeyp
			else
				VLC_ParNode = ""
			endif
		enddo
	
		select tmp_y06
		set order to 0
		go (VLN_Recno) in tmp_y06
		replace tmp_y06.nodeord with VLC_CurOrder,;
				tmp_y06.title 	with VLC_CurTitle;
		in tmp_y06
	endscan

	VLN_RetReportCount = 0


	if PLN_FilterType <> 2
		VLC_ValidUntilCaption = " (" + this.FOC_Caption("valido ate") + " "
		VLC_NotIdentifiedCaption = " (" + this.FOC_Caption("nao identificado") + ")"

		select tmp_y06
		set order to nodeord
		scan
			VLC_Description = alltrim(tmp_y06.y06_002_c)
			if tmp_y06.esp
				*- Verificação de específicos
				if !empty(tmp_y06.valid)
					VLC_Description = VLC_Description + VLC_ValidUntilCaption + dtoc(tmp_y06.valid) + ")"
				else
					&& VLC_Description = VLC_Description + VLC_NotIdentifiedCaption
				endif
			endif


			VLN_RetReportCount = VLN_RetReportCount + 1	
			dimension PRA_RetReportArray[VLN_RetReportCount, 12]
			PRA_RetReportArray[VLN_RetReportCount, 1] = tmp_y06.ukey
			PRA_RetReportArray[VLN_RetReportCount, 2] = VLC_Description
			PRA_RetReportArray[VLN_RetReportCount, 3] = nvl(tmp_y06.y06_006_m, "") && Sempre coloca a configuração correta nesse campo
			PRA_RetReportArray[VLN_RetReportCount, 4] = tmp_y06.y06_ukeyp
			PRA_RetReportArray[VLN_RetReportCount, 5] = tmp_y06.y06_005_n
			PRA_RetReportArray[VLN_RetReportCount, 6] = tmp_y06.y06_004_c
			PRA_RetReportArray[VLN_RetReportCount, 7] = tmp_y06.y06_001_c
			PRA_RetReportArray[VLN_RetReportCount, 8] = tmp_y06.esp
			PRA_RetReportArray[VLN_RetReportCount, 9] = tmp_y06.nodeord
			PRA_RetReportArray[VLN_RetReportCount, 10] = occurs(".", tmp_y06.nodeord) + 1
			PRA_RetReportArray[VLN_RetReportCount, 11] = alltrim(tmp_y06.title)
			PRA_RetReportArray[VLN_RetReportCount, 12] = alltrim(tmp_y06.y06_002_c)
		endscan

		if used("tmp_y06")
			use in tmp_y06
		endif
	else
		VLN_RetReportCount = reccount("tmp_y06")
	endif

	if used("Y06")
		use in y06
	endif

	if used("Y33")
		use in y33
	endif


	return VLN_RetReportCount

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Abre arquivos DBF locais, armazenados em packs
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_UseInPackage
lparameters PLC_FileName, PLC_Alias, PLC_Order, PLN_Area, PLL_Again

local VLL_Return

	if !(this.FOC_StartPar("ERPSYSTEM") == 'JSTART')  

		do components\controls\libraries\SearchEngine with .t., PLC_FileName, PLC_Alias, PLC_Order, PLN_Area, PLL_Again, VLL_Return

	else

		if used(PLC_Alias)
			use in (PLC_Alias)
		endif
		
		if empty(PLN_Area)
			select 0
		else
			select (PLN_Area)
		endif

		if !empty(PLL_Again)
			use (PLC_FileName) again noupdate alias (PLC_Alias)
		else
			use (PLC_FileName) noupdate alias (PLC_Alias)
		endif

		if !empty(PLC_Order)
			set order to (PLC_Order) in (PLC_Alias)
		endif

		VLL_Return = used(PLC_Alias)

	endif
	
	return VLL_Return

endproc

*/------------------------------------------------------------------------------------------------------------------------
*/ Abre arquivos DBF locais, armazenados em packs
*/------------------------------------------------------------------------------------------------------------------------
procedure FOL_UseInCNABPackage
lparameters PLC_FileName, PLC_Alias, PLC_Order, PLN_Area, PLL_Again

local VLL_Return
	if !(this.FOC_StartPar("ERPSYSTEM") == 'JSTART')  
		do components\controls\libraries\cnab with .t., PLC_FileName, PLC_Alias, PLC_Order, PLN_Area, PLL_Again, VLL_Return
	else
		if used(PLC_Alias)
			use in (PLC_Alias)
		endif
		if empty(PLN_Area)
			select 0
		else
			select (PLN_Area)
		endif
		if !empty(PLL_Again)
			use (PLC_FileName) again noupdate alias (PLC_Alias)
		else
			use (PLC_FileName) noupdate alias (PLC_Alias)
		endif
		if !empty(PLC_Order)
			set order to (PLC_Order) in (PLC_Alias)
		endif
		VLL_Return = used(PLC_Alias)
	endif
	
	return VLL_Return
endproc

*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_UpAndDownLoadResource
*/ Parametros       : PLN_Type  : 1 - UpLoad, 2 - DownLoad, 3 - Novo Resource, 
*/					  4 - Salva o Arquivo, 5 - Exclui o Arquivo.
*/					  PLC_Par	: Passar o Par de armazenamento do UXR
*/					  PLC_UKeyp	: Passar o UKeyp de armazenamento do UXR
*/ Descrição        : Trabalha com o Resource do usuário, para trabalhar com arquivos que
*/					  estar corrompidos.
*/ Retorno          : .T. ou .F.										  
*/----------------------------------------------------------------------------------------
procedure FOL_WorkResource
lparameters PLN_Type, PLC_Par, PLC_Ukeyp, PLC_File, PLL_Message

	if this.VOL_ServiceMode or empty(nvl(PLC_Ukeyp, "")) or (this.VON_DbType = 2 and !(this.FOC_StartPar("RESOURCE") == 'DBSERVER'))  
		return .T.
	endif

	local VLO_Zip, VLC_ZipFile, VLC_DirDownload, VLL_Continue, VLL_ContMessage, VLC_FileCutCopyFile
	VLO_Zip = null
	VLC_ZipFile = alltrim(this.VOC_DIRSETTINGS) + '\ResourceZip.Zip'
	VLC_DirTemp = alltrim(this.VOC_DIRSETTINGS) + '\TEMP'
	VLL_Continue = .T.
	VLC_FileResource = justfname(strtran(this.VOC_FileSettings, '.res'))
	VLC_FileCutCopyFile = justfname(strtran(this.VOC_CutCopyFile, '.res'))
	VLL_ContMessage = .T.

	VLO_Zip = createobject("CGO_Zip")

	do case
		case PLN_Type = 1 && UpLoad

			this.FOL_CloseFileInAllDataSession(VLC_FileResource)
			this.FOL_CloseFileInAllDataSession(VLC_FileCutCopyFile)
			
			VLL_Continue = this.FOL_usetable(this.VOC_FileSettings, "ActualRes") and !empty(tag(1, 'ActualRes'))
			
			if VLL_Continue
				VLL_Continue = this.FOL_usetable(this.VOC_CutCopyFile, "ActualCut") and !empty(tag(1, 'ActualCut')) and !empty(tag(2, 'ActualCut'))
			endif

			this.FOL_CloseTable("ActualRes")
			this.FOL_CloseTable("ActualCut")

			if !VLL_Continue
				VLL_ContMessage = .T.
				if PLL_Message
					VLL_ContMessage = messagebox("O Arquivo Settings atual esta danificado e não podera ser gravado, deseja recria-lo",4+16) = 6
				endif
				if VLL_ContMessage
					delete file (strtran(this.VOC_FileSettings, '.'+justext(this.VOC_FileSettings), '.*'))
					delete file (strtran(this.VOC_CutCopyFile, '.'+justext(this.VOC_CutCopyFile), '.*'))
					this.FOL_SetUserResource()
					VLL_Continue = .T.
				else
					VLL_Continue = .F.
				endif
			endif

			if VLL_Continue
				
				VLL_Continue = (VLO_Zip.FON_ZipFiles(VLC_ZipFile, strtran(this.VOC_FileSettings, '.'+justext(this.VOC_FileSettings), '.*'), .T.)=0)
				if VLL_Continue
					VLL_Continue = (VLO_Zip.FON_ZipFiles(VLC_ZipFile, strtran(this.VOC_CutCopyFile, '.'+justext(this.VOC_CutCopyFile), '.*'), .T.)=0)
				endif
				
				if VLL_Continue and file(VLC_ZipFile)
					* Excluindo o Arquivo
					this.FOL_SETPARAMETER(1, PLC_Par)
					this.FOL_SETPARAMETER(2, PLC_Ukeyp)
					this.FOL_EditCursor("UXR_PARUKEYP", "UXR_PARUKEYP", "UXRT", "Z", 0, 2)
					
					select UXRT
					replace all mycontrol	with '1' in UXRT
					delete all
					
					select UXRT
					this.FOL_NewRecordtoForm("UXRT")
					replace mycontrol	with '1',;
							uxr_001_m	with filetostr(VLC_ZipFile),;
							uxr_par		with PLC_Par,;
							uxr_ukeyp	with PLC_Ukeyp in UXRT
					
					this.FOL_CreateSqlString("UXRT", "UXR")
					
					VLL_Continue = this.FOL_BeginTRansaction("UXR")
					if VLL_Continue
						VLL_Continue = this.FOL_SaveCursorSql(null, "UXRT", "UXR", "UKEY")
					endif
					if !VLL_Continue
						this.FOL_RollBack("UXR")
					else
						this.FOL_CommitTrans("UXR")
					endif
					
				endif
			
				delete file (VLC_ZipFile)
			endif
		
		case PLN_Type = 2 && Download

			this.FOL_CloseFileInAllDataSession(VLC_FileResource)
			this.FOL_CloseFileInAllDataSession(VLC_FileCutCopyFile)

			this.FOL_SETPARAMETER(1, PLC_Par)
			this.FOL_SETPARAMETER(2, PLC_Ukeyp)
			this.FOL_EditCursor("UXR_PARUKEYP", "UXR_PARUKEYP", "UXRT", "Z", 0, 2)

			if reccount("UXRT") > 0
				VLC_StringToFile = 	UXRT.UXR_001_M
			else
				VLL_Continue = .F.
			endif

			if VLL_Continue
	
				if !directory(VLC_DirTemp)
					mkdir (VLC_DirTemp)
				endif
				
				this.FOL_SetUserResource()

				strtofile(VLC_StringToFile, VLC_DirTemp+'\DownZip.zip')
				VLL_Continue = file(VLC_DirTemp+'\DownZip.zip')

				if VLL_Continue
					VLL_Continue = (VLO_Zip.FON_UnZipFiles(VLC_DirTemp+'\DownZip.zip', "", VLC_DirTemp) = 0)
					if VLL_Continue
		
						if !this.FOL_usetable(this.VOC_FileSettings, "ActualRes", , , ,.T., .T.)
							delete file (strtran(this.VOC_FileSettings, '.'+justext(this.VOC_FileSettings), '.*'))
							this.FOL_CreateResourceFile(.t., .f.)
							this.FOL_usetable(this.VOC_FileSettings, "ActualRes", , , , , .T.)
						endif

						select ActualRes
						zap
						append from (VLC_DirTemp+"\resource.res")

						this.FOL_CloseTable("ActualRes")

						if !this.FOL_usetable(this.VOC_CutCopyFile, "ActualCut", , , , .T., .T.)
							delete file (strtran(this.VOC_CutCopyFile, '.'+justext(this.VOC_CutCopyFile), '.*'))
							this.FOL_CreateResourceFile(.f., .t.)
							this.FOL_usetable(this.VOC_CutCopyFile, "ActualCut", , , , , .T.)
						endif

						select ActualCut
						zap
						append from (VLC_DirTemp+"\resource1.res")

						this.FOL_CloseTable("ActualCut")

					endif

				endif

				delete file (VLC_DirTemp+'\DownZip.zip')
			endif

		case PLN_Type = 3 && Criar novo Resource
			
			this.FOL_CloseFileInAllDataSession(VLC_FileResource)
			this.FOL_CloseFileInAllDataSession(VLC_FileCutCopyFile)

			
			delete file (strtran(this.VOC_FileSettings, '.'+justext(this.VOC_FileSettings), '.*'))
			delete file (strtran(this.VOC_CutCopyFile, '.'+justext(this.VOC_CutCopyFile), '.*'))
			
			this.FOL_SetUserResource()
			VLL_Continue = .T.

		case PLN_Type = 4 && Salva o arquivo informado
			
			if !empty(PLC_File)
				if file(PLC_File)
					PLC_File = filetostr(PLC_File)
				endif
			endif
			
			this.FOL_SETPARAMETER(1, PLC_Par)
			this.FOL_SETPARAMETER(2, PLC_Ukeyp)
			this.FOL_EditCursor("UXR_PARUKEYP", "UXR_PARUKEYP", "UXRT", "Z", 0, 2)
			
			select UXRt
			Replace all mycontrol with '1' in UXRt
			delete all
		
			select UXRT
			this.FOL_NewRecordtoForm("UXRT")
			replace mycontrol	with '1',;
					uxr_001_m	with PLC_File,;
					uxr_par		with PLC_Par,;
					uxr_ukeyp	with PLC_Ukeyp in UXRT
			
			this.FOL_CreateSqlString("UXRT", "UXR")
			
			VLL_Continue = this.FOL_BeginTRansaction("UXR")
			if VLL_Continue
				VLL_Continue = this.FOL_SaveCursorSql(null, "UXRT", "UXR", "UKEY")
			endif
			if !VLL_Continue
				this.FOL_RollBack("UXR")
			else
				this.FOL_CommitTrans("UXR")
			endif

		case PLN_Type = 5 && Exclui o arquivo informado
			
			this.FOL_SETPARAMETER(1, PLC_Par)
			this.FOL_SETPARAMETER(2, PLC_Ukeyp)
			this.FOL_EditCursor("UXR_PARUKEYP", "UXR_PARUKEYP", "UXRT", "Z", 0, 2)
			
			select UXRt
			Replace all mycontrol with '1' in UXRt
			delete all
		
			this.FOL_CreateSqlString("UXRT", "UXR")
			
			VLL_Continue = this.FOL_BeginTRansaction("UXR")
			if VLL_Continue
				VLL_Continue = this.FOL_SaveCursorSql(null, "UXRT", "UXR", "UKEY")
			endif
			if !VLL_Continue
				this.FOL_RollBack("UXR")
			else
				this.FOL_CommitTrans("UXR")
			endif

		case PLN_Type = 6 && Caso não exista o resource gravado o sistema sobe o arquivo
		
			this.FOL_SetUserResource()
		
			* Excluindo o Arquivo
			this.FOL_SETPARAMETER(1, PLC_Par)
			this.FOL_SETPARAMETER(2, PLC_Ukeyp)
			this.FOL_EditCursor("UXR_PARUKEYP", "UXR_PARUKEYP", "UXRT", "Z", 0, 2)
			
			VLL_Continue = (reccount("UXRT")>0)

	endcase

	this.FOL_CloseTable("UXRT")

	return VLL_Continue

endproc
*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_CloseFileInAllDataSession
*/ Parametros       : PLN_Type  : Alias que o resource está utilizando.
*/ Descrição        : Fechar em todos os DataSession o arquivo informado
*/ Retorno          : .T. ou .F.										  
*/----------------------------------------------------------------------------------------
procedure FOL_CloseFileInAllDataSession
lparameters PLC_AliasFind
	
	VLN_OldDataSession = set("datasession")

	For c = 1 To 250
	  
		Try
		  
			Set Datasession To (c)
			if used(PLC_AliasFind)
				use in (PLC_AliasFind)
			endif
			 
			Catch TO VLO_error
			  
			Exit

			Finally 
		Endtry
	 
	Next 

	set datasession to VLN_OldDataSession
	
endproc

*/----------------------------------------------------------------------------------------
*/ Função Objeto	: FOL_CreateResourceFile
*/ Parametros       : PLL_Resource - Indica se cria o arquivo de resource
*/					: PLL_CutCopy  - Indica se cria o arquivo de Resource Copia e Cola
*/ Descrição        : Criar os arquivos de resource
*/ Retorno          : .T. ou .F.
*/----------------------------------------------------------------------------------------
procedure FOL_CreateResourceFile
lparameters PLL_Resource, PLL_CutCopy

	if PLL_Resource
		select 0
		create table (this.VOC_FileSettings) (id c(100), n1 n(10), n2 n(10), n3 n(10), n4 n(10), n5 n(10), ;
						l1 l( 1), l2 l( 1), l3 l( 1), m1 m)
		index on id tag id
		use
	endif

	if PLL_CutCopy
		select 0
		create table (this.VOC_CutCopyFile) (ukey c(20), id c(50), name c(50), n1 n(10), n2 n(10), n3 n(10), ;
						l1 l( 1), l2 l( 1), l3 l( 1), m1 m, m2 m, m3 m)
		index on id+name   tag id
		index on ukey tag ukey
		use
	Endif

endproc

*/----------------------------------------------------------------------------------------
*/ Função           : FOC_BuildInExpr
*/ Parametros       : PLC_Field - Campo que deve ser comparado
*/					: PLC_ExprList  - Lista com os argumentos a serem separados
*/					: PLC_Separator  - Separador dos argumentos
*/ Descrição        : Cria uma expressão para ser utilizada no comando SQL com a finalidade de substituir procuras
*/                     lentas utilizando contido
*/ Retorno          : Caractér
*/----------------------------------------------------------------------------------------
procedure FOC_BuildInExpr
lparameters PLC_Field, PLC_ExprList, PLC_Separator

	local VLC_Return, VLC_List
	VLC_Return = ""
	
	if empty(PLC_Separator)
		PLC_Separator = ","
	endif

	if !empty(PLC_Field)
		VLC_List = "'" + strtran(PLC_ExprList, PLC_Separator, "','") + "'"
		VLC_Return = PLC_Field + " IN (" + VLC_List + ")"
	endif
	
	return VLC_Return
endproc


*/----------------------------------------------------------------------------------------
*/ Função           : FOC_DateEvalExpr
*/ Parametros       : PLN_StdId - Identificação do padrão
*/					: PLN_FromOrTo  - Expressão para o início (1) ou fim (2) do intervalo
*/					: PLL_DateTimeFormat - Retorna no formato datetime
*/ Descrição        : Cria uma expressão VFP, cujo retorno pode ser avaliado com a função execscript para se
*/					obter as datas/horas programadas
*/ Retorno          : Caractér
*/----------------------------------------------------------------------------------------
procedure FOC_DateEvalExpr
	lparameters PLN_StdId, PLN_FromOrTo, PLL_DateTimeFormat

	local VLC_Return
	VLC_Return = ""

	do case
		case PLN_StdId = 1 && hoje
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return datetime(year(date()), month(date()), day(date()), 23, 59, 59)"
			endcase
		
		case PLN_StdId = 2 && ontem
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 1"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 1)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 1"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() - 1" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 3 && amanhã
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 1"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() + 1)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 1"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 1" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 4 && semana corrente
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - dow(date()) + 1"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - dow(date()) + 1)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - dow(date()) + 7"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() - dow(date()) + 7" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 5 && mês corrente
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return date(year(VLD_RefDate), month(VLD_RefDate), 1)"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), 1, 0, 0, 0)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return gomonth(date(), 1) - day(date())"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = gomonth(date(), 1) - day(date())" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase


		case PLN_StdId = 6 && ano corrente
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return date(year(VLD_RefDate), 1, 1)"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), 1, 1, 0, 0, 0)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date(year(date()), 12, 31)"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return datetime(year(date()), 12, 31, 23, 59, 59)"
			endcase

		case PLN_StdId = 7 && últimos 2 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 1"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 1)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 8 && últimos 3 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 2"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 2)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 9 && últimos 5 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 4"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 4)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 10 && últimos 7 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 6"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 6)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 11 && últimos 10 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 9"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 9)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 12 && últimos 15 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() - 14"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date() - 14)"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 13 && último mês
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return gomonth(date(), -1)"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(gomonth(date(), -1))"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 14 && último ano
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return gomonth(date(), -12)"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(gomonth(date(), -12))"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 15 && proximos 2 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 1"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 1" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 16 && proximos 3 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 2"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 2" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 17 && proximos 5 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 4"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 4" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 18 && proximos 7 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 6"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 6" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 19 && proximos 10 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 9"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 9" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 20 && proximos 15 dias
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date() + 14"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = date() + 14" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 21 && proximo mês
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return gomonth(date(), 1)"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = gomonth(date(), 1)" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

		case PLN_StdId = 22 && proximo ano
			do case
				case PLN_FromOrTo = 1 and empty(PLL_DateTimeFormat)
					VLC_Return = "return date()"
				case PLN_FromOrTo = 1 and !empty(PLL_DateTimeFormat)
					VLC_Return = "return dtot(date())"
				case PLN_FromOrTo = 2 and empty(PLL_DateTimeFormat)
					VLC_Return = "return gomonth(date(), 12)"
				case PLN_FromOrTo = 2 and !empty(PLL_DateTimeFormat)
					VLC_Return = "VLD_RefDate = gomonth(date(), 12)" + chr(13) +;
								 "return datetime(year(VLD_RefDate), month(VLD_RefDate), day(VLD_RefDate), 23, 59, 59)"
			endcase

	endcase

	return VLC_Return

endproc


*/----------------------------------------------------------------------------------------
*/ Função           : FOC_YMEvalExpr
*/ Parametros       : PLN_StdId - Identificação do padrão
*/					: PLN_FromOrTo  - Expressão para o início (1) ou fim (2) do intervalo
*/ Descrição        : Cria uma expressão VFP, cujo retorno pode ser avaliado com a função execscript para se
*/					obter as datas/horas programadas
*/ Retorno          : Caractér
*/----------------------------------------------------------------------------------------
procedure FOC_YMEvalExpr
	lparameters PLN_StdId, PLN_FromOrTo

	local VLC_Return
	VLC_Return = ""

	do case
		case PLN_StdId = 1 && mês corrente
			VLC_Return = "VLD_RefDate = date()" + chr(13) +;
						 "return substr(dtos(VLD_RefDate), 1, 6)"

		case PLN_StdId = 2 && mês passado
			VLC_Return = "VLD_RefDate = gomonth(date(), -1)" + chr(13) +;
						 "return substr(dtos(VLD_RefDate), 1, 6)"

		case PLN_StdId = 3 && próximo mês
			VLC_Return = "VLD_RefDate = gomonth(date(), 1)" + chr(13) +;
						 "return substr(dtos(VLD_RefDate), 1, 6)"

		case PLN_StdId = 4 && ano corrente
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date(year(date()), 1, 1)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date(year(date()), 12, 1)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 5 && Últimos 2 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -1)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 6 && Últimos 3 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -2)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 7 && Últimos 4 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -3)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 8 && Últimos 5 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -4)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 9 && Últimos 6 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -5)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 10 && Últimos 12 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -11)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 11 && Últimos 24 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = gomonth(date(), -23)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 12 && Próximos 2 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 1)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 13 && Próximos 3 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 2)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 14 && Próximos 4 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 3)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 15 && Próximos 5 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 4)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 16 && Próximos 6 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 5)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 17 && Próximos 12 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 11)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

		case PLN_StdId = 18 && Próximos 24 meses
			do case
				case PLN_FromOrTo = 1
					VLC_Return = "VLD_RefDate = date()" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
				case PLN_FromOrTo = 2
					VLC_Return = "VLD_RefDate = gomonth(date(), 23)" + chr(13) +;
								 "return substr(dtos(VLD_RefDate), 1, 6)"
			endcase

	endcase

	return VLC_Return

endproc


*-----------------------------------------------------------------------------------------------------------------------
*- Executa um comando sql na conexão padrão utilizando o ADO
*- Parâmetros:	- PLC_SQLCmd		: Comando sql que será enviado
*-				- PLL_Treated		: Indica se o comando SQL já foi rodado alguma vez (resolvido), matriz VOA_LastSql
*- Retorno   : Lógico
*-----------------------------------------------------------------------------------------------------------------------
procedure FOL_AdoSqlExec
	lparameters PLC_SQLCmd, PLC_ReturnCursor, PLL_Treated

	local VLO_Conn, VLO_Recordset, VLO_Comm, VLO_ObjErr, VLL_Error, VLL_Return, VLL_DebugMode, VLT_Start, VLN_StartSec
	
	if empty(this.VOC_ConnectionString)
		return VLL_Return
	endif

	if empty(PLC_ReturnCursor)
		PLC_ReturnCursor = ""
	endif

	VLL_DebugMode = !empty(this.VOC_DebugSQLCmdFile) and (bitand(Debug_SqlErr, this.VON_DebugType) = Debug_SqlErr)

	if VLL_DebugMode
		if empty(justext(this.VOC_DebugSQLCmdFile))
			this.VOC_DebugSQLCmdFile = alltrim(this.VOC_DebugSQLCmdFile) + ".dbf"
		endif

		if !file(this.VOC_DebugSQLCmdFile)
			create table (this.VOC_DebugSQLCmdFile) (start t, elap_time b(18), sql_cmd m, status c(4), debug_type n)
			this.FOL_CloseTable(juststem(this.VOC_DebugSQLCmdFile))
		endif
		VLT_Start = datetime()
		VLN_StartSec = seconds()
	endif

	VLC_String = this.FOC_SqlStrToExec(PLC_SQLCmd, PLL_Treated, .t.)


	*- Tratamento necessário para desviar os erros do objeto ADO.
	try
		*- Abertura da conexão padrão
		VLO_Conn = createobject("adodb.connection")
		VLO_Conn.open(this.VOC_ConnectionString)
		VLO_Conn.cursorlocation = 3 && adUseClient

		*- Objeto que receberá o Cursor objeto
		VLO_Recordset = createobject("adodb.recordset")
		with VLO_Recordset
			.activeconnection = VLO_Conn
			.cursorlocation = 3 && adUseClient
			.cursortype = 3  && adOpenStatic
			.locktype= 1  && adLockReadOnly
		endwith

		*- Commando SQL a ser executado
		VLO_Comm = createobject("adodb.command")
		VLO_Comm.activeconnection = VLO_Conn
		VLO_Comm.commandtext = VLC_String

		VLO_Recordset.open(VLO_Comm)
		
	catch to VLO_ObjErr
		if vartype(VLO_ObjErr) = 'O'
			VLL_Error = .t. 
		endif

	endtry

	if !VLL_Error
		with VLO_Recordset
			if .fields.count > 0
				dimension VLA_TblStruct[.fields.count, AFIELDS2NDCOL]
				for VLN_FldCounter = 1 to .fields.count
					VLO_FieldObj = .fields(VLN_FldCounter-1)
					VLC_Type = ""
					VLN_Size = 0
					VLN_Precision = 0

					do case
						case VLO_FieldObj.type = 129 && Caracter
							VLC_Type = "C"
							VLN_Size = VLO_FieldObj.definedsize

						case VLO_FieldObj.type = 3 && Inteiro
							VLC_Type = "I"
							VLN_Size = VLO_FieldObj.definedsize

						case VLO_FieldObj.type = 5 && Double
							VLC_Type = "B"
							VLN_Size = 8
							VLN_Precision = 18

						case inlist(VLO_FieldObj.type, 200, 201) && Memo
							VLC_Type = "M"
							VLN_Size = 4

						case VLO_FieldObj.type = 135 && Datetime
							VLC_Type = "T"
							VLN_Size = 14

						case VLO_FieldObj.type = 131 && Numérico/Double
							VLN_Size = VLO_FieldObj.definedsize
							VLN_Precision = VLO_FieldObj.numericscale

							if VLN_Precision <= 8
								VLC_Type = "N"
							else
								VLC_Type = "B"
							endif

						otherwise
							loop
					endcase

					VLA_TblStruct[VLN_FldCounter, 1] = VLO_FieldObj.name
					VLA_TblStruct[VLN_FldCounter, 2] = VLC_Type
					VLA_TblStruct[VLN_FldCounter, 3] = VLN_Size
					VLA_TblStruct[VLN_FldCounter, 4] = VLN_Precision
					VLA_TblStruct[VLN_FldCounter, 5] = bittest(VLO_FieldObj.attributes, adFldIsNullable)
					VLA_TblStruct[VLN_FldCounter, 6] = .f.
					VLA_TblStruct[VLN_FldCounter, 7] = ""
					VLA_TblStruct[VLN_FldCounter, 8] = ""
					VLA_TblStruct[VLN_FldCounter, 9] = ""
					VLA_TblStruct[VLN_FldCounter, 10] = ""
					VLA_TblStruct[VLN_FldCounter, 11] = ""
					VLA_TblStruct[VLN_FldCounter, 12] = ""
					VLA_TblStruct[VLN_FldCounter, 13] = ""
					VLA_TblStruct[VLN_FldCounter, 14] = ""
					VLA_TblStruct[VLN_FldCounter, 15] = ""
					VLA_TblStruct[VLN_FldCounter, 16] = ""
					VLA_TblStruct[VLN_FldCounter, 17] = 0
					VLA_TblStruct[VLN_FldCounter, 18] = 0

				endfor

				create cursor (PLC_ReturnCursor) from array VLA_TblStruct

				if used(PLC_ReturnCursor)
					select (PLC_ReturnCursor)
					this.FOL_CursorBuffering(3)
					for VLN_RecCounter = 1 to .recordcount
						append blank in (PLC_ReturnCursor)
						for VLN_FldCounter = 1 to .fields.count
							replace (VLA_TblStruct[VLN_FldCounter, 1]) with .fields(VLN_FldCounter-1).value in (PLC_ReturnCursor)
						endfor
						.movenext()
					endfor
				endif
				
				VLL_Return = .t.
			endif
		endwith
		VLO_Conn.close()
	endif

	*- Código para debugar velocidade de execução
	if VLL_DebugMode
		VLN_ElapsedSec = seconds() - VLN_StartSec
		VLC_Status = iif(VLL_Return, " OK ", "ERRO")

		use (this.VOC_DebugSQLCmdFile) alias DEBUGFILE in 0 again
		append blank in debugfile
		replace debugfile.start		 with VLT_Start		,;
				debugfile.elap_time	 with VLN_ElapsedSec,;
				debugfile.sql_cmd	 with VLC_String	,;
				debugfile.status	 with VLC_Status	,;
				debugfile.debug_type with Debug_SqlErr	in DEBUGFILE
		use in debugfile
	endif

	return VLL_Return

endproc


*-----------------------------------------------------------------------------------------------------------------------
*- Retorna uma string com o tamanho de um arquivo
*- Parâmetros:	- PLN_FSize	: Tamanho do arquivo em bytes
*- Retorno   : Caractér
*-----------------------------------------------------------------------------------------------------------------------
procedure FOC_FileSizeStr
	lparameters PLN_FSize

	local VLC_Return

	do case
		case PLN_FSize <= 1023
			VLC_Return = transform(PLN_FSize, "@ 9,999") + " bytes"
		case between(PLN_FSize, 1024, (2^20)-1)
			VLC_Return = transform(PLN_FSize/1024, "@ 9,999.99") + " Kb"
		otherwise
			VLC_Return = transform(PLN_FSize/(2^20), "@ 9,999.99") + " Mb"
	endcase

	return alltrim(VLC_Return)
endproc


*** <summary>
*** Muda os atributos de um arquivo
*-	R ou -R : Apenas Leitura (Readonly)
*-	H ou -H : Invisível (Hidden)
*-	S ou -S : Sistema (System)
*-	D ou -D : Diretório (Directory)
*-	A ou -A : Arquivo (Archive)
*-	N ou -N : Normal (Normal)
*-	T ou -T : Temporário (Temporary)
*-	C ou -C : Compactado (Compressed)
*** </summary>
*** <param name="PLC_FileFullPath">Caminho completo do arquivo</param>
*** <param name="PLC_Attributes">Lista de atributos separador por espaço, utilizando - para retirar o atributo</param>
*** <remarks></remarks>
procedure FOL_SetFileAttributes
	lparameters PLC_FileFullPath, PLC_Attributes


this.


	local VLN_CurFileAttributes, VLN_NewFileAttributes

	*- Lê os atributos correntes do arquivo read current attributes for this file 
	VLN_CurFileAttributes = GetFileAttributes(PLC_FileFullPath)

	if VLN_CurFileAttributes = -1 or empty(PLC_Attributes)
		*- Arquivo não existe
	    return .f.
	endif
	
	VLN_NewFileAttributes = VLN_CurFileAttributes
	PLC_Attributes = upper(PLC_Attributes)

	if at("-R", PLC_Attributes) > 0 
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_READONLY
		endif
	else
		if at("R", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_READONLY
		endif
	endif

	if at("-H", PLC_Attributes) > 0 
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_HIDDEN
		endif
	else
		if at("H", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_HIDDEN
		endif
	endif

	if at("-S", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_SYSTEM
		endif
	else
		if at("S", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_SYSTEM
		endif
	endif

	if at("-D", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_DIRECTORY
		endif
	else
		if at("D", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_DIRECTORY
		endif
	endif

	if at("-A", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_ARCHIVE
		endif
	else
		if at("A", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_ARCHIVE
		endif
	endif

	if at("-N", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_NORMAL
		endif
	else
		if at("N", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_NORMAL
		endif
	endif

	if at("-T", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_TEMPORARY
		endif
	else
		if at("T", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_TEMPORARY
		endif
	endif

	if at("-C", PLC_Attributes) > 0
		if bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED
			VLN_NewFileAttributes = VLN_NewFileAttributes - FILE_ATTRIBUTE_COMPRESSED
		endif
	else
		if at("C", PLC_Attributes) > 0 and !bitand(VLN_NewFileAttributes, FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED
			VLN_NewFileAttributes = VLN_NewFileAttributes + FILE_ATTRIBUTE_COMPRESSED
		endif
	endif

    *- Coloca os novos atributos no arquivo
	if VLN_NewFileAttributes <> VLN_CurFileAttributes
	    =SetFileAttributes(PLC_FileFullPath, VLN_NewFileAttributes)
	endif 


	
	
	return .t.
endproc


protected function Rodrigo_Bruscain

endproc 

Enddefine


Defi Clas sgo_teste as Form
	VON_prop1 = 0					&&-nao aparece no intellisense
	VON_prop2 = 0
	VON_prop3 = 0
	VON_prop4 = 0

	Function FOL_Func1
		this.
	Endfunc 
	
	Protected Function FOC_Func2
	Endfunc 
Enddefine  


Defi Clas sgo_teste2 as Form
	VON_prop1 = 0
	VON_prop2 = 0
	VON_prop3 = 0
	VON_prop4 = 0

	Function FOL_Func1
		
	Endfunc 

	Function FOL_Func1
	Endfunc 
	
	Protected Function FOC_Func2
	Endfunc 
Enddefine  


function teste_A
	ggg = createobject("form")
	with ggg
		
endfunc  


procedure teste_B

endfunc 

