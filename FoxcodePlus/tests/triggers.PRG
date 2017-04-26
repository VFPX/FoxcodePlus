*/------------------------------------------------------------------------------------------/*
*/ Função  			: triggers             		                                            /*
*/ Descrição		: Programa principal que executa a chamada de todos os 				    /*
*/                    específicos dentro do sistema.                                        /*
*/ Retorno			: sem retorno                                                           /*
*/ Empresa			: Clone©											  			        /*
*/------------------------------------------------------------------------------------------/*
lparameters PLO_Form, PLO_Source, PLC_CodeId
local VLC_FormName, VLN_Position, VLC_Exact



 

 
VLC_FormName = iif(!isnull(PLO_Form), upper(PLO_Form.name), [NULL])

if 'RIGHTCLICK' $ PLC_CodeId and VGO_Custom.VOC_USERCODE $ "STAR_"
	wait window PLO_Source.name + chr(13) +;
		PLO_Source.VOC_FullName
	_cliptext = proper(PLO_Source.VOC_FullName)
endif

*- Estancia classe da Clone de funcoes e propriedades
if vartype(VGO_clone) <> 'O'
	public VGO_clone
	VGO_clone = newobject('VGO_clone', 'VGO_clone.vcx')
endif

*- Trata variaveis
PLC_CodeId = upper(alltrim(PLC_CodeId))
PPO_form = VLC_FormName

public VLL_PROCESSADOINTG, VLC_HISTORINT
VLL_PROCESSADOINTG = .f. 
VLC_HISTORINT = ""

*- variavel auxiliar para correção do desconto e variavel publica para recalcular o desconto geral do pedido de venda
public VPL_CORRIGE_DESCONTO, VPL_RecalculaDescGeral





*- Inicia verificacao de pontos do trigger
do case
	case VLC_FormName == "LOGIN"
		*-- carrega o arquivo de procedures
		do case
						 
			case PLC_CodeId $ "LOGIN.UNLOAD"
				do prg_arrays.fxp
				set procedure to prg_classes_especificas.fxp additive

				*- Adição da Libera de funções padrão da area de serviços da STARSOFT
				public VGO_ss
				VGO_ss = newobject('cgo_ss','sgo_ss.fxp')
					
				
				*- Somente para administradores do sistema ***\\***
				if VGO_Custom.VOC_USERCODE $ "STAR_"
					on key label CTRL+F2 activ window command
					on key label CTRL+F3 set
					on key label CTRL+F4 set step on
					on key label alt+F9 browse normal
					on key label alt+F12 suspend
					on key label alt+F10 _cliptext = vgo_dat.voc_sqlerror
		 
					set procedure to backupgenfix.fxp additive
					on key label alt+C FLL_CopyGenFix()
					on key label alt+R FLL_RestoreGenFix()

						
						
					set procedure to sgo_y07vss.fxp additive
					on key label alt+V FLL_RetiraArquivos()
				endif
				
				*- Habilita o campo data de entrada na nota de entrada.
				on key label ALT+0 do tgr_e10_014_d_enabled.fxp with _screen.ActiveForm 

				*- Garante não dar timeout de conexao com o banco
				do tgr_connect.FXP

				*- Adiciona propriedade com o ID do usuario
				VGO_Custom.FOL_cSQLexec("SELECT USR_CODE, USR_611_N FROM STAR_DATA@USR (NOLOCK) WHERE USR_UKEY LIKE '%"+alltrim(VGO_Custom.VOC_USERCODE)+"%'","TmpUser")
				VGO_clone.voc_userID = alltrim(TmpUser.USR_CODE)

				

				if VGO_toolbar.voa_toolbars[6,1]._ssm_webb.visible
					VGO_toolbar.voa_toolbars[6,1]._ssm_webb.visible = ( VGO_Clone.VOC_userID $ "ADMINISTRATOR" )
				endif

				if TmpUser.USR_611_N = 0
					do tgr_sessao.prg with PLO_Form
				endif
				VGO_dat.FOL_CloseTable("TmpUser")

				do proc0049

				*- Desabilito impressao de filtros (resume) após imprimir qualquer relatorio (economia de papél)
				if !empty(VGO_Gen.VOC_DirSettings)
					if !used("UserResource")
						use (VGO_Gen.VOC_FileSettings) in 0 alias userresource
					endif
					replace userresource.l3 with .f. in userresource for id = "frxformat-01"
					VGO_Dat.FOL_CloseTable("userresource")
				endif

				*- funcionalidades do custo
				do triggers_custos with PLO_Form, PLO_Source, PLC_CodeId
				
				*- Desabilita o botão de troca de empresas para o usuário
				VGO_Toolbar.VOA_Toolbars[6,1]._ssm_cias.visible = .F.
		endcase

	*- Tudo isso abaixo é porque a tela não salva no local selecionado.
	case VLC_FormName == "H11_001"
		do case
			case PLC_CodeId == "H11_001.INIT"
				PLO_Form.addproperty("VOC_InitialFileName","")
				PLO_Form.addproperty("VOL_ChangeOption",.f.)
				PLO_Form.addproperty("VOL_ChangeDestiny", .f.)
				PLO_Form.addproperty("VOL_WasClickedInDoIt", .f.)
				PLO_Form.addproperty("VOC_DestinyFileName","")
				do tgr_h11_001_init.fxp with PLO_Form

			case PLC_CodeId = upper("H11_001.container1.h11_ukey.textalias1.lostfocus")
				PLO_Form.VOL_ChangeOption = .t.

			case PLC_CodeId = upper("H11_001.container1.document.document.gotfocus")
				if PLO_Form.VOL_ChangeOption = .t.
					PLO_Form.VOL_ChangeOption = .f.
					PLO_Form.VOC_InitialFileName = PLO_Form.container1.document.document.value
				endif
				PLO_Form.VOL_ChangeDestiny = .t.

			case PLC_CodeId = upper("H11_001.container1.document.FOL_update")
				if PLO_Form.VOL_ChangeDestiny
					PLO_Form.VOL_ChangeDestiny = .f.
					PLO_Form.VOC_DestinyFileName = forceext(PLO_Form.container1.document.document.value,"TXT")
				endif

			case PLC_CodeId = upper("H11_001.doit.click")
				PLO_Form.VOL_WasClickedInDoIt = .t.
		endcase


	case VLC_FormName == "A03"				&&-Tela de Clientes
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				do tgr_a03_listaapp.fxp with PLO_Form
				
			case PLC_CodeId == "A03.LOAD"
*!*					PLO_Form.addproperty("Voc_MyExecuteCursor")
*!*					PLO_Form.addproperty("Voc_LastA03Ukey603")
*!*					PLO_Form.addproperty("Voc_LastA03Ukey605")
*!*					PLO_Form.Voc_MyExecuteCursor 	= ""
*!*					PLO_Form.Voc_LastA03Ukey603		= ""
*!*					PLO_Form.Voc_LastA03Ukey605		= ""
				do tgr_A03_load.fxp with PLO_Form


			case PLC_CodeId == "A03.FOL_EXECUTECURSOR"
				if !PLO_Form.Voc_MyExecuteCursor == A03.Ukey or vgo_gen.voa_TriggerParameters[1]
					PLO_Form.Voc_MyExecuteCursor = A03.Ukey
					with Vgo_Dat
						** Canais
						.FOL_SqlExec("Select * from STAR_DATA@AA601 (NOLOCK)","AA601_all")
						** Localizações
						.FOL_SqlExec("Select * from STAR_DATA@AA602 (NOLOCK)","AA602_all")
						** Canais do Cliente
						.Voa_Index[1] = "AA601_Ukey"
						.FOL_SetParameter(1,A03.Ukey)
						.FOL_EditCursor("AA603_FRM","AA603_FRM","AA603T","Z",1,1)
						** Localizações do Cliente
						.Voa_Index[1] = "AA602_Ukey"
						.FOL_SetParameter(1,A03.Ukey)
						.FOL_EditCursor("AA605_FRM","AA605_FRM","AA605T","Z",1,1)
					endwith

					PLO_Form.Voc_LastA03Ukey603		= ""
					PLO_Form.Voc_LastA03Ukey605		= ""
					if type("PLO_Form.Pageframe1.Page5.Moverlists1") = "O"
						=PLO_Form.Pageframe1.Page5.Moverlists1.FOL_CreateList()
					endif
				endif

			case PLC_CodeId == "A03.INIT"
				do tgr_A03_init.fxp with PLO_Form

			case PLC_CodeId == "A03.REFRESH"
				do tgr_A03_REFRESH.fxp with PLO_Form

				with PLO_Form.PAGEFRAME1.PAGE2.A03_A02_UKEY.TEXTALIAS1
					.VOC_Searchname= "A02_ESP"
					.VOC_cursorcat = "A02_ESP"
					.VOC_CUrsorcatrecords = "A02_CLI_COUNT"
				endwith

			case PLC_CodeId $ "A03.FOL_EDIT"
				do tgr_A03_edit.fxp with PLO_Form

			case PLC_CodeId $ "A03.FOL_SAVE"
				do tgr_A03_FOL_Save.fxp with PLO_Form

			case PLC_CodeId $ "A03.FOL_AFTERSAVE"
				do tgr_A03_FOL_AfterSave.fxp with PLO_Form

			case PLC_CodeId $ "A03.FOL_BEFOREDELETE"
				do tgr_A03_FOL_BeforeDelete.fxp with PLO_Form

			case PLO_Form.VOL_EDITING and VGO_Clone.VOC_userID $ "AGNALDO/CELIA/MAURO/JOSE/FATIMA/REGIS"
				**	 Objetos que não desabilitaram diretamente pela classe
				PLO_Form.Pageframe1.Page4.a03_a03_020_n.enabled = .f.
				PLO_Form.Pageframe1.Page4.a03_a03_020_n.vol_Enabled = .f.
		endcase
	
	case VLC_FormName == "A08"
		do case
			case PLC_CodeId $ "A08.INIT"
				do tgr_a08_init.fxp with PLO_Form
		endcase
	
	case VLC_FormName == "A11" && Feitosa Productions (nao sei pra que to fazendo isso, mas....)
		do case
			case PLC_CodeId $ "A11.INIT"
				do tgr_a11_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "A13"				&&-Condicao de pagamentos
		do case
			case PLC_CodeId = "A13.INIT"
				do tgr_A13_init.fxp with PLO_Form
		
			case PLC_CodeId == "LOADTOOLBARAPPS"
				VGO_CUSTOM.FOL_cAddToolbarApps ("Clientes para esta condição", "VGO_Gen.FOL_DoForm('A13_Cli')", PLO_Form.VOL_Editing and reccount("A13") > 0)
		endcase
		
	case VLC_FormName == "A14"
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"
				VGO_CUSTOM.FOL_cAddToolbarApps ("LAYOUT de Tabelas para Cálculos de Fretes",'VGO_DAT.FOL_DoNormalForm([JJ612])',!empty(A14.UKEY) )
				VGO_CUSTOM.FOL_cAddToolbarApps ("CONFIGURAÇÃO de Valores para Cálculos de Fretes",'VGO_DAT.FOL_DoNormalForm([JJ613])',!empty(A14.UKEY) )
		endcase

	case VLC_FormName == "A21" 				&&-Cadastro de Tipo de Documento
		do case
			case PLC_CodeId $ "A21.INIT"
				do tgr_a21_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "A23" 				&&-Cadastro de Estados
		do case
			case PLC_CodeId $ "A23.INIT"
				do tgr_a23_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "A24" 				&&-Cadastro de Cidades
		do case
			case PLC_CodeId $ "A24.INIT"
				do tgr_a24_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "A33" 				&&-Cadastro de vendedores
		do case
			case PLC_CodeId $ "A33.INIT"
				do tgr_a33_init.fxp with PLO_Form
			case PLC_CodeId $ "A33.FOL_BEFORESAVE"
				do tgr_A33_FOL_BeforeSave.fxp with PLO_Form
				
			case PLC_CodeId $ "A33.REFRESH"
				with PLO_Form.paGEFRAME1.pAGE3.a33_a02_ukey.tEXTALIAS1
					.VOC_Searchname= "A02_VENDEDOR"
					.VOC_cursorcat = "A02_VENDEDOR"
					.VOC_CUrsorcatrecords = "A02_VENDEDOR_COUNT"
				endwith

			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				*- Representatividade
				*VGO_Custom.FOL_cAddtoolbarapps ("Representatividade", "VGO_Custom.FOL_cDoForm('AA606.scx')",!PLO_Form.VOL_Editing or PLO_Form.VOL_Adding)
				VGO_Custom.FOL_cAddtoolbarapps ("Representatividade (Gerente)", "VGO_Custom.FOL_cDoForm('AA606.scx')", !PLO_Form.VOL_Editing and reccount("A33") > 0 and a33.array_003 = 1)
				VGO_Custom.FOL_cAddtoolbarapps ("Representatividade (Zona)", "VGO_Custom.FOL_cDoForm('AA608.scx')", !PLO_Form.VOL_Editing and reccount("A33") > 0 and a33.array_003 = 1)
		endcase

	case VLC_FormName == "B04_001" 				&&-Cadastro de Rateios
		do case
			case PLC_CodeId $ "B04_001.FOL_BEFORESAVE"
				do tgr_b04_001_beforesave.fxp with PLO_Form
		endcase

	case VLC_FORMNAME == "B06"
		do case
			case PLC_CodeId $ "LOADTOOLBARAPPS"
				vgo_custom.FOL_cAddtoolbarapps ("Alteração das datas do lançamento contábil", 'vgo_dat.FOL_donormalform([B06_ALTERADATAS])', PLO_Form.VOL_Editing)
		endcase

	case VLC_FormName =="B11"
		do case
			case PLC_CodeId $ "LOADTOOLBARAPPS"
				vgo_custom.FOL_cAddtoolbarapps ("Alteração do Periodo da Conta Contábil", 'vgo_dat.FOL_donormalform([B11_ALTERAPERIODO])', PLO_Form.VOL_Editing)
		endcase

	case VLC_FormName == "D04" 				&&-Itens do estoque
		do case			
			case PLC_CodeId == upper(alltrim("D04.PAGEFRAME1.PAGE6.D04_D04_629_N.CLICK"))
				if PLO_Form.PAGEFRAME1.PAGE6.D04_D04_629_N.value = 1
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.VOL_ENABLED = .f.
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.value = 0

					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.refresh()
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_629_N.refresh()
				else
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.VOL_ENABLED = .t.
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.value = 0

					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N.refresh()
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_629_N.refresh()
				endif			
			
			case PLC_CodeId == "D04.FOL_EXECUTECURSOR"
				if !PLO_Form.Voc_MyExecuteCursor == D04.Ukey or vgo_gen.voa_TriggerParameters[1]
					PLO_Form.Voc_MyExecuteCursor = D04.Ukey
					with Vgo_Dat
						** Canais
						.FOL_SqlExec("Select * from STAR_DATA@AA601 (NOLOCK)","AA601_all")
						** Canais do Produto
						.Voa_Index[1] = "AA601_Ukey"
						.FOL_SetParameter(1,D04.Ukey)
						.FOL_EditCursor("AA604_FRM","AA604_FRM","AA604T","Z",1,1)
					endwith

					PLO_Form.Voc_LastD04Ukey604		= ""						
*!*						if type("PLO_Form.Pageframe1.Page8.Moverlists1") = "O"
*!*							=PLO_Form.Pageframe1.Page8.Moverlists1.FOL_CreateList()
*!*						endif				
					do TGR_D04_REFRESH
				endif
			
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Últimas Vendas do Item", "VGO_Custom.FOL_cDoForm('D04_ULTIMASVENDAS')", not VGO_Dat.FOL_AcessMenu("D04_ULTIMASVENDAS") and !PLO_Form.VOL_EDITING and not empty(PLO_Form.PAGEFRAME1.PAGE1.D04_D04_001_C.value))
				VGO_Custom.FOL_cAddtoolbarapps("Edição do Preço deste Item", "VGO_Custom.FOL_cDoForm('D04_ALTERAPRECO')", !PLO_Form.VOL_EDITING and not empty(PLO_Form.PAGEFRAME1.PAGE1.D04_D04_001_C.value))
				VGO_Custom.FOL_cAddtoolbarapps("Análise de Venda do Item", "VGO_Custom.FOL_cDoForm('D04_ANAL_ITEM')", !PLO_Form.VOL_EDITING and not empty(PLO_Form.PAGEFRAME1.PAGE1.D04_D04_001_C.value))
			case PLC_CodeId $ "D04.INIT"
				do tgr_d04_init.fxp with PLO_Form
			case upper(PLC_CodeId) = upper("d04.pageframe1.page1.d04_d04_001_c.lostfocus")
				=PLO_Form.pageframe1.page6.d04_d04_621_d.refresh()
				=PLO_Form.pageframe1.page6.iABC.refresh()
			case PLC_CodeId $ "D04.FOL_SAVE"
				do tgr_d04_Save.fxp with PLO_Form
			case upper(PLC_CodeId) = upper("d04.pageframe1.page1.d04_d04_008_c.Lostfocus") and PLO_Form.Vol_adding
				replace D04.D04_600_C with D04.D04_008_C
				*				PLO_Form.pageframe1.page2.destecnica.refresh
				PLO_Form.pageframe1.page6.destecnica.refresh
			case PLO_Form.VOL_EDITING
				do tgr_D04_edit.fxp with PLO_Form
				
				if PLO_Form.PAGEFRAME1.PAGE6.D04_D04_629_N.value = 1
					with PLO_Form.PAGEFRAME1.PAGE6.D04_D04_630_N
						.VOL_ENABLED = .f.
						.value = 0

						.refresh()
					endwith
					
					PLO_Form.PAGEFRAME1.PAGE6.D04_D04_629_N.refresh()
				endif
				
			case PLC_CodeId $ "D04.LOAD"
				do TGR_D04_LOAD  with PLO_Form
		endcase

	case VLC_FORMNAME == "D07"
		do case
			case PLC_CodeId == "D07.INIT"
				do TGR_D07_INIT with PLO_Form
		endcase

	case VLC_FORMNAME == "D12"
		do case
			case PLC_CodeId == "D12.INIT"
				if FLL_VerifAcesso("EFET_INVEN")
					PLO_Form.d12_d12_015_n.visible = .f.
					PLO_Form.COMMANDBUTTON1.left = PLO_Form.d12_D12_015_N.left
				endif
			case PLC_CodeId == "D12.FOL_SAVE"
				do tgr_d12_save.prg
			case PLC_CodeId == "D12.REFRESH"
				do tgr_d12_refresh.prg
		endcase

	case VLC_FormName == "D13_001"		&&-Capa Entrada de Materiais filipe 05092008
		do case
			case PLC_CodeId = "D13_001.REFRESH"  && Filipe Oliveira - pesquisa clientes ativos e em estudo. 05/09/2008
				with PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
		endcase

	case VLC_FormName == "D13_002"		&&-Capa Saida de Materiais filipe 08092008
		do case
			case PLC_CodeId = "D13_002.REFRESH"  && Filipe Oliveira - pesquisa clientes ativos e em estudo. 08/09/2008
				with PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
		endcase

	case VLC_FormName == "D13_005"		&&-Capa Transferencia de Locais
		do case
			case PLC_CodeId == "D13_005.LOAD"
				PLO_Form.addproperty("VOC_MyLastExecuteCursor")
				PLO_Form.addproperty("VOL_MyExecute")
				PLO_Form.VOC_MyLastExecuteCursor = ""
				PLO_Form.VOL_MyExecute = .t.

			case PLC_CodeId == "D13_005.FOL_EXECUTECURSOR"
				if PLO_Form.VOL_MyExecute or !(PLO_Form.VOC_MyLastExecuteCursor == nvl(d13.ukey, ""))
					PLO_Form.VOC_MyLastExecuteCursor = d13.ukey
					PLO_Form.VOL_MyExecute = .f.
					with VGO_Dat
						*- Kit (produto pai)
						.FOL_SetParameter(1, "D13")
						.FOL_SetParameter(2, d13.ukey)
						.VOA_Index[1] = "jj629t.jj629_par + Nvl(jj629t.jj629_ukeyp,Space(20))"
						.FOL_EditCursor("JJ629", "JJ629", "JJ629T", "Z", 1, 2)
					endwith
				endif

			case PLC_CodeId = "D13_005.FOL_CANCEL"
				PLO_Form.VOL_MyExecute = .t.		&&- Refaço os cursores especificos do D13_005.FOL_EXECUTECURSOR

			case PLC_CodeId = "D13_005.INIT"
				*- Configura gravação para os kits
				PLO_Form.FOL_SetTable(2, .t., null, "JJ629T", "JJ629", "Ukey", "JJ629")

			case PLC_CodeId == "D13_005.FOL_SAVE"
				do tgr_D13_005_SAVE.fxp with PLO_Form 	&&- Gravo no código de destino do item

			case PLC_CodeId == "D13_005.FOL_AFTERSAVE"
				if VGO_Gen.VOL_TriggerReturn
					PLO_Form.VOL_MyExecute = .t.		    &&- Refaço os cursores especificos do D13_005.FOL_EXECUTECURSOR
				endif

				do tgr_D13_005_AFTERSAVE.fxp with PLO_Form 	&&- Atualizo os acumuladores após o save porque a qtd de entrada ainda não está no banco durante o save. 	

		endcase

	case VLC_FormName == "D13_007"		&&-Capa Requisição de Materiais
		do case
			case  PLC_CodeId == "D13_007.D13_D13_003_C.INIT"
				PLO_Form.D13_D13_003_C.vol_chkifempty = .t.
		endcase

	case VLC_FormName == "D13_008"		&&-Capa da reserva da materiais
		do case
			case PLC_CodeId == "D13_008.LOAD"
				PLO_Form.addproperty("VOC_MyLastExecuteCursor")
				PLO_Form.addproperty("VOL_MyExecute")
				PLO_Form.VOC_MyLastExecuteCursor = ""
				PLO_Form.VOL_MyExecute = .t.

			case PLC_CodeId == "D13_008.FOL_EXECUTECURSOR"
				if PLO_Form.VOL_MyExecute or !(PLO_Form.VOC_MyLastExecuteCursor == nvl(d13.ukey, ""))
					PLO_Form.VOC_MyLastExecuteCursor = d13.ukey
					PLO_Form.VOL_MyExecute = .f.
					with VGO_Dat
						*- Kit (produto pai)
						.FOL_SetParameter(1, "D13")
						.FOL_SetParameter(2, d13.ukey)
						.VOA_Index[1] = "jj629t.jj629_par + Nvl(jj629t.jj629_ukeyp,Space(20))"
						.FOL_EditCursor("JJ629", "JJ629", "JJ629T", "Z", 1, 2)
					endwith
				endif

			case PLC_CodeId = "D13_008.FOL_CANCEL"
				PLO_Form.VOL_MyExecute = .t.		&&- Refaço os cursores especificos do D13_008.FOL_EXECUTECURSOR

			case PLC_CodeId = "D13_008.INIT"
				*- Configura gravação para os kits
				PLO_Form.FOL_SetTable(2, .t., null, "JJ629T", "JJ629", "Ukey", "JJ629")

			case PLC_CodeId == "D13_008.FOL_AFTERSAVE"
				if VGO_Gen.VOL_TriggerReturn
					PLO_Form.VOL_MyExecute = .t.		    &&- Refaço os cursores especificos do D13_008.FOL_EXECUTECURSOR
				endif
		endcase

	case VLC_FormName == "D13_017"		&&-Capa Alocação para Produção
		do case
			case  PLC_CodeId == "D13_017.INIT"
				do tgr_D13_017_init.fxp with PLO_Form && Apaga a alocação na lista de Materiais da OP
			case  PLC_CodeId == "D13_017.FOL_AFTERSAVE"
				do tgr_D13_017_SAVE.fxp && Grava alocação na lista de materiais da OP
		endcase


	case VLC_FormName == "D14_001"			&&-Item da entrada de Materiais
		do case
			case PLC_CodeId $ "D14_001.D14_D04_UKEY.TEXTALIAS1.LOSTFOCUS"
				if empty(alltrim(PLO_Form.TEXTALIASLABEL_D34.TEXTALIAS1.value))
					PLO_Form.TEXTALIASLABEL_D34.TEXTALIAS1.setfocus()
				endif

			case  PLC_CodeId $ "D14_001.D14_D14_002_B.SPINNER1.GOTFOCUS"	&& foco da quantidade
				*- Carrega o valor do custo no Valor unitário se estive como zerado
				do tgr_d14_001_d14_002_b_gotfocus.fxp with PLO_Form

		endcase
	
	case VLC_FormName == "D14_004"			&&-Item da previsão entrada de Materiais
		do case
			case PLC_CodeId $ "D14_004.D14_D14_003_B.SPINNER1.GOTFOCUS"
				if d14t.d14_003_b = 0				
					*- Traz o valor da última nota de entrada para o item selecionado
					VGO_Gen.FOL_SqlExec("select top 1 e11.e11_005_b from STAR_DATA@e11 e11 (nolock) inner join STAR_DATA@e10 e10 (nolock) on e11.e10_ukey = e10.ukey where e11.d04_ukey = ?d14t.d04_ukey order by e10.e10_014_d desc, e10.timestamp desc", "TmpE11")
					replace d14_003_b with TmpE11.e11_005_b in d14t
					VGO_Gen.FOL_CloseTable("TmpE11")
				endif
		endcase

	* Chamada espeficica para timer da tela de calculo de estoque e calculo de custos - 30042010
	case VLC_FormName == "D15"
		do case
			case PLC_CodeId == "D15.INIT"
				do ss_d15-timer.prg with PLO_Form 
	
			&&- Colocando uma trava para não fazer calculo de estoque anterior a 01/01/2011
			CASE PLC_CodeId == "D15.CONTAINER3.INITIAL_DATE.LOSTFOCUS"	
				do tgr_d15_initial_date_lostfocus.prg WITH PLO_Form		
		endcase
	
	case VLC_FormName == "D15_003"
		do case
			&&- Colocando uma trava para não fazer calculo de estoque anterior a 01/01/2011
			CASE PLC_CodeId == "D15_003.CONTAINER.INITIAL_DATE.LOSTFOCUS"				
				do tgr_d15_003_initial_date_lostfocus.prg WITH PLO_Form		
		endcase	
	
	case VLC_FormName == "V22_001"
		do case
			case PLC_CodeId == "V22_001.INIT"
				do ss_V22_001-timer.prg with PLO_Form		
		
			case PLC_CodeId == "V22_001.CONTAINER1.TEXTALIASLABEL_V03.TEXTALIAS1.LOSTFOCUS"	
				do tgr_v22_001_textaliaslabel_v03_lostfocus.prg WITH PLO_Form		
		endcase


	case VLC_FormName == "V22_001"
		do case
			case PLC_CodeId == "V22_001.INIT"
				do ss_V22_001-timer.prg with PLO_Form
		endcase

	case VLC_FormName == "D16"				&&-Classificais fiscais
		do case
			case PLC_CodeId $ "D16.REFRESH"
				if VGO_Gen.VOC_CompanyCode = "SRTRU"
					PLO_Form.D16_D16_601_B.enabled = PLO_Form.VOL_EDITING
					PLO_Form.D16_D16_602_B.enabled = PLO_Form.VOL_EDITING
				endif
				
			case PLC_CodeId $ "D16.INIT"
				do TGR_D16_INIT with PLO_Form
		endcase

	case VLC_FormName == "D22_003"			&&-Itens Previsão de Saída
		do case
			case PLC_CodeId $ "D22_003.INIT"
				do tgr_d22_003_init.fxp with PLO_Form
				
			case PLC_CodeId $ "D22_003.LOAD"	
				PLO_Form.height = PLO_Form.height + 23
		endcase

	case VLC_FormName == "D22_005"			&&-Itens Transferencia de locais
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Lançamento de Kits", "VGO_Custom.FOL_cDoForm('jj629_1', _screen.activeform)",  not PLO_Form.VOL_EDITING and not _screen.activeform.VOL_NoEdit)

			case PLC_CodeId == "D22_005.TEXTALIASLABEL_D04.FOL_UPDATE" and wontop()="D22_005"
				if "." $ PLO_Form.TEXTALIASLABEL_D04.TEXTALIAS1.value and empty(nvl(d22t.jj629_Ukey,""))
					messagebox("Lançamento manual de kit ou parte de kit não permitido."+chr(13)+;
						"Favor utilizar a tela de lançamento de Kits na lista de aplicações",48,"Atenção")

					replace d04_ukey with null in d22t
					PLO_Form.TEXTALIASLABEL_D04.refresh()
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId $ "D22_005.INIT"
				do tgr_D22_005_init.fxp with PLO_Form 			&&-Cria os campos necessários para o Item Destino

			case "D22_005.TEXTALIASLABEL_D07" $ PLC_CodeId or "REFRESH" $ PLC_CodeId
				*- Caso seja item filho proveniente de kit não permito alterar o código do item, quantidade e local
				if not empty(nvl(d22t.jj629_ukey,""))
					PLO_Form.TEXTALIASLABEL_D07.TEXTALIAS1.enabled = .f.
					PLO_Form.TEXTALIASLABEL_D04.TEXTALIAS1.enabled = .f.
					PLO_Form.D22_D22_002_B.SPINNER1.enabled = .f.
				endif

				do tgr_d22_005_d04_fol_update.fxp with PLO_Form 			&&*- Travando o campo de valor unitário caso o item tenha seu código maior de 5 posições

			case PLC_CodeId == "D22_005.FOL_ADDNEW"
				do tgr_d22_005_EDIT.fxp with PLO_Form

			case  PLC_CodeId == "D22_005.D22_D22_002_B.SPINNER1.GOTFOCUS"	&& foco da quantidade
				*- Carrega o valor do custo no Valor unitário se estive como zerado
				do tgr_d22_005_d22_002_b_gotfocus.fxp with PLO_Form

			case PLC_CodeId == "D22_005.FOL_BEFORESAVE"
				do tgr_d22_005_BEFORESAVE.fxp

			case PLC_CodeId $ "D22_005.FOL_BEFOREDELETE" and wontop()="D22_005"		&&- Uso o Wontop() se não a trigger também é acionada pela tela de Lançamento de Kit
				if not empty(nvl(d22t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId $ "D22_005.D22_D22_003_B.SPINNER1.GOTFOCUS"
				if d22t.d22_003_b = 0				
					*- Traz o valor da última nota de entrada para o item selecionado
					VGO_Gen.FOL_SqlExec("select top 1 e11.e11_005_b from STAR_DATA@e11 e11 (nolock) inner join STAR_DATA@e10 e10 (nolock) on e11.e10_ukey = e10.ukey where e11.d04_ukey = ?d22t.d04_ukey order by e10.e10_014_d desc, e10.timestamp desc", "TmpE11")
					replace d22_003_b with TmpE11.e11_005_b in d22t
					VGO_Gen.FOL_CloseTable("TmpE11")
				endif		
		endcase

	case VLC_FormName == "D22_008"			&&-Itens da reserva de materiais
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Lançamento de Kits", "VGO_Custom.FOL_cDoForm('jj629_2', _screen.activeform)",  not PLO_Form.VOL_EDITING and not _screen.activeform.VOL_NoEdit)

			case PLC_CodeId == "D22_008.D22_D04_UKEY.FOL_UPDATE" and wontop()="D22_008"
				if "." $ PLO_Form.D22_D04_UKEY.TEXTALIAS1.value and empty(nvl(d22t.jj629_Ukey,""))
					messagebox("Lançamento manual de kit ou parte de kit não permitido."+chr(13)+;
						"Favor utilizar a tela de lançamento de Kits na lista de aplicações",48,"Atenção")

					replace d04_ukey with null in d22t
					PLO_Form.D22_D04_UKEY.refresh()
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId == "D22_008.FOL_EDIT"
				*- Caso seja item filho proveniente de kit não permito alterar o código do item, quantidade e local
				if not empty(nvl(d22t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser alterado ou excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId $ "D22_008.FOL_BEFOREDELETE" and wontop()="D22_008"		&&- Uso o Wontop() se não a trigger também é acionada pela tela de Lançamento de Kit
				*- Caso seja item filho proveniente de kit não permito alterar o código do item, quantidade e local
				if not empty(nvl(d22t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser alterado ou excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif
		endcase

	case VLC_FormName == "D22_017" && Itens de Alocação da produção
		do case
			case PLC_CodeId $ "D22_017.FOL_EDIT//D22_017.FOL_BEFOREDELETE"
				do tgr_d22_017_Fol_Edit.fxp && Não permito alterar registro integrado.
				
			case  PLC_CodeId == "D22_017.D22_D22_002_B.SPINNER1.GOTFOCUS"	&& foco da quantidade
				*- Carrega o valor do custo no Valor unitário se estive como zerado
				do tgr_d22_017_d22_002_b_gotfocus.fxp with PLO_Form
		endcase

	case VLC_FORMNAME == "DD01"
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"
				vgo_custom.FOL_cAddtoolbarapps ("Totalização das Fichas por Item", 'vgo_dat.FOL_donormalform([dd01_001])', PLO_Form.VOL_Editing )
		endcase

	case VLC_FORMNAME == "D45_001"
		do case
			*- Bloqueando alguns objetos na tela e tratando o numero do lote para nf de entrada e apontamento de mat prima
			case PLC_CodeId == "D45_001.ACTIVATE"
				do trg_d45_001_activate.FXP with PLO_Form
		endcase

	case VLC_FORMNAME == "D45_003"
		do case
			*- Libera alguns botões na tela 
			case PLC_CodeId == "D45_003.ACTIVATE"
				do trg_d45_003_activate.FXP with PLO_Form
		endcase

 
	case VLC_FormName == "E08"				&&-Pedido de Compra 
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"
				VGO_Custom.FOL_cAddtoolbarapps("Notas integrantes (EDI-DOCCOB)", "VGO_Custom.FOL_cDoForm('ee01_edinotas.scx')", !empty(E08.UKEY) )
				VGO_Custom.FOL_cAddtoolbarapps("Cancelamento de liberação (COMPRA)", "DO TGR_E08_TOOLBARAPPS_CANLIB", !empty(E08.UKEY) and E08.E08_600_N = 1 and !FLL_VerifAcesso("CANC_LIBCOM"))
				vgo_custom.FOL_cAddtoolbarapps("Transporte de Itens do Pedido de Compra do CloneImports", 'vgo_dat.FOL_donormalform("importa_clone",_screen.activeform)', PLO_Form.VOL_adding and !empty(e08.A08_ukey) and !isnull(e08.A08_ukey) and reccount("E09t") = 0)

*!*				case upper(PLC_CodeId) = "E08.E08_E08_003_D.REFRESH"
*!*					PLO_Form.e08_e08_003_d.date1.VOL_Enabled = .F.

			case upper(PLC_CodeId) = "E08.FOL_SAVE"
				do TGR_E08_SAVE

			case upper(PLC_CodeId) = "E08.INIT"
				do TGR_E08_INIT with PLO_Form

			case upper(PLC_CodeId) = "E08.FOL_EDIT"
				if E08.E08_600_N = 1
					messagebox("Não é permitido alterar/editar um pedido já liberado. Operação Cancelada!",16,"Atenção")
					VGO_GEN.VOL_TRIGGERRETURN = .f.
				endif

			case upper(PLC_CodeId) = "E08.FOL_DELETE"
				if E08.E08_600_N = 1
					messagebox("Não é permitido Excluir um pedido já liberado. Operação Cancelada!",16,"Atenção")
					PLO_Form.refresh()
					VGO_GEN.VOL_TRIGGERRETURN = .f.
				endif
				if VGO_GEN.VOL_TRIGGERRETURN
					do tgr_e08_Delete.fxp
				endif
			case upper(PLC_CodeId) = "E08.FOL_AFTERSAVE"
				do tgr_e08_aftersave.fxp

			case "E08.PAGEFRAME1.PAGE4.E08_A21_UKEY.TEXTALIAS1.INIT" $ upper(PLC_CodeId)  && 	 No Init da Tela
				PLO_Form.PAGEFRAME1.PAGE4.e08_a21_ukey.textALIAS1.vol_chkifempty=.t.

			case "E08.PAGEFRAME1.PAGE4.E08_A21_UKEY.TEXTALIAS1.LOSTFOCUS" $ upper(PLC_CodeId) && marco o tipo de documento em todas as parcelas
				do e05_TipoDoc with E08.A21_Ukey

			case upper("E08.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON1") $ upper(PLC_CodeId)
				PLO_Form.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON1.visible=.f.

			case upper("E08.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON2") $ upper(PLC_CodeId)
				PLO_Form.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON2.visible=.f.

			case upper(PLC_CodeId) = "E08.FOL_BEFORESAVE"
				if vgo_gen.VOL_TRIGGERRETURN
					vgo_gen.VOL_TRIGGERRETURN = E08_BeforeSave()
				endif
		endcase

	case VLC_FormName == "E08_004"					&&-Notas Fiscal entrada utilizando um pedido de compra
		do case
			*-Colocando a data de entrada da nf com a data corrente
			case PLC_CodeId = "E08_004.UNLOAD"
				if not empty(alltrim(nvl(e10.e08_ukey,'')))
					replace e10.e10_014_d with date()
				endif
			endcase

	case VLC_FormName == "E09"				&&-Itens do Pedido de Compra
		do case			
			case PLC_CodeId $ "E09.REFRESH"
				*- Desabilita o campo VALOR UNITÁRIO se for maior que Zeros
				do tgr_E09_init.fxp with PLO_Form
				
			case PLC_CodeId $ "E09.FOL_SAVE"
				do FLL_VerifCES.fxp with "E09T"
		endcase


	case VLC_FormName == "E10"				&&-Nota Fiscal de Entrada
		do case

			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				*VGO_Custom.FOL_cAddtoolbarapps("Cancelamento da Liberação", "VGO_Custom.FOL_cDoForm('JJ604',E10.UKEY)", NOT EMPTY(PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1.VALUE) )
				VGO_Custom.FOL_cAddtoolbarapps("Cancelamento de liberação (RMA)", "DO TGR_E10_TOOLBARAPPS_CANLIB", !empty(E10.UKEY) and E10.E10_632_N = 1 and !FLL_VerifAcesso("CANC_LIBCOMRMA"))

				*- Adiciona a opção de informação do Protocolo de intenção - CASOUZA - 23/08/07
				VGO_Custom.FOL_cAddtoolbarapps("PTA - Protocolo de Intenção", "VGO_Custom.FOL_cDoForm('e10cl_pta')", !eof("E10"))

			case "E10.FOL_EDIT" $ upper(PLC_CodeId) && Força a serie interna quando edito o documento.
				if E10.E10_632_N = 1
					do TGR_E10_LIBERADO with "Alterar"
				endif
				*				Replace E10_033_n WITH 1  && Solicitação da Marcia Fiscal, para não ser mais obrigatório a Série da NF na Entrada 10/07/2007 Mantis 0000489

			case "E10.FOL_BEFOREDELETE" $ upper(PLC_CodeId)
				if E10.E10_632_N = 1
					do TGR_E10_LIBERADO with "Apagar"
				endif

			case PLC_CodeId $ "E10.FOL_ADDNEW"
				*Torna o campo Numero da Nota Fiscal não obrigatório, mas depois volta no fol_beforesave
				PLO_Form.e10_e10_001_c.vol_chkifempty = .f.


			case PLC_CodeId $ "E10.FOL_BEFORESAVE"
				* Busca numeração de NF de Venda para Entrada de Importação
				do tgr_E10_beforesave.fxp with PLO_Form
				if vgo_gen.VOL_TRIGGERRETURN
					vgo_gen.VOL_TRIGGERRETURN = E10_BeforeSave()
				endif

				*!*CASE "E10.E10_E10_033_N.INIT" $ UPPER(PLC_CodeId) && Desabilita o campo Série interna
				*!*PLO_Form.E10_E10_033_N.VOL_ENABLED = .f.  && Solicitação da Marcia Fiscal, para não ser mais obrigatório a Série da NF na Entrada 10/07/2007 Mantis 0000489

			case "E10.PAGEFRAME1.PAGE4.E10_A21_UKEY.TEXTALIAS1.INIT" $ upper(PLC_CodeId) && Init do Tipo de Documento (Aba Financeiro)
				PLO_Form.PAGEFRAME1.PAGE4.e10_a21_ukey.textALIAS1.vol_chkifempty=.t.

			case "E10.PAGEFRAME1.PAGE1.E10_T88_UKEY.TEXTALIAS1.INIT" $ upper(PLC_CodeId) && Init do Tipo de Documento (Aba Financeiro)
				PLO_Form.PAGEFRAME1.PAGE1.e10_T88_ukey.textALIAS1.vol_chkifempty=.t.

			*case upper("fol_replacefields") $ upper(PLC_CodeId)  && Marca a serie interna automaticamente na adição de um novo registro
				*	Replace E10_033_n WITH 1 && Solicitação da Marcia Fiscal, para não ser mais obrigatório a Série da NF na Entrada 10/07/2007 Mantis 0000489

			case "E10.PAGEFRAME1.PAGE4.E10_A21_UKEY.TEXTALIAS1.LOSTFOCUS" $ upper(PLC_CodeId) && marco o tipo de documento em todas as parcelas
				do e05_TipoDoc with E10.A21_Ukey

			case upper("E10.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON1") $ upper(PLC_CodeId)
				PLO_Form.PAGEFRAME1.PAge3.Cgo_showgridcontainer1.Grid1.Column1.COMMANDBUTTON1.visible=.f.

			case PLC_CodeId = "E10.INIT"
				PLO_Form.CGO_A03A08SELECTION1.A08_TEXT.TEXTALIAS1.tabindex=4
				PLO_Form.CGO_A03A08SELECTION1.A08_TEXT.TEXTALIAS1.vol_chkifempty=.t.

				do tgr_e10_init with PLO_Form
				 
				*Colocando a trava de quem pode inserir nf em determinada empresa via trigger porque o acesso é por empresa assim tendo que tirar totalmente o acesso.
				*do TGR_E10_VerificaAcesso with PLO_Form
				
				*-- Ajuste da sequencia dos campos da nota
			case PLC_CodeId == "E10.E10_E10_007_C.LOSTFOCUS"
				PLO_Form.E10_E10_003_D.setfocus()

			case PLC_CodeId == "E10.E10_E10_033_N.LOSTFOCUS"
				PLO_Form.E10_J09_UKEY.TEXTALIAS1.setfocus()

			case PLC_CodeId == "E10.E10_J09_UKEY.TEXTALIAS1.LOSTFOCUS"
				PLO_Form.E10_E10_003_D.setfocus()

			case PLC_CodeId == "E10.E10_E10_003_D.DATE1.LOSTFOCUS"
				PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.setfocus()

			case PLC_CodeId == "E10.CGO_A03A08SELECTION1.ARRAY_109.LOSTFOCUS"
				PLO_Form.CGO_A03A08SELECTION1.A08_TEXT.TEXTALIAS1.setfocus()

			case PLC_CodeId == "E10.CGO_A03A08SELECTION1.A08_TEXT.TEXTALIAS1.LOSTFOCUS"
				PLO_Form.PAGEFRAME1.PAGE1.E10_E10_014_D.DATE1.setfocus()

			case PLC_CodeId = "E10.REFRESH"  && Filipe Oliveira - pesquisa clientes ativos e em estudo. 03/09/2008
				with PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith

				if !PLO_Form.VOC_MyExecuteCursor == E10.Ukey or VGO_Gen.VOA_TriggerParameters[1]
					PLO_Form.VOC_MyExecuteCursor = E10.Ukey

					*- Verifica se a nota de entrada foi referenciada em algum pedido de RMA
					VGO_Gen.FOL_SqlExec("SELECT UKEY, J07_001_C FROM STAR_DATA@J07 (NOLOCK) WHERE J07.E10CL_UKEY = ?E10.UKEY", "TmpJ07")

					*- Caso tenha encontrado valoriza o textbox
					if reccount("TmpJ07") > 0
						PLO_Form.pageframe1.page1.e10_j07_001_c.value = TmpJ07.J07_001_C
						PLO_Form.pageframe1.page1.e10_j07_001_c.refresh()
					else
						PLO_Form.pageframe1.page1.e10_j07_001_c.value = ""
						PLO_Form.pageframe1.page1.e10_j07_001_c.refresh()
					endif

					*- Verifica se o pedido de RMA foi faturado
					VGO_Gen.FOL_SqlExec("SELECT UKEY, J10_001_C FROM STAR_DATA@J10 (NOLOCK) WHERE J10.J07_UKEY = ?TmpJ07.UKEY", "TmpJ10")

					*- Caso tenha encontrado valoriza o textbox
					if reccount("TmpJ10") > 0
						PLO_Form.pageframe1.page1.e10_j10_001_c.value = TmpJ10.J10_001_C
						PLO_Form.pageframe1.page1.e10_j10_001_c.refresh()
					else
						PLO_Form.pageframe1.page1.e10_j10_001_c.value = ""
						PLO_Form.pageframe1.page1.e10_j10_001_c.refresh()
					endif

					VGO_Gen.FOL_CloseTable("TmpJ07")
					VGO_Gen.FOL_CloseTable("TmpJ10")
				endif


			case PLC_CodeId = "E10.PAGEFRAME1.PAGE1.E10_T88_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Especie da NF
				with PLO_Form.pageframe1.page1
					.e10_t88_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith

			case PLC_CodeId = "E10.PAGEFRAME1.PAGE1.E10_T89_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Modelo da NF
				with PLO_Form.pageframe1.page1
					.e10_t89_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith
		endcase

	case VLC_FormName == "E11"	
		do case	
			case PLC_CodeId $ "E11.FOL_SAVE"
				do FLL_VerifCES.fxp with "E11T"
				
				* Verificando se a nota pode ser lançada sem que seja por um pedido de compra
				do tgr_e11_save.fxp with PLO_Form				
		
		case PLC_CodeId $ "E11.REFRESH"		
				* Verificando se o campo e valor unitario pode ficar habilitado
				do tgr_e11_refresh.fxp with PLO_Form	
				 		
		case PLO_Form.VOL_EDITING
				do tgr_e11_edit.fxp with PLO_Form
			
		endcase
 
	case VLC_FormName == "E10_001"				&&-Devolução de Entrada

		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Dados da Logistica - Dev. de Entrada", "VGO_Custom.FOL_cDoForm('JJ610',E10.UKEY)", not empty(PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1.value) )

			case PLC_CodeId $ "E10_001.FOL_ADDNEW"
				if PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.value = 1
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .t.
				else
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .f.
				endif
				PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.refresh()


			case PLC_CodeId $ "E10_001.INIT"
				* Inclui campos na tela de Devolução de Entrada
				do tgr_E10_001_init.fxp with PLO_Form

			case PLC_CodeId = "E10_001.FOL_AFTERSAVE"
				* Calcula a M3 da NF de devolução de entrada
				do tgr_E10_001_aftersave_m3.fxp with PLO_Form


			case PLC_CodeId $ "E10_001.FOL_EDIT"
				if PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.value = 1
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .t.
				else
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .f.
				endif

			case PLC_CodeId $ "E10_001.CGO_A03A08SELECTION1.ARRAY_109.CLICK"
				if PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.value = 1
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .t.
				else
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.VOL_ENABLED = .f.
					replace E10_UKEY1 with "" in E10
				endif
				PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.refresh()

			case PLC_CodeId $ "E10_001.PAGEFRAME1.PAGE1.E10_E10_UKEY1.TEXTALIAS1.CLICK"
				if PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.value = 1
					if empty(nvl(PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1.value,''))
						messagebox("Informe Código do Cliente. Operação Cancelada!",48,"Atenção")
						PLO_Form.CGO_A03A08SELECTION1.A03_TEXT.TEXTALIAS1.setfocus()
						VGO_GEN.VOL_TRIGGERRETURN = .f.
						return .f.
					endif
				endif

				if E10.E10_PAR = 'A03'
					PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.TEXTALIAS1.VOA_FILTER[1] = E10.E10_UKEYP
				endif

			case PLC_CodeId $ "E10_001.FOL_BEFORESAVE"
				* Deixa o campo data de saida da NF em branco
				do tgr_E10_001_beforesave.fxp with PLO_Form
				if PLO_Form.CGO_A03A08SELECTION1.ARRAY_109.value < 2 and alltrim(TMP_TipoNota.t04_001_c) <> "S0165"
					if empty(nvl(E10.E10_UKEY1,''))
						messagebox("Informe a nota de origem desta devolução.",48,"Atenção")
						PLO_Form.PAGEFRAME1.PAGE1.setfocus()
						PLO_Form.PAGEFRAME1.PAGE1.E10_E10_UKEY1.setfocus()
						VGO_GEN.VOL_TRIGGERRETURN = .f.
						return .f.
					endif
				else
					replace E10_UKEY1 with "" in E10
				endif

				if VGO_GEN.VOL_TRIGGERRETURN = .t.
					do tgr_E10_001_beforesave.fxp with PLO_Form
				endif

			case PLC_CodeId == "E10_001.REFRESH"  && Filipe Oliveira 08/09/2008
				with PLO_Form.cGO_A03A08SELECTION1.A03_TEXT.tEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith

			case PLC_CodeId = "E10_001.PAGEFRAME1.PAGE1.E10_T88_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Especie da NF
				with PLO_Form.pageframe1.page1
					.e10_t88_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith

			case PLC_CodeId = "E10_001.PAGEFRAME1.PAGE1.E10_T89_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Modelo da NF
				with PLO_Form.pageframe1.page1
					.e10_t89_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith
		endcase
	
	case VLC_FormName == "E11_001"	
		do case	
			case PLC_CodeId $ "E11_001.FOL_SAVE"
				do FLL_VerifCES.fxp with "E11T"
		endcase

	case VLC_FormName == "E11_002"				&&-Importação dos Itens da Devolução de Entrada

		do case
			case PLC_CodeId $ "E11_002.FOL_CLOSE"
				* Verifica se for S0010 ou S0080 força local Não Conforme.
				do tgr_E11_002_FOL_CLOSE.fxp with PLO_Form

		endcase


	case VLC_FORMNAME == "F09_002"
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Montagem de Lotes", "VGO_Custom.FOL_cDoForm('F09_MONTALOTE')",PLO_Form.VOL_EDITING )
		endcase

	case VLC_FormName == "F11_001"              &&-Contas a Pagar
		do case
			case PLC_CodeId $ "F11_001.FOL_ADDNEW"
				*- Passa valor 0 para o array_603
				PLO_Form.PAGEFRAME1.PAGE1.F11_ARRAY_603.value = 0
			case PLC_CodeId = "F11_001.PAGEFRAME1.PAGE1.F11_ARRAY_478.LOSTFOCUS" and PLO_Form.PAGEFRAME1.PAGE1.F11_ARRAY_478.value <> 9
				*- Se Tipo de Docto <> Impostos gravar Tipo de Imposto = 0  combo "Em Branco"
				PLO_Form.PAGEFRAME1.PAGE1.F11_ARRAY_603.value = 0
			case PLC_CodeId = "F11_001.PAGEFRAME1.PAGE2.F11_A21_UKEY.TEXTALIAS1.INIT"
				*- Forço obrigatoriedade do campo de tipo de documento
				PLO_Form.PAGEFRAME1.PAGE2.F11_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.
			case PLC_CodeId = "F11_001.PAGEFRAME1.PAGE2.F11_A21_UKEY.TEXTALIAS1.LOSTFOCUS"
				*- Repasso tipo de documento para as parcelas
				do F11_TipoDoc with F11.a21_UKEY
			case PLC_CodeId $ "F11_001.FOL_BEFORESAVE"
				if vgo_gen.VOL_TRIGGERRETURN
					vgo_gen.VOL_TRIGGERRETURN = F11_001_BeforeSave()
				ENDIF
			case PLC_CodeId == "F11_001.INIT"
				*do tgr_f11_001_init.fxp with PLO_Form  && Inserindo opção de não provisionar na contabilidade
			
		endcase

	case VLC_FormName == "F11_002"              &&-COMISSÕES A PAGAR
		do case
			case "F11_A21_UKEY.TEXTALIAS1.INIT" $ PLC_CodeId
				*- Forço obrigatoriedade do campo de tipo de documento
				PLO_Form.PAGEFRAME1.PAGE2.F11_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.
			
			case "F11_A21_UKEY.TEXTALIAS1.LOSTFOCUS" $ PLC_CodeId
				*- Repasso tipo de documento para as parcelas
				do F11_TipoDoc with F11.a21_UKEY
		endcase

	case VLC_FormName == "F11_003"              &&-ADIANTAMENTO DE VALORES A PAGAR
		do case
			case "F11_A21_UKEY.TEXTALIAS1.INIT" $ PLC_CodeId
				*- Forço obrigatoriedade do campo de tipo de documento
				PLO_Form.PAGEFRAME1.PAGE2.F11_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.

			case "F11_A21_UKEY.TEXTALIAS1.LOSTFOCUS" $ PLC_CodeId
				*- Repasso tipo de documento para as parcelas
				do F11_TipoDoc with F11.a21_UKEY
		endcase

	case VLC_FormName == "F12_001"              &&-Filipe Oliveira
		do case
			case PLC_CodeId = "F12_001.REFRESH"
				*- Buscanco informações de dt de coleta e dt de entrega da nf origem do titulo .
				text to vlc_dadosentrega noshow
					select distinct jj600.jj600_004_d as dt_coleta, j10.j10_006_D as dt_entrega from emp04.J10 j10 (nolock)
					inner join emp04.JJ600 jj600 (nolock) on j10.JJ600_UKEY = jj600.UKEY where j10.ukey = ?f12.f12_iukeyp
				endtext
				Vgo_Dat.FOL_SqlExec(vlc_dadosentrega,"TMP_DadosEntrega")
				
			case PLC_CodeId = "F12_001.ACTIVATE"
				with PLO_Form.paGEFRAME1.pAGE1.A03A08SELECTION1.A03_TEXT.TEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
				
			*- Procedimento para trocar o tipo de documento sem refazer as parcelas
			case PLC_CODEID = "F12_001.INIT"
				PLO_Form.addproperty("VOC_AntigaA21Ukey")
				PLO_Form.addproperty("VOC_NovaA21Ukey")
				PLO_Form.VOC_AntigaA21Ukey = ""
				PLO_Form.VOC_NovaA21Ukey = ""
			
				*- inserindo alguns campos na tela de titulo a receber
				do tgr_f12_001_init.fxp with PLO_Form				
			
			case PLC_CODEID = "F12_001.FOL_EDIT"
				PLO_Form.VOC_AntigaA21Ukey = f12.a21_ukey
			
			case PLC_CODEID = "F12_001.FOL_BEFORESAVE"
				PLO_Form.VOC_NovaA21Ukey = f12.a21_ukey
				replace f12.a21_ukey with PLO_Form.VOC_AntigaA21Ukey in f12
				
			case PLC_CODEID = "F12_001.FOL_SAVE"
				replace f12.a21_ukey with PLO_Form.VOC_NovaA21Ukey in f12
			*-----------------------------------------------------------------------
		endcase

	case VLC_FormName == "F12_003"              &&-Filipe Oliveira
		do case
			case PLC_CodeId = "F12_003.ACTIVATE"
				with PLO_Form.PAGEFRAME1.PAGE1.Selection1.A03_TEXT.TEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
		endcase

	case VLC_FormName == "F14_001"              &&-PARCELAS CONTAS A PAGAR
		do case
			case PLC_CodeId = "F14_001.F14_A21_UKEY.TEXTALIAS1.INIT"
				PLO_Form.F14_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.
		endcase

	case VLC_FormName == "F14_002"              &&-PARCELAS COMISSÕES A PAGAR
		do case
			case PLC_CodeId = "F14_002.F14T_A21_UKEY.TEXTALIAS1.INIT"
				PLO_Form.F14T_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.
				
				*-Inserindo o objeto imagem para dizer se a parcela foi estornada parcial, estornada total ou não estornada
				PLO_Form.addobject("IMAGEN","cgo_STAMP")
					with PLO_Form.IMAGEN
						.height = 84
						.left = 240
						.top = 36
						.width = 260
						.visible = .f.						
					endwith				
		
		case PLC_CodeId == "F14_002.REFRESH"
				do tgr_f14_002_refresh.fxp with PLO_Form			
		endcase

	case VLC_FormName == "F14_003"              &&-PARCELAS ADIANTAMENTO DE VALORES A PAGAR
		do case
			case PLC_CodeId = "F14_003.F14_A21_UKEY.TEXTALIAS1.INIT"
				PLO_Form.F14_A21_UKEY.TEXTALIAS1.vol_chkifempty = .t.
		endcase

	case VLC_FormName == "F15_002"				&&-Parcelas de Quitações a Receber
		do case
			case PLC_CodeId = "F15_002.INIT"
				do tgr_f15_002_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "F15_031"				&&-Parcelas de Estorno de Titulos a Receber
		do case
			case PLC_CodeId = "F15_031.INIT"
				do tgr_f15_031_init.fxp with PLO_Form
		endcase
 
	case VLC_FORMNAME == "F24"
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Copia CNAB", "VGO_Custom.FOL_cDoForm('F24_COPIACNAB')",.t. )
		endcase

	case VLC_FormName == "F28" 				&&-Ficha financeira
		do case
			case PLC_CodeId == "F28.INIT" 	&& Inclui botão para calcular o crédito consolidando o SSA com o Siscorp
				do tgr_f28_init.fxp with PLO_Form

			case PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Histórico de acompanhamento", "VGO_Custom.FOL_cDoForm('F28_historico',_screen.ActiveForm)" , not empty(PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.value) or not VGO_Dat.FOL_AcessMenu("F28_HISTORICO"))
				VGO_Custom.FOL_cAddtoolbarapps("Credit score - crédito sugerido", "VGO_Custom.FOL_cDoForm('CreditScore',_screen.ActiveForm)", not empty(PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.value) or not VGO_Dat.FOL_AcessMenu("CREDITSCORE"))

			case "F28.F28_F28_001_D.DATE1.GOTFOCUS"  $  PLC_CodeId
				*- Quando sair do campo código do cliente.
				do tgr_f28_001_D_GOTFOCUS.fxp with PLO_Form

			case "GOTFOCUS" $ PLC_CodeId   				&& Controla alteração do limite de crédito.
				do tgr_f28_update.fxp with PLO_Form

			case PLC_CodeId == "FORMSET1.F28.GRID1.ANALITICO.COMMANDBUTTON1.CLICK"
				do tgr_f28_analitico.fxp with PLO_Form, PLO_Form.parent
		endcase

	case VLC_FormName == "F17_007"
		do case
			case PLC_CodeId == "F17_007.FOL_BEFORESAVE"
				do Tgr_F17_007_BeforeSave.prg  with PLO_Form 
		endcase

	case VLC_FormName == "F18_002"
		do case
			case PLC_CodeId == "F18_002.FOL_AFTERSAVE"
				do tgr_f18_002_serasa.fxp

			case PLC_CodeId == "F18_002.FOL_BEFOREDELETE"
				do tgr_f18_002_serasa.fxp
		endcase


	case VLC_FORMNAME == "F18_003"
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Movimentação de Conta Corrente COMISSÕES", "DO TGR_F22_MOVIMENTO",PLO_Form.VOL_EDITING and !FLL_VerifAcesso("MOV_CC"))
		endcase

	case VLC_FormName == "F18_003_1"
		do case
			case PLC_CodeId = "F18_003_1.DOIT.CLICK"
				do trg_f18_003_1_Doit_click.fxp with PLO_Form
				
			case PLC_CodeId = "F18_003_1.RW.CLICK"
				do trg_f18_003_1_Rw_click.fxp 			
		endcase

	case VLC_FormName == "F18_023" 						&&-CNAB
		do case
			case PLC_CodeId == "F18_023.FOL_SAVE"
				if VGO_Custom.VOL_TriggerReturn = .t. &&-Se todas as operações do botão SAVE for concluida.
					*- Grava os dados bancarios na parcela do titulo e marca que o arquivo tem CNAB
					do tgr_f18_023_save.fxp with PLO_Form
				endif

			case PLC_CodeId == "F18_023.FOL_DELETE"
				if VGO_Custom.VOL_TriggerReturn = .t.
					*- Remove os dados bancarios da parcela do titulo no financeiro
					do tgr_f18_023_delete.fxp with PLO_Form
				endif
		endcase

	case VLC_FormName == "F18_025"
		do case
			case PLC_CodeId = "F18_025.DOIT.CLICK"                       && Adalberto A. dos Santos
				do tgr_f18_025_doit.fxp with PLO_Form
		endcase

	case VLC_FormName == "F18_031" && - Colocando o tipo nota da parcela estornada na parcela do titulo a receber
		do case
			case PLC_CodeId = "F18_031.FOL_AFTERSAVE"                      
				do tgr_f18_031_fol_aftersave.fxp with PLO_Form
		endcase
		
	case VLC_FormName == "F40"              &&-tipo de operação conta corrente
		do case
			case PLC_CodeId == "F40.INIT"
				do tgr_f40_init.fxp with PLO_Form  && liberando a opção de vinculo de conta contabil
		endcase

	case VLC_FormName == "I18" 						&&-Problemas do Laudo Técnico
		do case
			case PLC_CodeId == "I18.INIT"
				do tgr_i18_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "H03_001"
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"
				* 		     VGO_Custom.FOL_cAddtoolbarapps ("Correção dos Livros Fiscais (J10-H03)", "VGO_Custom.FOL_cDoForm ('J10_CORRIGELIVRO.SCX') ", !PLO_Form.VOL_EDITING)
				VGO_Custom.FOL_cAddtoolbarapps ("Fechamento da Área Livros Fiscais (PAR-00156)", "VGO_Custom.FOL_cDoForm ('ALTERDT_PAR00156.SCX') ", !PLO_Form.VOL_EDITING)
		endcase

	case VLC_FormName == "H03_002"
		do case
			case PLC_CodeId == "LOADTOOLBARAPPS"
				* 		     VGO_Custom.FOL_cAddtoolbarapps ("Correção dos Livros Fiscais (J10-H03)", "VGO_Custom.FOL_cDoForm ('J10_CORRIGELIVRO.SCX') ", !PLO_Form.VOL_EDITING)
				VGO_Custom.FOL_cAddtoolbarapps ("Fechamento da Área Livros Fiscais (PAR-00156)", "VGO_Custom.FOL_cDoForm ('ALTERDT_PAR00156.SCX') ", !PLO_Form.VOL_EDITING)
		endcase

	case VLC_FormName == "J01"						&&-Orçamento
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				if PLO_Form.VOL_EDITING=.f. and not empty(PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.value)
					VGO_Custom.FOL_cAddtoolbarapps ("Pedidos que utilizaram este orçamento", "VGO_Custom.FOL_cDoForm ('J01_PEDIDO.scx') ", !PLO_Form.VOL_EDITING)
				else
					VGO_Custom.FOL_cAddtoolbarapps ("Pedidos que utilizaram este orçamento", "VGO_Custom.FOL_cDoForm ('J01_PEDIDO.scx') ", .f.)
				endif

			case PLC_CodeId = "J01.PAGEFRAME1.PAGE4.J01_A13_UKEY.TEXTALIAS1.FOL_UPDATE" && Comissao
				do Tgr_J01_A13_Update1.Fxp with PLO_Form
				with PLO_Form.Pageframe1.Page1
					.J01_D01_Ukey.refresh()
					.J01_D01_Ukey.TextAlias1.refresh()
					.J01_D02_Ukey.refresh()
					.J01_D02_Ukey.TextAlias1.refresh()
				endwith

			case PLC_CodeId == "J01.PAGEFRAME1.PAGE1.J01_D01_UKEY.REFRESH" && Comissao
				PLO_Form.PAGEFRAME1.PAGE1.J01_D01_UKEY.VOL_ENABLED = .f.

			case PLC_CodeId == "J01.PAGEFRAME1.PAGE1.J01_D02_UKEY.REFRESH" && Comissao
				PLO_Form.PAGEFRAME1.PAGE1.J01_D02_UKEY.VOL_ENABLED = .f.

			case PLC_CodeId = upper("J01.pageFRAME1.page2.container1.cmd_new.Click") && Comissao1
				if empty(J01.A13_Ukey) or isnull(J01.A13_Ukey)
					=messagebox("Favor Preencher antes a condicao de pagamento.",16,"Atenção!")
					vgo_gen.vol_trIGGERRETURN = .f.
				endif

			case PLC_CodeId == "J01.A03A08SELECTION1.LOSTFOCUS" && Comissao
				Vgo_Dat.FOL_SqlExec("Select A13_Ukey From STAR_DATA@A03 (NOLOCK) Where A03.Ukey = ?J01.A03_Ukey","A03_A13")
				if used("A03_A13")
					select J01
					replace A13_Ukey with A03_A13.A13_Ukey
					Vgo_Dat.FOL_Closetable("A03_A13")
				endif
				do Tgr_J01_A13_Update.Fxp
				with PLO_Form.Pageframe1.Page1
					.J01_D01_Ukey.refresh()
					.J01_D01_Ukey.TextAlias1.refresh()
					.J01_D02_Ukey.refresh()
					.J01_D02_Ukey.TextAlias1.refresh()
				endwith
				PLO_Form.Pageframe1.Page4.J01_A13_Ukey.refresh()
				PLO_Form.Pageframe1.Page4.J01_A13_Ukey.TextAlias1.refresh()

			case PLC_CodeId = "J01.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.FOL_UPDATE"
				*- Emite mensagem se o cliente estiver bloqueado/inativo ou em estudo, mas permite fazer o orçamento
				do tgr_j01_bloq.fxp with PLO_Form

			case PLC_CodeId = "J01.INIT"
				do tgr_J01_init.fxp with PLO_Form

			case PLC_CodeId = "J01.REFRESH"
				with PLO_Form.A03A08SELECTION1.A03_TEXT.tEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
				
			case PLO_Form.VOL_EDITING
				*- Desabilita alguns campos e altera o lay-out de outros
				do tgr_J01_edit.fxp with PLO_Form

		endcase


	case VLC_FormName == "J02"				 		&&-Itens do Orcamento
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				VGO_Custom.FOL_cAddtoolbarapps("Configuração de E/S utilizada", "VGO_Custom.FOL_cDoForm('J02J08J11_T06')", .T.)
			
			case PLC_CodeId $ "J02.FOL_BEFORESAVE"
				if empty(J02t.j02_007_c)
					PLO_Form.CGO_RECALCULAST.FOL_EXECUTE() && executa a rotina para recalculo da ST
				endif

			case PLC_CodeId $ "J02.FOL_CLOSE" && Comissao
				if !PLO_Form.Vol_Noedit
					do Tgr_J02_Close.fxp
				endif

			case "J02.ACTIVATE" $ upper(PLC_CodeId)
				do J02_Recalculo.fxp with PLO_Form
				
				*- Insere o campo de Qrde de Embalagem abaixo da Qtde Disponivel para o item
				if vartype(PLO_Form.VOO_FormSec) == "O"
					*- Cria o objeto responsavel por atualizar a Qtde de Embalagem
					if vartype(PLO_Form.VOO_FormSec.OOO_Package) <> "O"
						PLO_Form.VOO_FormSec.newobject("OOO_Package", "CGO_Package", "SGO_Package.FXP")
					
						*- Atualiza a Qtde de Embalagem quando rodar o refresh da Qtde disponivel
						bindevent(PLO_Form.VOO_FormSec.stockbalanceavailable1.spinner1, "REFRESH", PLO_Form.VOO_FormSec.OOO_Package, "FOL_ShowPackage")
					endif
				endif

			case PLC_CodeId = "J02.J02_D01_UKEY.REFRESH" && Comissao
				PLO_Form.J02_D01_UKEY.Vol_Enabled = .f.

			case PLC_CodeId = "J02.J02_D02_UKEY.REFRESH" && Comissao
				PLO_Form.J02_D02_UKEY.Vol_Enabled = .f.

			case PLC_CodeId = "J02.J02_T02_UKEY0.TEXTALIAS1.GOTFOCUS"  &&Unidade de Medida
				*- Desabilita a Lista de Preços para não permitir alteração
				PLO_Form.J02_D01_UKEY.TEXTALIAS1.enabled = .f.
				PLO_Form.J02_D02_UKEY.TEXTALIAS1.enabled = .f.
				*- Deixa o campo valor total desabilitado
				PLO_Form.J02_J02_006_B.SPINNER1.enabled = .f.
				*- Deixa o campo Natureza de Operação desabilitado
				PLO_Form.J02_A28_UKEY.TEXTALIAS1.enabled = .f.
				*- Na alteração ou inserção da qtde do desconto mostra o valor liquido
				*do tgr_J02_QTDE.fxp with PLO_Form

			case PLC_CodeId = "J02.J02_T04_UKEY.TEXTALIAS1.GOTFOCUS"  &&Tipo de E/S
				*- Desabilita a Lista de Preços para não permitir alteração
				PLO_Form.J02_D01_UKEY.TEXTALIAS1.enabled = .f.
				PLO_Form.J02_D02_UKEY.TEXTALIAS1.enabled = .f.
				*- Deixa o campo valor total desabilitado
				PLO_Form.J02_J02_006_B.SPINNER1.enabled = .f.
				*- Deixa o campo Natureza de Operação desabilitado
				PLO_Form.J02_A28_UKEY.TEXTALIAS1.enabled = .f.

			case PLC_CodeId = "J02.J02_T04_UKEY.TEXTALIAS1.FOL_UPDATE"  &&Tipo de E/S
				*- Verifica se o Tipo de E/S está em Branco
				do tgr_J02_t04_ukey_fol_update.fxp with PLO_Form

			case PLC_CodeId $ "J02.J02_J02_005_B.SPINNER1.FOL_UPDATE"
				VLC_Percentual = PLO_Form.j02_j02_604_c.FOC_EncontraPercDesc(j02t.j02_005_b, j02t.j02_606_b)
			
				replace j02t.j02_604_c with VLC_Percentual in j02t
				replace j02t.j02_605_b with iif(empty(VLC_Percentual), 0, (j02t.j02_606_b - j02t.j02_005_b) * j02t.j02_003_b) in j02t
				PLO_Form.j02_j02_605_b.refresh()	

			case PLC_CodeId $ "J02.J02_J02_003_B.SPINNER1.GOTFOCUS"   && Quantidade
				*- Desabilita a Lista de Preços para não permitir alteração
				PLO_Form.J02_D01_UKEY.TEXTALIAS1.enabled = .f.
				PLO_Form.J02_D02_UKEY.TEXTALIAS1.enabled = .f.
				*- Deixa o campo valor total desabilitado
				PLO_Form.J02_J02_006_B.SPINNER1.enabled = .f.
				*- Deixa o campo Natureza de Operação desabilitado
				PLO_Form.J02_A28_UKEY.TEXTALIAS1.enabled = .f.
				
				if !(j02t.d04_ukey == PLO_Form.j02_j02_604_c.VOC_D04Ukey) AND PLO_Form.VOL_Adding
					replace j02_606_b with j02_005_b in j02t
					PLO_Form.j02_j02_606_b.refresh()	
					PLO_Form.j02_j02_604_c.VOC_D04Ukey = j02t.d04_ukey
				endif

			case PLC_CodeId $ "J02.J02_J02_003_B.SPINNER.LOSTFOCUS"   && Quantidade
				*- Digita 2 enter quando troca a quantidade
				if (PLO_Form.j02_j02_003_b.spinner1.value <> PLO_Form.j02_j02_003_b.spinner1.von_lastvalue ) and ;
						(PLO_Form.j02_j02_003_b.spinner1.von_lastvalue <> 0) and !empty(PLO_Form.j02_j02_007_c.value)
					keyboard 'S'
					keyboard 'S'
				endif

			case PLC_CodeId $ "J02.J02_J02_003_B.SPINNER1.FOL_UPDATE" && Quantidade
				*- Essa trigger procura o local que tenha estoque para o produto
				do tgr_J02_local.FXP with PLO_Form
				*- Essa trigger tem diversas funcionalidades
				do tgr_J02_Fol_Update.FXP with PLO_Form
				*- Na alteração ou inserção da qtde do desconto mostra o valor liquido
				*do tgr_J02_QTDE.fxp with PLO_Form
				PLO_Form.j02_j02_005_b.setfocus()

			case PLC_CodeId $ "J02.J02_J02_007_C.GOTFOCUS"  && Percentual do desconto
				*- Desabilita a Lista de Preços para não permitir alteração
				PLO_Form.J02_D01_UKEY.TEXTALIAS1.enabled = .f.
				PLO_Form.J02_D02_UKEY.TEXTALIAS1.enabled = .f.
				*- Deixa o campo valor total desabilitado
				PLO_Form.J02_J02_006_B.SPINNER1.enabled = .f.
				*- Deixa o campo Natureza de Operação desabilitado
				PLO_Form.J02_A28_UKEY.TEXTALIAS1.enabled = .f.
				
				if empty(J02t.j02_007_c)
					PLO_Form.CGO_RECALCULAST.FOL_EXECUTE() && executa a rotina para recalculo da ST
				endif

			*case PLC_CodeId $ "J02.J02_J02_007_C.LOSTFOCUS" && Percentual do desconto
				*- Deixa o campo quantidade como inteiro ou decimal dependendo do tipo de Saída
				*do tgr_j02_j02_007_c_lostfocus.fxp with PLO_Form

			case PLC_CodeId $ "J02.FOL_ADDNEW" or PLC_CodeId $ "J02.J02_J02_007_C.FOL_UPDATE" && Percentual do desconto
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j02_j02_604_c.VOC_D04Ukey = ""
								
				*- adicionar e no momento da inserção/alteração do desconto/qtde mostra o valor liquido
				*- Tb roda um programa pra verificar a localização se está correta conforme o tipo da nota
				*do tgr_J02_QTDE.FXP with PLO_Form

			case PLC_CodeId = "J02.INIT"
				*-Deixa o campo quantidade como inteiro e valor do item com duas casas decimais
				do tgr_J02_init.fxp with PLO_Form && cria tambem objeto para recalculo ST

			case PLC_CodeId == "J02.FOL_EDIT" or PLC_CodeId == "J02.FOL_MAKESEARCH"
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j02_j02_604_c.VOC_D04Ukey = ""
				
				*- Atualiza o campo valor liquido
				*do tgr_J02_edit.fxp with PLO_Form

			case PLC_CodeId = "J02.FOL_SAVE"
				do FLL_VerifCES.fxp with "J02T"
				
				*- Verifica duplicidade de itens no pedido
				do tgr_j02_003_b_itemduplicado.fxp with PLO_Form
				if VGO_Custom.VOL_TriggerReturn
					* Corrige desconto e outras funcionalidades
					do tgr_j02_save.fxp with PLO_Form
				endif

*!*				case PLC_CodeId = "J02.FOL_UPDATE"
*!*					*- Deixa o campo quantidade como inteiro ou decimal dependendo do tipo de Saída
*!*					do tgr_J02_QTDE.fxp with PLO_Form

			case PLO_Form.VOL_EDITING
				*- Desabilita alguns campos e altera o lay-out de outros
				do tgr_J02_edit.fxp with PLO_Form

*!*				case not ("LOAD" $ PLC_CodeId) and not ("INIT" $ PLC_CodeId) and not ("EXECUTECURSOR" $ PLC_CodeId) && Comissao
*!*					*- Atualiza o valor do campo Valor Unitário Líquido
*!*					PLO_Form.liq.value = round((PLO_Form.J02_J02_008_B.value/PLO_Form.J02_J02_003_b.value),2)
		
		endcase

	case VLC_FormName == "J02_001" &&- Transporte de Orçamento para Pedido de venda
		do case
			case  PLC_CodeId == "J02_001.UNLOAD"
				*- Transporte do desconto especifico
				do tgr_j02_001_unload.fxp
		endcase


	case VLC_FormName == "J07"						&&-Pedido de venda
		do case
			*- Acionando a Barra de Ferramentas APPS
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				do tgr_j07_barra.fxp with PLO_Form

			case PLC_CodeId == "J07.FOL_ADDNEW"
				*- Deixo as variaveis abaixo vazia e falsa para o novo pedido de venda
				PLO_Form.VOC_PercDescontogeral = "" 
				PLO_Form.VON_FezCalcDescGeral = .f.
				
			case PLC_CodeId = "J07.FOL_CANCEL"
				PLO_Form.VOC_MyExecuteCursor = "Reprocessar" && Forco a execucao do executecursor especifico. O vgo_gen.voa_TriggerParameters[1] nao esta funcionando

			case PLC_CodeId = "J07.FOL_AFTERSAVE"
				PLO_Form.VOC_MyExecuteCursor = "Reprocessar" && Forco a execucao do executecursor especifico O vgo_gen.voa_TriggerParameters[1] nao esta funcionando
				do Tgr_J07_FOL_AfterSave.Fxp with PLO_Form
						
			case PLC_CodeId = "J07.PAGEFRAME1.PAGE6.J07_A13_UKEY.TEXTALIAS1.FOL_UPDATE" && Comissao
				do Tgr_J07_A13_Update1.Fxp with PLO_Form

			case PLC_CodeId == "J07.REFRESH"  && Filipe Oliveira
				with PLO_Form.A03A08SELECTION1.A03_TEXT.tEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
				
				do Tgr_J07_Refresh.Fxp with PLO_Form

			case PLC_CodeId == upper("J07.pageFRAME1.page4.j07_J07_610_N.Click")  && Filipe Tipo Bonificação
				do Tgr_J07_610_Click.Fxp with PLO_Form

			case PLC_CodeId == "J07.PAGEFRAME1.PAGE1.J07_D01_UKEY.REFRESH" && Comissao
				PLO_Form.PAGEFRAME1.PAGE1.J07_D01_UKEY.VOL_ENABLED = .f.

			case PLC_CodeId == "J07.PAGEFRAME1.PAGE1.J07_D02_UKEY.REFRESH" && Comissao
				PLO_Form.PAGEFRAME1.PAGE1.J07_D02_UKEY.VOL_ENABLED = .f.

			case PLC_CodeId == "J07.A03A08SELECTION1.LOSTFOCUS" && Comissao
				VGO_Dat.FOL_SqlExec("Select A13_Ukey From STAR_DATA@A03 (NOLOCK) Where A03.Ukey = ?J07.A03_Ukey","A03_A13")
				if used("A03_A13")
					select J07
					replace A13_Ukey with A03_A13.A13_Ukey
					VGO_Dat.FOL_Closetable("A03_A13")
				endif
				do Tgr_J07_A13_Update.Fxp
				PLO_Form.PAGEFRAME1.PAGE6.refresh()
				PLO_Form.PAGEFRAME1.PAGE1.refresh()

			case PLC_CodeId = "J07.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.FOL_UPDATE"
				*- Não permite lançamento de pedidos com cliente bloqueado/inativo/Em Estudo
				do tgr_j07_bloq.fxp with PLO_Form
 
			case PLC_CodeId = "J07.PAGEFRAME1.PAGE4.E10CL_UKEY.TEXTALIAS1.FOL_UPDATE"
				*- Caso a nf de entrada de rma seja alterada apago o cursor j07_e10_nf - 28112008
				do trg_j07_e10cl_ukey_fol_update.fxp with PLO_Form

			case PLC_CodeId = "J07.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.LOSTFOCUS"
				*- Não permite lançamento de pedidos com cliente bloqueado/inativo/Em Estudo
				do tgr_j07_bloq.fxp with PLO_Form

			case PLC_CodeId = "J07.PAGEFRAME1.PAGE3.J07_J07_010_C.LOSTFOCUS"
				*- Desconto Geral, novo procedimento de desconto dado para alguns clientes com isso fica desconto 
				*- Interno no item mais esse desconto geral, mais no item irá apresentar apenas o desconto individual				
				if PLO_Form.VON_FezCalcDescGeral = .f. and !empty(alltrim(j07.j07_010_c)) and PLO_Form.VOL_EDITING
					* Deixo verdadeira a variavel de deconto geral
					store .t. to VPL_RecalculaDescGeral	
					do tgr_j07_j07_010_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t.
				endif 
				if PLO_Form.VON_FezCalcDescGeral = .t. and !empty(alltrim(j07.j07_010_c)) and PLO_Form.VOL_EDITING
					* Deixo verdadeira a variavel de deconto geral
					store .t. to VPL_RecalculaDescGeral	
					do tgr_j07_j07_010_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t.
				endif	
				
			case PLC_CodeId = "J07.FOL_BEFORESAVE"
				PLO_Form.OOL_Comissao_Esp.J07_MontaComissao()
				*- Verifica liberação automatica no tipo de nota e também verifica se deve ser dado o update nos 
				*- campos de desconto geral do pedido de venda
				do Tgr_j07_BeforeSave.prg  with PLO_Form 
				
				if VGO_Custom.VOL_TriggerReturn
					*- Verifica se não há mais de um tipo de ST no pedido e realiza outras verificações para pode gravar o pedido
					VGO_Custom.VOL_TriggerReturn = FLL_ChkJ07(PLO_Form)
				endif
				
			case PLC_CodeId = "J07.FOL_EDIT"
				* Impede edição de pedido já liberado ou faturado
				do tgr_j07_edit_delete.fxp with PLO_Form

				*- Duplica o cursor de itens para resgatar a qtde anterior (usado na lista de faltas)
				select * from j08t into cursor TmpJ08Ant

			case PLC_CodeId = "J07.FOL_BEFOREDELETE"
				* Impede exclusão de pedido já liberado ou faturado
				do tgr_j07_edit_delete.fxp with PLO_Form
				if VGO_Custom.VOL_TriggerReturn
					do tgr_j07_fol_beforedelete.prg with PLO_Form
				endif

				*- Deleta itens da lista de faltas vinculadas ao pedido
				VGO_Gen.FOL_SqlExec("DELETE FROM STAR_DATA@JJ636 WHERE JJ636_PAR = 'J07' AND JJ636_UKEYP = ?j07.ukey")

			case PLC_CodeId = "J07.FOL_DELETE"
				* Corrige F04 F46 para pode excluir o pedido
				do tgr_j07_fol_delete.prg with PLO_Form

			case PLC_CodeId = "J07.LOAD"
				PLO_Form.addproperty("Voc_MyExecuteCursor")
				PLO_Form.Voc_MyExecuteCursor 	= ""
				
				* Propriedade que verifica se teve desconto geral no pedido
				PLO_Form.addproperty("VON_FezCalcDescGeral",.f.)
				PLO_Form.VON_FezCalcDescGeral = .f.
								
				* Propriedade será utilizada para saber o percentual do desconto geral do pedido de venda
				PLO_Form.addproperty("VOC_PercDescontoGeral","")
				PLO_Form.VOC_PercDescontoGeral = ""
					
			case PLC_CodeId = "J07.INIT"
				*- Configura gravação para os kits
				PLO_Form.FOL_SetTable(1,.t.,.null.,"JJ629T","JJ629","Ukey")
				*- Configura gravação para as substituições de itens RMA
				PLO_Form.FOL_SetTable(2,.t.,.null.,"J07_JJ632T","JJ632","Ukey")

				*- Desabilita alguns campos e altera o lay-out de outros
				do tgr_J07_edit.fxp with PLO_Form
				do tgr_J07_INIT.fxp with PLO_Form
				PLO_Form.newobject("OOL_Comissao_Esp", "J07Calcs", "j07_Comissao.fxp")  && Comissao1
				
				*- Cria o objeto responsavel pelo envio de e-mail para o envio de email para o representante
				*- Caso o pedido de venda tenha origem web e seja excluido
				PLO_Form.newobject("OOL_WebSales", "cgo_websales", "sgo_websales.fxp")  
				 
				*- Cria o objeto responsavel pelo calculo da m3
				PLO_Form.newobject("OOO_M3", "CGO_M3", "SGO_M3.FXP")
				
			case PLC_CodeId = upper("J07.pageFRAME1.page2.container1.cmd_new.Click") && Comissao1
				if empty(J07.A13_Ukey) or isnull(J07.A13_Ukey)
					messagebox("Favor Preencher antes a condicao de pagamento.",16,"Atenção!")
					vgo_gen.vol_trIGGERRETURN = .f.
				endif
				if vgo_gen.vol_trIGGERRETURN = .t.
					Vgo_Dat.FOL_SqlExec("Select A03_002_C From STAR_DATA@A03 (NOLOCK) Where A03.Ukey = ?J07.A03_Ukey","A03_002")
					if used("A03_002")
						if empty(j07.j07_010_c) and  "G BARBOSA" $ alltrim(A03_002.a03_002_c)
							messagebox("Favor Preencher antes o desconto geral do cliente.",16,"Atenção!")
							vgo_gen.vol_trIGGERRETURN = .f.
						endif
						Vgo_Dat.FOL_Closetable("A03_002")
					endif
				endif

			case PLC_CodeId = upper("J07.pageFRAME1.page2.container1.CMD_AUXBUTTON02.Click") && Comissao1
				if empty(J07.A13_Ukey) or isnull(J07.A13_Ukey)
					messagebox("Favor Preencher antes a condicao de pagamento.",16,"Atenção!")
					vgo_gen.vol_trIGGERRETURN = .f.
				endif
				if vgo_gen.vol_trIGGERRETURN = .t.
					Vgo_Dat.FOL_SqlExec("Select A03_002_C From STAR_DATA@A03 (NOLOCK) Where A03.Ukey = ?J07.A03_Ukey","A03_002")
					if used("A03_002")
						if empty(j07.j07_010_c) and  "G BARBOSA" $ alltrim(A03_002.a03_002_c)
							messagebox("Favor Preencher antes o desconto geral do cliente.",16,"Atenção!")
							vgo_gen.vol_trIGGERRETURN = .f.
						endif
						Vgo_Dat.FOL_Closetable("A03_002")
					endif
				endif

			case PLC_CodeId = upper("J07.pageFRAME1.page2.container1.CMD_AUXBUTTON03.Click") && Comissao1
				if empty(J07.A13_Ukey) or isnull(J07.A13_Ukey)
					=messagebox("Favor Preencher antes a condicao de pagamento.",16,"Atenção!")
					vgo_gen.vol_trIGGERRETURN = .f.
				endif
				if vgo_gen.vol_trIGGERRETURN = .t.
					Vgo_Dat.FOL_SqlExec("Select A03_002_C From STAR_DATA@A03 (NOLOCK) Where A03.Ukey = ?J07.A03_Ukey","A03_002")
					if used("A03_002")
						if empty(j07.j07_010_c) and  "G BARBOSA" $ alltrim(A03_002.a03_002_c)
							=messagebox("Favor Preencher antes o desconto geral do cliente.",16,"Atenção!")
							vgo_gen.vol_trIGGERRETURN = .f.
						endif
						Vgo_Dat.FOL_Closetable("A03_002")
					endif
				endif

			case PLC_CodeId == "J07.FOL_AFTERDELETE"
				do tgr_j07_fol_afterdelete.prg with PLO_Form

			case PLC_CodeId == "J07.FOL_EXECUTECURSOR"
				if 	!(PLO_Form.VOC_MyExecuteCursor == nvl(J07.ukey, "")) or vgo_gen.voa_TriggerParameters[1]
					PLO_Form.Voc_MyExecuteCursor = J07.Ukey
										
					with VGO_Dat
						.Voa_Index[1] = "JJ629_UkeyP"
						.FOL_SetParameter(1,"J07")
						.FOL_SetParameter(2,J07.Ukey)
						.FOL_EditCursor("JJ629_FRM","JJ629_FRM","JJ629T","Z",1,2)
					endwith

					with VGO_Dat 					&& Abrindo o cursor do grid de substituição de itens RMA
						.VOA_INDEX[1]= "J07_JJ632T.UKEY"
						.FOL_SetParameter(1,J07.Ukey)
						.FOL_EditCursor("JJ632_ESTRU", "JJ632_ESTRU", "J07_JJ632T", "Z", 1, 1)
					endwith
					
					with VGO_Dat && Cria o cursor para a lista de faltas
						.FOL_EditCursor("JJ636_FRM", "JJ636_FRM", "JJ636T", "Z", 0, 0)
					endwith
					
					* Informação para apresentar número do pedido web.
					Vgo_Dat.FOL_SqlExec("Select Wj07_001_C From STAR_DATA@WJ07 (NOLOCK) Where WJ07.J07_UKEY = ?J07.UKEY","WJ07_TMP")
								
					if vartype(PLO_Form.wj07_wj07_001_c) = "O"
						if reccount('WJ07_TMP') > 0
							PLO_Form.wj07_wj07_001_c.visible = .t.
							PLO_Form.label_wj07_001_c.visible = .t.
						else
							PLO_Form.wj07_wj07_001_c.visible = .f.
							PLO_Form.label_wj07_001_c.visible = .f.
						endif

						PLO_Form.wj07_WJ07_001_C.tooltiptext = "Número do pedido na WEB"
					endif
					
					* Setando a variavel de controle para pedidos de venda com Desconto geral 
					if j07.j07_011_b <> 0
						* Propriedade que verifica se teve desconto geral no pedido
						PLO_Form.VON_FezCalcDescGeral = .T.
						PLO_Form.VOC_PercDescontogeral = j07.j07_010_c
					else
						PLO_Form.VON_FezCalcDescGeral = .F.	
						PLO_Form.VOC_PercDescontogeral = ""
					endif				
				endif
		
			case PLC_CodeId == "J07.PAGEFRAME1.PAGE4.J07_A14_UKEY.TEXTALIAS1.INIT"
				*- Torna o campo transportadora obrigatorio
				PLO_Form.pageframe1.page4.j07_a14_ukey.textalias1.VOL_ChkIfEmpty = .T.
		endcase


	case VLC_FormName == "J07_004"					&&-Notas Fiscal utilizando um pedido
		do case
			case PLC_CodeId = "J07_004.FOL_SAVE"
				*- Verifica se o pedido foi liberado ao utiliza-lo em uma NF.
				do tgr_j07_004_save.fxp with PLO_Form

			case PLC_CodeId = "J07_004.UNLOAD"
				*- ativa variável para correção do desconto
				store .t. to VPL_CORRIGE_DESCONTO
				
				*- Busca os dados do volume do pedido transportado
				do tgr_j10_volumes.fxp WITH PLO_Form.VOC_Pedido
				
				*- Transporte de campos novos de desconto para a nf
				do tgr_j10_desc.fxp with PLO_Form.VOC_Pedido
		endcase

	case VLC_FormName == "J07_LIBERACAO"	 &&& Tela de Liberação de Pedido
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				VGO_Custom.FOL_cAddtoolbarapps ("Posição do Financeiro", "VGO_Custom.FOL_cDoForm ('J07_LIBPOSFIN.scx') ", not VGO_Dat.FOL_AcessMenu("J07_LIBPOSFIN") and not empty(j07tmp.ukey))
				VGO_Custom.FOL_cAddtoolbarapps ("Totalizações", "VGO_Custom.FOL_cDoForm ('J07_TOT.scx') ", not VGO_Dat.FOL_AcessMenu("J07_TOT"))
		endcase

	case VLC_FormName == "J07_CONFIRMACAO"	 && Comissao
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				VGO_Custom.FOL_cAddtoolbarapps ("Totalizações", "VGO_Custom.FOL_cDoForm ('J07_TOT.scx') ", not VGO_Dat.FOL_AcessMenu("J07_TOT"))
		endcase

	case VLC_FormName == "J08"	&&-Itens do Pedido de venda
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Lançamento de Kits", "VGO_Custom.FOL_cDoForm('JJ629.scx',_screen.activeform)", not PLO_Form.VOL_EDITING and not _screen.activeform.VOL_NoEdit)
				VGO_Custom.FOL_cAddtoolbarapps("Substituição de Itens de RMA", "VGO_Custom.FOL_cDoForm('JJ632')", PLO_Form.vol_adding = .T. or PLO_Form.VOL_EDITING = .T.)
				VGO_Custom.FOL_cAddtoolbarapps("Configuração de E/S utilizada", "VGO_Custom.FOL_cDoForm('J02J08J11_T06')", .T.)

			case PLC_CodeId = "J08.D04D33SELECTION1.FOL_UPDATE"
				if "." $ PLO_Form.D04D33SELECTION1.TEXTALIASD04.TEXTALIAS1.value and empty(nvl(j08t.jj629_Ukey,""))
					messagebox("Lancamento manual de kit ou parte de kit não permitido."+chr(13)+;
							   "Favor utilizar a tela de lançamento de Kits na lista de aplicações  ",48,"Atenção")

					replace d04_ukey with null in j08t
					PLO_Form.D04D33SELECTION1.TEXTALIASD04.refresh()
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId = "J08.J08_D01_UKEY.REFRESH"
				PLO_Form.J08_D01_UKEY.Vol_Enabled = .f.

			case PLC_CodeId = "J08.J08_D02_UKEY.REFRESH"
				PLO_Form.J08_D02_UKEY.Vol_Enabled = .f.

			case PLC_CodeId $ "J08.J08_J08_005_B.FOL_UPDATE"
				PLO_Form.ool_J08_Functions.Von_LastTablePrice = PLO_Form.J08_J08_005_B.value

			case PLC_CodeId $ "J08.FOL_CLOSE" && Comissao
				if !PLO_Form.Vol_Noedit
					do Tgr_J08_close.fxp
					PLO_Form.OOL_Comissao_Esp.J07_MontaComissao()
				endif

			case PLC_CodeId $ "J08.J08_T04_UKEY.TEXTALIAS1.GOTFOCUS"  &&Tipo de E/S
				* Desabilita alguns campos para Vendas
				do tgr_j08_disable_vendas.fxp with PLO_Form
				
			case PLC_CodeId $ "J08.J08_T02_UKEY0.TEXTALIAS1.GOTFOCUS" &&Unidade de Medida
				* Desabilita alguns campos para Vendas
				do tgr_j08_disable_vendas.fxp with PLO_Form
				
			case PLC_CodeId $ "J08.J08_J08_003_B.SPINNER1.GOTFOCUS" && Quantidade
				* Desabilita alguns campos para Vendas
				do tgr_j08_disable_vendas.fxp with PLO_Form
				* Deixa o campo quantidade como inteiro ou decimal dependendo do tipo de Saída
				do tgr_j08_j08_003_b_gotfocus.fxp with PLO_Form
				
				*- Guarda valor unitario inicial para o item
				if PLO_Form.j08_j08_005_b.spinner1.VON_InitialValue <= 0
					PLO_Form.j08_j08_005_b.spinner1.VON_InitialValue = PLO_Form.j08_j08_005_b.value
				endif 
				
				*- Caso o item mudado atualiza valor unitario bruto
				if !(j08t.d04_ukey == PLO_Form.j08_j08_604_c.VOC_D04Ukey) AND PLO_Form.VOL_Adding
					replace j08_606_b with j08_005_b in j08t
					PLO_Form.j08_j08_606_b.refresh()	
					PLO_Form.j08_j08_604_c.VOC_D04Ukey = j08t.d04_ukey
				endif
				
			case PLC_CodeId $ "J08.J08_J08_003_B.SPINNER1.FOL_UPDATE"  && Quantidade
				* Essa trigger tem diversas funcionalidades
				do tgr_J08_local.FXP with PLO_Form
				do tgr_J08_Fol_Update.FXP with PLO_Form
				PLO_Form.j08_j08_005_b.setfocus()
				 
			case PLC_CodeId $ "J08.J08_J08_005_B.SPINNER1.LOSTFOCUS" && Valor Unitário
				* Calcula Desconto para o Magazine
				do tgr_j08_005_b_lostfocus.fxp with PLO_Form
				
			case PLC_CodeId $ "J08.J08_J08_005_B.SPINNER1.FOL_UPDATE"
				VLC_Percentual = PLO_Form.j08_j08_604_c.FOC_EncontraPercDesc(j08t.j08_005_b, j08t.j08_606_b)
				
				replace j08t.j08_604_c with VLC_Percentual in j08t
				replace j08t.j08_605_b with iif(empty(VLC_Percentual), 0, (j08t.j08_606_b - j08t.j08_005_b) * j08t.j08_003_b) in j08t
				PLO_Form.j08_j08_605_b.refresh()	

			case PLC_CodeId $ "J08.INIT"
				*- Deixa o campo quantidade como inteiro e valor do item com duas casas decimais
				do tgr_J08_init.fxp with PLO_Form && cria também objeto recalculo da ST
				PLO_Form.newobject("OOL_Comissao_Esp", "J07Calcs", "j07_Comissao.fxp") && Comissao

			case PLC_CodeId $ "J08.FOL_ADDNEW" 
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j08_j08_604_c.VOC_D04Ukey = ""

			case PLC_CodeId $ "J08.FOL_SAVE"
				do FLL_VerifCES.fxp with "J08T"
				
				if VGO_Custom.VOL_TriggerReturn
					*- Verifica duplicidade de itens no pedido
					do tgr_j08_003_b_itemduplicado.fxp with PLO_Form
					
					if VGO_Custom.VOL_TriggerReturn
						* Corrige desconto e outras funcionalidades
						do tgr_j08_save.fxp with PLO_Form
					endif
				endif
				
				if VGO_Custom.VOL_TriggerReturn
					*- Executa rotina que checa disponibilidade no estoque e caso não tenha envia para a lista de faltas
					do tgr_j08_avaliadisp.fxp with PLO_Form
				endif

			case "J08.ACTIVATE" $ upper(PLC_CodeId)
				*- Deixa o campo quantidade como inteiro e valor do item com duas casas decimais
				*do tgr_J08_volta_600.fxp with PLO_Form
				* Programa abaixo serve tanto para o calculo desconto geral como a troca de condição de pagamento
				do j08_Recalculo_DescontoGeral.fxp with PLO_Form
							
				*- Insere o campo de Qrde de Embalagem abaixo da Qtde Disponivel para o item
				if vartype(PLO_Form.VOO_FormSec) == "O"
					*- Cria o objeto responsavel por atualizar a Qtde de Embalagem
					if vartype(PLO_Form.VOO_FormSec.OOO_Package) <> "O"
						PLO_Form.VOO_FormSec.newobject("OOO_Package", "CGO_Package", "SGO_Package.FXP")
					
						*- Atualiza a Qtde de Embalagem quando rodar o refresh da Qtde disponivel
						bindevent(PLO_Form.VOO_FormSec.stockbalanceavailable1.spinner1, "REFRESH", PLO_Form.VOO_FormSec.OOO_Package, "FOL_ShowPackage")
					endif
				endif

			case PLC_CodeId $ "J08.FOL_BEFOREDELETE" and wontop()="J08"		&&- Uso o Wontop() se não a trigger também é acionada pela tela de Lançamento de Kit
				if not empty(nvl(j08t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser alterado ou excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId $ "J08.REFRESH"
				*- Deixa o campo quantidade como inteiro ou decimal dependendo do tipo de Saída
				do tgr_j08_disable_vendas.fxp with PLO_Form
				*- Deixa o campo quantidade como inteiro ou decimal dependendo do tipo de Saída
				do tgr_j08_j08_003_b_gotfocus.fxp with PLO_Form

			case PLC_CodeId == "J08.FOL_EDIT"
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j08_j08_604_c.VOC_D04Ukey = ""
								
				*- Propriedade usada para resgatar o valor digitado para comparar disponibilidade em caso de edição do item
				PLO_Form.addproperty("VON_QtdeAnterior", 0)
				
				VGO_Gen.FOL_SqlExec("SELECT J08_003_B FROM STAR_DATA@J08 (NOLOCK) WHERE UKEY = ?J08T.UKEY", "TmpJ08")
				PLO_Form.VON_QtdeAnterior = TmpJ08.J08_003_B
				VGO_Gen.FOL_CloseTable("TmpJ08")
			
				*- Atualiza o campo valor liquido
				do tgr_J08_edit.fxp with PLO_Form
			
				PLO_Form.d04d33selection1.textaliasd04.textalias1.VOC_FirstUkey = j08t.d04_ukey
				PLO_Form.d04d33selection1.textaliasd04.textalias1.VOC_LastUkey = j08t.d04_ukey
				
				*- Nao permito alterar nenhuma informação do item se o item pertencer a um kit
				if not empty(nvl(j08t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser alterado ou excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLC_CodeId == "J08.FOL_MAKESEARCH"				
				*- Atualiza o campo valor liquido
				do tgr_J08_edit.fxp with PLO_Form
				
				*- Nao permito alterar nenhuma informação do item se o item pertencer a um kit
				if not empty(nvl(j08t.jj629_ukey,""))
					messagebox("O item pertence a um Kit e não pode ser alterado ou excluído  ",48,"Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
					return .f.
				endif

			case PLO_Form.VOL_EDITING
				*- Desabilita alguns campos ao adicionar/editar
				do tgr_J08_edit.fxp with PLO_Form

*!*				case not ("LOAD" $ PLC_CodeId) and not ("INIT" $ PLC_CodeId) and not ("EXECUTECURSOR" $ PLC_CodeId)
*!*					*- Atualiza o valor do campo Valor Unitário Líquido
*!*					do tgr_j08_liq.fxp with PLO_Form
			
		endcase


	case VLC_FormName == "J10" 						&&-Nota Fiscal de Venda

		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				do tgr_j10_barra.fxp with PLO_Form

			case PLC_CodeId  $ "J10.ACTIVATE"
				if VPL_CORRIGE_DESCONTO = .t.
					* procedimento para abrir a tela de itens apos a importação para nota fiscal
					do tgr_j10_corrige_desconto.fxp with PLO_Form
					PLO_Form.FOL_makeparcelas() && refaz as parcelas na importação do pedido apos correção do desconto
				endif

			case PLC_CodeId = "J10.PAGEFRAME1.PAGE3.J10_J10_033_C.LOSTFOCUS"
				*- Desconto Geral, novo procedimento de desconto dado para alguns clientes com isso fica desconto 
				*- Interno no item mais esse desconto geral, mais no item irá apresentar apenas o desconto individual				
				if PLO_Form.VON_FezCalcDescGeral = .f. and !empty(alltrim(j10.j10_033_c)) and PLO_Form.VOL_EDITING
					do tgr_j10_j10_033_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t.
				endif 
				if PLO_Form.VON_FezCalcDescGeral = .t. and !empty(alltrim(j10.j10_033_c)) and PLO_Form.VOL_EDITING
					do tgr_j10_j10_033_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t.
				endif		
				
			case PLC_CodeId $ "J10.FOL_ADDNEW"
				*- Valoriza o campo serie e peso ao adicionar
				do tgr_j10_addnew.fxp with PLO_Form

			case PLC_CodeId $ "J10.LOAD"
				* Propriedade que verifica se teve desconto geral na nota
				PLO_Form.addproperty("VON_FezCalcDescGeral",.f.)
				PLO_Form.VON_FezCalcDescGeral = .f.
				
			case PLC_CodeId = "J10.INIT"	&& 'J10.INIT'
				* Cria a Váriavel auxiliar na correção do desconto
				store .f. to vpl_corrige_desconto
				* Desabilita alguns campos e altera o lay-out de outros
				do tgr_J10_init.fxp with PLO_Form 

			case PLC_CodeId = "J10.PAGEFRAME1.PAGE2.CONTAINER1.CMD_AUXBUTTON02.INIT"
				* Desabilita o botão importação do pedido por dentro do item
				PLO_Form.PAGEFRAME1.PAGE2.CONTAINER1.CMD_AUXBUTTON02.visible = .f.

			case PLC_CodeId = "J10.FOL_BEFOREDELETE"
				do tgr_J10_fol_beforedelete.fxp

			case PLC_CodeId = "J10.FOL_BEFORESAVE"				
				*- Desconto Geral, novo procedimento de desconto dado para alguns clientes com isso fica desconto 
				*- Interno no item mais esse desconto geral, mais no item irá apresentar apenas o desconto individual				
				if PLO_Form.VON_FezCalcDescGeral = .f. and !empty(alltrim(j10.j10_033_c)) and PLO_Form.VOL_EDITING
					do tgr_j10_j10_033_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t.
				endif 
				if PLO_Form.VON_FezCalcDescGeral = .t. and !empty(alltrim(j10.j10_033_c)) and PLO_Form.VOL_EDITING
					do tgr_j10_j10_033_c_lostfocus.fxp with PLO_Form
					PLO_Form.VON_FezCalcDescGeral = .t. 
				endif	
												
				* Deixa o campo data de saida da NF em branco
				do tgr_J10_beforesave.fxp with PLO_Form	&& Filipe ativou 2/12/2008
				do tgr_j10_LinkComission with PLO_Form && Acrescentar esta linha!!!!!!!!!!!
			
			case PLC_CodeId = "J10.FOL_SAVE"
				*- Busca os dados do volume do pedido transportado
				do tgr_j10_volumes.fxp with j10.j07_ukey
				
			case PLC_CodeId = "J10.FOL_AFTERSAVE"
				* Calcula a M3 da NF
				*do TGR_J10_AFTERSAVE_INTEGRACAO with PLO_Form
				do tgr_J10_aftersave_dtasaida.fxp with PLO_Form
				do tgr_j10_fol_aftersave.fxp with PLO_Form
				
			case PLC_CodeId == "J10.REFRESH"  && Filipe Oliveira
				
				do TGR_J10_REFRESH with PLO_Form							
*!*				PLO_Form.A03A08SELECTION1.A03_TEXT.tEXTALIAS1.VOC_Searchname= "J07_A03_ESP"
*!*				PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.VOC_cursorcat = "J07_A03_ESP"
*!*				PLO_Form.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.VOC_CUrsorcatrecords = "J07_A03_COUNT"

			case PLC_CodeId = "J10.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.FOL_UPDATE"
				* Não permite lançamento de pedidos com cliente bloqueado/inativo/Em Estudo - Filipe 22092008
				do tgr_j10_bloq.fxp with PLO_Form

			case PLC_CodeId = "J10.A03A08SELECTION1.A03_TEXT.TEXTALIAS1.LOSTFOCUS"
				* Não permite lançamento de pedidos com cliente bloqueado/inativo/Em Estudo - Filipe 22092008
				do tgr_j10_bloq.fxp with PLO_Form

			case PLC_CodeId = "J10.J10_J09_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de serie da NF
				with PLO_Form					 
					.j10_j09_ukey.TEXTALIAS1.vol_chkifempty = .t.
				endwith					
			
			case PLC_CodeId = "J10.PAGEFRAME1.PAGE1.J10_T88_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Especie da NF
				with PLO_Form.pageframe1.page1
					.j10_t88_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith

			case PLC_CodeId = "J10.PAGEFRAME1.PAGE1.CGO_TEXTALIASLABEL_T89.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Modelo da NF
				with PLO_Form.pageframe1.page1
					.cgo_textaliaslabel_t89.textalias1.VOL_ChkIfEmpty = .t.
				endwith
		endcase


	case VLC_FormName == "J10_001"				&&-Devolução de Saída

		do case
			
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				*-- Adiciona a opção de informação do Protocolo de intenção - CASOUZA - 23/08/07
				VGO_Custom.FOL_cAddtoolbarapps("PTA - Protocolo de Intenção", "VGO_Custom.FOL_cDoForm('j10cl_pta')", !eof("J10"))

			case PLC_CodeId $ "J10_001.ACTIVATE"
				if VPL_CORRIGE_DESCONTO = .t.
					* Corrige O desconto apos a importação para nota de devolução
					do tgr_j10_001_corrige_desconto.fxp with PLO_Form
					PLO_Form.FOL_makeparcelas() && refaz as parcelas na importação do pedido apos correção do desconto
				endif
		
			case PLC_CodeId $ "J10_001.FOL_ADDNEW"
				*Torna o campo Numero da Nota Fiscal não obrigatório, mas depois volta no fol_beforesave
				with PLO_Form.j10_j10_001_c
					.vol_chkifempty = .f.
					.vol_enabled = .t.
					.enabled = .t. 
				endwith

			case PLC_CodeId $ "J10_001.FOL_BEFORESAVE"
				* Busca numeração de NF de Venda para Entrada de Importação
				do tgr_j10_001_beforesave.fxp with PLO_Form

			case PLC_CodeId == "J10_001.REFRESH"  && Filipe Oliveira 09/09/2008
				with PLO_Form.A03A08SELECTION1.A03_TEXT.tEXTALIAS1
					.VOC_Searchname= "J07_A03_ESP"
					.VOC_cursorcat = "J07_A03_ESP"
					.VOC_CUrsorcatrecords = "J07_A03_COUNT"
				endwith
				
			case PLC_CodeId = "J10_001.PAGEFRAME1.PAGE1.J10_T88_UKEY.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Especie da NF
				with PLO_Form.pageframe1.page1
					.j10_t88_ukey.textalias1.VOL_ChkIfEmpty = .t.
				endwith

			case PLC_CodeId = "J10_001.PAGEFRAME1.PAGE1.CGO_TEXTALIASLABEL_T891.TEXTALIAS1.INIT"
				*- Torna obrigatório o campo de Modelo da NF
				with PLO_Form.pageframe1.page1
					.cgo_textaliaslabel_t891.textalias1.VOL_ChkIfEmpty = .t.
				endwith
	endcase
  
	case VLC_FormName == "J10_006"				&&- Tela Geração de Nota Fiscal Eletrônica - NF-e
		do case
				case PLC_CodeId $ "J10_006.INIT"
					with PLO_Form.Container1
						.TxtFileName.Arq_txt.value = "R:\INTERCAMBIO\NFE\GERADOS\"
						.TxtVersionAplic.value = "2.2.0"	
						.DtSaida.Value = 1						
					endwith	 
					
					with PLO_Form
					    .VON_Enquad = 2
						.VON_Versao = 2
			 		endwith	 
					 
					*- programa que altera a select para criar o desconto do suframa no campo desconto da NFe.
					do tgr_j10_006_init.FXP
					
				case PLC_CodeId  == "J10_006.CANCEL.CLICK"
		                do tgr_j10_006_fol_close.fxp with PLO_Form	
		                
		endcase
 
	case VLC_FormName == "J10_011"				&&- Tela Confirmação de Nota Fiscal Eletrônica - NF-e
		do case
				case PLC_CodeId $ "J10_011.INIT"
					 PLO_Form.Container1.TxtFileName.Arq_txt.value = "R:\INTERCAMBIO\NFE\AUTORIZADOS\"
			   		 
               	*- Envia por email o XML da NFe para os clientes
                case PLC_CodeId $ "J10_011.FOL_CLOSE"
               	   		 do tgr_J10_011_fol_chkexecute.fxp with PLO_Form	 
		endcase

	case VLC_FormName == "J11"						&&-Itens da Nota Fiscal

		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				VGO_Custom.FOL_cAddtoolbarapps("Configuração de E/S utilizada", "VGO_Custom.FOL_cDoForm('J02J08J11_T06')", .T.)
					
				*Deixa o campo quantidade como inteiro e valor do item com duas casas decimais
			case PLC_CodeId $ "J11.INIT"
				do tgr_J11_init.fxp with PLO_Form &&cria objeto recalculo da ST
						
			case PLC_CodeId $ "J11.FOL_ADDNEW"
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j11_j11_604_c.VOC_D04Ukey = ""
				
			case PLC_CodeId $ "J11.FOL_EDIT"
				*- Limpa a variavel de controle do item obteve o desconto
				PLO_Form.j11_j11_604_c.VOC_D04Ukey = ""	
				
			case PLC_CodeId $ "J11.J11_J11_005_B.SPINNER1.FOL_UPDATE" 
					VLC_Percentual = PLO_Form.j11_j11_604_c.FOC_EncontraPercDesc(j11t.j11_005_b, j11t.j11_606_b)
				
					replace j11t.j11_604_c with VLC_Percentual in j11t
					replace j11t.j11_605_b with iif(empty(VLC_Percentual), 0, (j11t.j11_606_b - j11t.j11_005_b) * j11t.j11_003_b) in j11t
					PLO_Form.j11_j11_604_c.refresh()
					PLO_Form.j11_j11_605_b.refresh()
							
			Case PLC_CodeId $ "J11.J11_J11_003_B.SPINNER1.GOTFOCUS"   && Quantidade
				if !(j11t.d04_ukey == PLO_Form.j11_j11_604_c.VOC_D04Ukey) AND PLO_Form.VOL_Adding
					replace j11_606_b with j11_005_b in j11t
					PLO_Form.j11_j11_606_b.refresh()	
					PLO_Form.j11_j11_604_c.VOC_D04Ukey = j11t.d04_ukey
				endif
						
			case PLC_CodeId $ "J11.J11_J11_003_B.SPINNER1.FOL_UPDATE" && Quantidade
				*- Colocando o focus no campo valor unit para calculo novo de desconto 
				PLO_Form.j11_j11_005_b.setfocus()
								
			case PLC_CodeId $ "J11.FOL_SAVE"
				do FLL_VerifCES.fxp with "J11T"
				do tgr_J11_fol_save.fxp with PLO_Form
				*PLO_Form.refresh()
				
			case PLC_CodeId $ "J11.ACTIVATE"
			
				* Chamando o meu recalculo de desconto geral
				if !empty(alltrim(j10.j10_033_c))
					do j11_Recalculo_DescontoGeral.fxp with PLO_Form
				endif			
			
				*- Insere o campo de Qrde de Embalagem abaixo da Qtde Disponivel para o item
				if vartype(PLO_Form.VOO_FormSec) == "O"				
					*- Cria o objeto responsavel por atualizar a Qtde de Embalagem
					if vartype(PLO_Form.VOO_FormSec.OOO_Package) <> "O"
						PLO_Form.VOO_FormSec.newobject("OOO_Package", "CGO_Package", "SGO_Package.FXP")
						
						*- Atualiza a Qtde de Embalagem quando rodar o refresh da Qtde disponivel
						bindevent(PLO_Form.VOO_FormSec.stockbalanceavailable1.spinner1, "REFRESH", PLO_Form.VOO_FormSec.OOO_Package, "FOL_ShowPackage")
					endif
				endif
			
			* Desabilita alguns campos ao adicionar/editar
			case PLO_Form.VOL_EDITING
				do tgr_J11_edit.fxp with PLO_Form

*!*				case not ("LOAD" $ PLC_CodeId) and not ("INIT" $ PLC_CodeId) and not ("EXECUTECURSOR" $ PLC_CodeId) && Comissao
*!*					* Atualiza o valor do campo Valor Unitário Líquido
*!*					if VGO_gen.voc_companycomplement == 'EMP02'
*!*						PLO_Form.liq.value = round((PLO_Form.J11_J11_008_B.value/PLO_Form.J11_J11_003_b.value),2)
*!*					endif

		endcase

	case VLC_FormName == "J11_001"					&&-Itens da Nota Fiscal de Devolução
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Adicionando na Lista de Aplicações
				VGO_Custom.FOL_cAddtoolbarapps("Configuração de E/S utilizada", "VGO_Custom.FOL_cDoForm('J02J08J11_T06')", .T.)
			
			case PLC_CodeId $ "J11_001.INIT"
				do tgr_J11_001_init.fxp with PLO_Form &&cria objeto recalculo da ST
				bindevent(PLO_Form,"fol_calcimp",PLO_Form.cgo_recalculast,"fol_execute",1) &&inclue na calcimp
			
			case PLC_CodeId $ "J11_001.J11_J11_003_B.SPINNER1.FOL_UPDATE" && Quantidade
				do tgr_j11_001_qtde.fxp

			case PLC_CodeId $ "J11_001.FOL_AFTERSAVE"
				PLO_Form.CGO_RECALCULAST.FOL_EXECUTE() &&recalcula se tiver ST
				PLO_Form.ool_j11_functions.FOL_CalcFormulas() &&atualia formulas
					 
				* Desabilita alguns campos ao adicionar/editar
			case PLO_Form.VOL_EDITING
				do tgr_j11_001_edit.fxp with PLO_Form

			case PLC_CodeId $ "J11_001.FOL_SAVE"
				do FLL_VerifCES.fxp with "J11T"
			
		endcase



	case VLC_FormName == "J11_002"					&&-Nota Fiscal de Devolução Utilizando um Nota Fiscal de Venda
		do case
			case PLC_CodeId $ "J11_002.UNLOAD"
				*- ativa variável para correção do desconto
				store .t. to VPL_CORRIGE_DESCONTO
				
				if used("TmpJ11JJ644")
					VGO_Gen.FOL_PushSelect()
					
					select TmpJ11JJ644
					
					scan
						replace j11t.jj644_ukey with TmpJ11JJ644.jj644_ukey in j11t for j11t.d04_ukey = TmpJ11JJ644.d04_ukey and j11t.j11_ukeyp = TmpJ11JJ644.ukey
					endscan
					
					select j11t
					go top
					
					VGO_Gen.FOL_CloseTable("TmpJ11JJ644")
					
					VGO_Gen.FOL_PopSelect()
				endif
				
			case PLC_CodeId = "J11_002.DOIT.CLICK"
				*- Aqui vou tirar o desconto geral da nota de saida, está sendo desconto 2 vezes por isso vou limpar o campo j10_033_c do cursor j11_d
				do trg_j11_002_doit_click.fxp with PLO_Form
				
			case PLC_CodeId = "J11_002.CONTAINER1.GRIDCONTAINER1.CMD_AUXBUTTON01.INIT"
				PLO_Form.container1.gridcontainer1.cmd_auxbutton01.VOC_SearchName = "J11_AUTDEV"
		endcase



	case VLC_FormName == "J13"					&&-Previsão de vendas
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"	&& Acionando a Barra de Ferramentas APPS
				VGO_Custom.FOL_cAddtoolbarapps("Utiliza Quantidade Média de Vendas", "VGO_Custom.FOL_cDoForm('J13_utilqmv.scx')", PLO_Form.VOL_EDITING and empty(nvl(J13.JJ607_UKEY,'')))
				VGO_Custom.FOL_cAddtoolbarapps("Previsão de Vendas - 12 Meses", "VGO_Custom.FOL_cDoForm('JJ642.scx')", !PLO_Form.VOL_Editing and !empty(nvl(J13.UKEY,'')))

			case PLC_CodeId $ "J13.INIT"
				do TGR_J13_INIT with PLO_Form
			case PLC_CodeId $ "J13.REFRESH"
				do TGR_J13_IMAGE1 with PLO_Form
			case PLC_CodeId $ "J13.FOL_EDIT"
				do TGR_J13_VERIFICA with 'E'
			case PLC_CodeId $ "J13.FOL_BEFOREDELETE"
				do TGR_J13_VERIFICA with 'D'
				
				do tgr_j13_fol_beforedelete.fxp
				
			case PLC_CodeId $ "J13.FOL_SAVE"
				VGO_Gen.VOL_TriggerReturn = FLL_ExportPrevisao()
				
			case PLC_CodeId $ "J13.FOL_AFTERSAVE"
				VGO_Gen.VOL_TriggerReturn = FLL_PrevDozeMeses() && Previsão de Vendas - 12 Meses
		endcase

	case VLC_FORMNAME == "J14"
		do case
			case PLC_CodeId $ "J14.INIT"
				do TGR_J14_INIT with PLO_Form
		endcase

	case VLC_FormName == "JJ643" 			&&- Autorização de Devolução
		do case
			case  PLC_CodeId == "LOADTOOLBARAPPS"
				VGO_Custom.FOL_cAddtoolbarapps("Recebimento", "VGO_Custom.FOL_cDoForm('jj643_receb.scx')", reccount("jj643") > 0 and not VGO_Dat.FOL_AcessMenu("JJ643_RECEB.JJ643") and !PLO_Form.VOL_Editing and !PLO_Form.VOL_Adding)
		endcase


	case VLC_FormName == "K01"				&&-Materiais do PCP

		if type("PLO_Form.pageframe1.page1.k01_k01_001_c")#"U"
			with PLO_Form.pageframe1.page1
				.k01_k01_001_c.vol_enabled = .f.
				.k01_k01_002_c.vol_enabled = .f.
			endwith
		endif

		if PLO_Form.VOL_EDITING
			* Habilitar o COMBO Categoria do Material
			do tgr_K01flexfor_edit.fxp with PLO_Form
		endif


		do case
			case PLC_CodeId == "K01.PAGEFRAME1.PAGE1.CGO_STOCKGRIDSELECTION1.CGO_TEXTALIASLABEL_D041.TEXTALIAS1.FOL_UPDATE"
				do tgr_K01_D04_Update.prg with plo_source, PLO_Form

			case PLC_CodeId == "K01.INIT"	
				do tgr_K01_INIT with PLO_Form

			case PLC_CodeId == "K01.FOL_SAVE"
				do tgr_K01_SAVE with PLO_Form

			case PLC_CodeId == upper("K01.PAGEFRAME1.PAGE1.CGO_stockgridselection1.Cgo_textaliaslabel_d041.init") && Acerta a picture do Item do Estoque
				with PLO_Form.pAGEFRAME1.pAGE1.Cgo_stockgridselection1.Cgo_textaliaslabel_d041.tEXTALIAS1
					.inputmask = vgo_gen.voa_picts[23,2]
					.format = vgo_gen.voa_picts[23,1]
				endwith
		endcase

	case VLC_FormName == "K11"				&&- Operações
		do case
			case PLC_CodeId == "K11.INIT"
				do K11_INIT with PLO_Form
		endcase

	case VLC_FormName == "K12"				&&-REcursos da operação
		do case
			case PLC_CodeId == "K12.INIT"
				do K12_INIT with PLO_Form
		endcase

	case VLC_FormName == "K13"				&&-Estrutura
		do case
			case PLC_CodeId == "K13.INIT"
				do K13_INIT with PLO_Form

			case "CTTREE.CLICK" $ upper(PLC_CodeId)
				do tgr_k13_click.fxp with PLO_Form

			case PLC_CodeId == "K13.UNLOAD"
				if vartype(OOO_ToolbarRev)="O"
					OOO_ToolbarRev.release
					release OOO_ToolbarRev
				endif

			case PLC_CodeId == "K13.FOL_EDIT"
				do tgr_k13_fol_edit.fxp with PLO_Form
		endcase

	case VLC_FormName == "K14"				&&-Recursos da operação da estrutura
		do case
			case PLC_CodeId == "K14.INIT"
				do K14_INIT with PLO_Form
		endcase

	case VLC_FormName == "K16"				&&-Componentes da estrutura
		do case
			case PLC_CodeId == "K16.REFRESH"
				*- Força a variável que trava a inserção de materiais substitutos com .F. para possibilitar a inserção dos memsmos sempre
				PLO_Form.VOL_Locked = .F.
		endcase
		
	case VLC_FormName == "K19"				&&-Ordem de produção
		do case
			case PLC_CodeId == "K19.K19_K19_009_B.SPINNER1.FOL_UPDATE"
				do TGR_K19_009_B_Update
		endcase

	case VLC_FormName == "K28"						 &&-Calculo do MRP
		do case
			case PLC_CodeId == "K28.INIT"
				DO tgr_k28_init WITH PLO_Form   &&filipe reativou dia 11112009 solicitação de Débora PCP

			case PLC_CodeId == "K28.COMMANDBUTTON2.CLICK"
				Vln_data = set("Datasession")
				do Tgr_K28_K31_Desmembra.prg
				PLO_Form.commandbutton1.click()					&&-Roda o calculo padrao do sistema.
				*!*					DO tgr_k28_commandbutton2.fxp WITH PLO_Form		&&-Roda o nosso calculo para ajuste.
				set datasession to Vln_data
				do Tgr_K28_K31_Realinha.prg
				do Tgr_K28_K34_Valores.prg		&& Valores de Compras

			case upper(PLC_CodeId) == upper("K28.Refresh")
				if type("PLO_Form.pageframe1.page5.gridcontainer3.grid1.ColumnDesc")#"O"
					** No init dava problema!
					with PLO_Form.PAGEFRAME1
						do K01_Descricao_Grid with .page5.gRIDCONTAINER3.gRID1,2,"k32t.MATERIAL"
						do K01_Descricao_Grid with .PAGE4.GRIDCONTAINER2.GRID1,2,"k31t.MATERIAL"
						do K01_Descricao_Grid with .PAGE6.GRIDCONTAINER4.GRID1,2,"k33t.MATERIAL"
						do K01_Descricao_Grid with .PAGE7.GRIDCONTAINER5.GRID1,2,"k34t.MATERIAL"
						do K01_Descricao_Grid with .PAGE8.GRIDCONTAINER6.GRID1,2,"k35t.MATERIAL"
						do K01_Descricao_Grid with .PAGE9.GRIDCONTAINER4.GRID1,2,"k54t.K01_001_C"
					endwith
					
					with PLO_Form.pageframe1.page5.gridcontainer3.grid1
						.column2.columnorder=3
						.column8.columnorder=3
						.column2.columnorder=4
						.column4.columnorder=5
						.column5.columnorder=6
					endwith
				endif
		endcase

	case VLC_FormName == "K31"				&&-Nivel 0
		do case
			case upper(PLC_CodeId) = 'K31.K31T_K01_UKEY.TEXTALIAS1.FOL_UPDATE'
				do Tgr_K31_K01_Update
		endcase

	case VLC_FormName == "K40"				&&-Apontamento de materiais
		do case
			case PLC_CodeId == "K40.INIT"
				do Tgr_K40_Init with PLO_Form

			case PLC_CodeId == "K40.EXPQTT.FOL_UPDATE"
				do Tgr_K40_003_B_Update with PLO_Form
 
			case PLC_CodeId == "K40.FOL_BEFORESAVE"
				do tgr_k40_BeforeSave

			case PLC_CodeId == "K40.EXPQTT.SPINNER1.GOTFOCUS"
				*-  Utilizado no ponto de trigger FOL_Update do campo Quantidade para explosao
				PLO_Form.VON_ControleExplosao = PLO_Form.expqtt.spinner1.value
		endcase

	case VLC_FormName == "K49"				&&-operações da estrutura
		do case
			case PLC_CodeId == "K49.INIT"
				do K49_INIT with PLO_Form
		endcase

	case VLC_FormName == "K51"		&&-Apontamento de Recursos
		do case
			case PLC_CodeId == "K51.INIT"
				do Tgr_K51_Init with PLO_Form
			case PLC_CodeId == "K51.CGO_K19SELECTION1.K19_TEXT.TEXTALIAS1.FOL_UPDATE"
				do Tgr_K51_K19_Update
		endcase

	case VLC_FormName == "K60"				&&-operações do Roteiro
		do case
			case PLC_CodeId == "K60.INIT"
				do K60_INIT with PLO_Form
		endcase

	case VLC_FormName == "K61"				&&-Recursos das operações do Roteiro
		do case
			case PLC_CodeId == "K61.INIT"
				do K61_INIT with PLO_Form
		endcase

	case VLC_FormName == "M05"
		do case
			case upper(PLC_CodeId) = "M05.INIT"
				PLO_Form.VOL_accounting = .t.
		endcase

	case VLC_FORMNAME =="T01_002"			&& Relatório de Log de inconsistencias somente contabilidade
		do case
			case PLC_CodeId =="T01_002.ACTIVATE"
				select * from T01_T where ARRAY_434 = 2 into cursor T01_TA readwrite
			case PLC_CodeId =="T01_002.UNLOAD"
				if used('B06')
					select T01_TA
					if reccount('T01_TA') > 0
						VGO_GEN.VOC_REPORTNAME = "Log de Inconsistências - Integração Contábil"
						report form LOG_CONTABIL.FRX preview
					endif
				endif
				VGO_DAT.FOL_CLOSETABLE('T01_TA')
		endcase

	case VLC_FormName == "T04"				&&-Configurações de entrada e saída.
		do case
			case PLC_CodeId == "T04.INIT"
				* Inclui Campos para controle de limite de credito e inadimplencia
				do tgr_t04_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "T06"				&&-Itens Configurações de entrada e saída.
		do case
*!*				case PLC_CodeId == "T06.FOL_SAVE"
*!*					do tgr_t06_save.fxp
		
			case PLC_CodeId = "LOADTOOLBARAPPS"
				VGO_Custom.FOL_CAddToolBarApps ("Duplica Configuração da Natureza", [vgo_gen.fol_doform("t06_dup", t06t.t04_ukey, t06t.ukey, _screen.activeform )], !eof("T06T") and !_Screen.ActiveForm.vol_editing and !_screen.activeform.voo_refreshobj.Parent.Parent.VOL_editing)
		endcase

	case VLC_FormName == "T07"				&&-Configurações de entrada e saída.
		do case
			case PLC_CodeId == "T07.T07_A40_UKEY.TEXTALIAS1.LOSTFOCUS"
				if t04.t04_009_n = 2
					VGO_Gen.FOL_SqlExec("SELECT A40_001_C FROM STAR_DATA@A40 (NOLOCK) WHERE UKEY = ?T07T.A40_UKEY", "TmpA40")
					
					do case
						case upper(alltrim(TmpA40.A40_001_C)) == "ICMS"
							replace t07t.t07_006_n with 1 , ;
									t07t.t07_018_n with 0 , ;
									t07t.t07_008_n with 1 in t07t
						
						case upper(alltrim(TmpA40.A40_001_C)) == "IPI"
							replace t07t.t07_006_n with 1 , ;
									t07t.t07_018_n with 1 , ;
									t07t.t07_008_n with 0 in t07t
									
						otherwise
							replace t07t.t07_006_n with 0 , ;
									t07t.t07_018_n with 0 , ;
									t07t.t07_008_n with 0 in t07t
					endcase
		
					do tgr_t07_refresh.fxp with PLO_Form
*!*						PLO_Form.t07_t07_006_n.refresh()
*!*						PLO_Form.t07_t07_008_n.refresh()
*!*						PLO_Form.t07_t07_009_n.refresh()
*!*						PLO_Form.t07_t07_015_n.refresh()
*!*						PLO_Form.t07_t07_016_n.refresh()
*!*						PLO_Form.t07_t07_017_n.refresh()
*!*						PLO_Form.t07_t07_022_n.refresh()
		
					VGO_Gen.FOL_CloseTable("TmpA40")
				endif
		endcase


	case VLC_FormName == "T30"
		do case
			case PLC_CodeId == "T30.FOL_SAVE"
				VGO_Dat.FOL_PushSelect()
				select t32t
				count to VLN_Reccount                             &&- Faço assim pq qdo deleto do cursor o reccount ainda conta
				if VLN_Reccount = 0
					messagebox("É necessário escolher pelo menos um vendedor ou cliente  ", 48, "Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.
				endif
 
				scan
					if not t32t.t32_par $ "A03/A33"
						messagebox("A configuração é permitida somente para clientes ou vendedores  ", 48, "Atenção")
						VGO_Gen.VOL_TriggerReturn = .f.
						exit
					endif
				endscan
				VGO_Dat.FOL_PopSelect()


				*Case PLC_CodeId == "T30.INIT"
			case PLC_CodeId == upper("T30.pageframe1.page1.label_a03_ukey.Rightclick")
				do tgr_t30_init.fxp with PLO_Form 		&&- exibe o código do cliente no list de configuração do comissionamento (page1) quando der right clique sobre o label "cliente"
		endcase


	case VLC_FormName == "T31"				&& Comissao
		do case
			case PLC_CodeId == "T31.INIT"
				do tgr_t31_init.fxp with PLO_Form
		endcase

	case VLC_FormName == "Y14"				&&-Acessos específicos
		do case
			case '.FOL_LOADTREEDATA' $ PLC_CodeId
				do tgr_login_init.fxp with PLO_Form
		endcase


	case VLC_FormName == "USR" 				&&-Usuários do Sistema
		do case
			case PLC_CodeId $ "USR.INIT"
				do tgr_USR_init.fxp with PLO_Form
			
			case PLC_CodeId $ "USR.LOAD"
				PLO_Form.addproperty("Voc_LastUSRUkey627")
				PLO_Form.addproperty("Voc_MyExecuteCursor")
				PLO_Form.Voc_LastUSRUkey627 	= ""
				PLO_Form.Voc_MyExecuteCursor 	= ""

			case PLC_CodeId == "USR.FOL_EXECUTECURSOR"
				if !PLO_Form.Voc_MyExecuteCursor == USR.Ukey or vgo_gen.voa_TriggerParameters[1]
					PLO_Form.Voc_MyExecuteCursor = USR.Ukey
					with Vgo_Dat
						** Zonas
						.FOL_SqlExec("Select * from STAR_DATA@A25 (NOLOCK)","A25_ALL")
						** Zonas do Usuário
						.Voa_Index[1] = "A25_UKEY"
						.FOL_SetParameter(1,USR.Ukey)
						.FOL_EditCursor("JJ627_FRM","JJ627_FRM","JJ627T","Z",1,1)
					endwith

					PLO_Form.Voc_LastUSRUkey627		= ""
					if type("PLO_Form.Pageframe1.Page5.Moverlists1") = "O"
						=PLO_Form.Pageframe1.Page5.Moverlists1.FOL_CreateList()
					endif
				endif
				
			case PLC_CodeId $ "USR.FOL_SAVE"
				do tgr_USR_FOL_Save.fxp with PLO_Form
		endcase


	*- O pontos abaixo são relacionado ao modulo de custo
	*- para isto existe um trigger somente para o custo
	case VLC_FormName == 'V02'
		do triggers_custos with PLO_Form, PLO_Source, PLC_CodeId

	*-	Fechamento da contabilidade
	case 'B14_002.INIT' == upper(PLC_CodeId)
		do triggers_custos with PLO_Form, PLO_Source, PLC_CodeId

	case PLC_CodeId == "ACCINTEG.ACTIVATE" or PLC_CodeId == "RDACCINTEG.ACTIVATE"
		do triggers_custos with PLO_Form, PLO_Source, PLC_CodeId

	case PLC_CodeId == "COSTINTEG.ACTIVATE" or PLC_CodeId == "RDCOSTINTEG.ACTIVATE"
		do triggers_custos with PLO_Form, PLO_Source, PLC_CodeId
		


	
	otherwise 
		if upper(PLC_CodeId) $ "AFTERLAUNCHMENU"
			do tgr_definebars.fxp with PLO_Source, PLC_CodeId
		endif 
endcase






*-
*- Faço a verificação se na tela existem campos de pesquisa de local (d07) 
*- para bloqueio no FOL_BEFORESAVE de locais inativos
*- 
if ".INIT" $ PLC_CodeId and "CGO_TEXTALIASLABEL_D07" $ upper(PLO_Source.class)
	if not pemstatus(PLO_Form,"VOA_HasD07",5)
		PLO_Form.AddProperty("VOA_HasD07[5]")		&&- Verifico até 5 campos de locais no mesmo form
	endif 
	
	local VLN_x		&& pq nao aparece
	VLN_x = 10




	for VLN_x = 1 to alen(PLO_Form.VOA_HasD07,1)
		if vartype(PLO_Form.VOA_HasD07[VLN_x]) <> "O"
			PLO_Form.VOA_HasD07[VLN_x] = PLO_Source
		endif 
	endfor
endif 


*- se na tela existe pesquisa de local, antes de salvar verifico se o local esta ativo
if PLC_CodeId == VLC_FormName+".FOL_BEFORESAVE"	and pemstatus(PLO_Form,"VOA_HasD07",5) AND NOT inlist(nvl(VLC_FormName," "),"DD01", "DD03", "D12", "D17_001", "D17_002","D17_003", "D04")
	local VLC_D07ukey, VLN_x
	for VLN_x = 1 to alen(PLO_Form.VOA_HasD07,1)
		if vartype(PLO_Form.VOA_HasD07[VLN_x])="O"
			VLC_D07ukey = nvl(evaluate(PLO_Form.VOA_HasD07[VLN_x].VOC_AliasUkey),"")
			if not empty(VLC_D07ukey)
				VGO_Dat.FOL_PushSelect()
				VGO_Dat.FOL_SqlExec("select d07_001_c, d07_601_n from STAR_DATA@D07 D07 (nolock) where ukey = '"+VLC_D07ukey+"'","TMP_D07Disabled")
				if TMP_D07Disabled.d07_601_n = 1
					messagebox("O local de movimentação de estoque "+alltrim(TMP_D07Disabled.d07_001_c)+" está inativo.  "+chr(13)+;
							   "Não é permitida a movimentação nesse local.  ", 48, "Atenção")
					VGO_Gen.VOL_TriggerReturn = .f.		   
				endif 
				VGO_Dat.FOL_CloseTable("TMP_D07Disabled")
				VGO_Dat.FOL_PopSelect()
			endif 
			exit
		endif 
	endfor 
endif 


return


