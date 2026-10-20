-- chamados
CREATE TABLE chamados (
        id TEXT PRIMARY KEY,
        data_criacao TEXT,
        titulo TEXT,
        cidade_predio TEXT,
        unidade TEXT,
        localidade_fisica TEXT,
        usuario TEXT,
        id_cliente TEXT,
        descricao TEXT,
        tag TEXT,
        ip_origem TEXT,
        status TEXT DEFAULT 'Aberto',
        data_atualizacao TEXT,
        base TEXT
    , andamento TEXT, link TEXT, hostname TEXT, tag_manual INTEGER DEFAULT 0, dados_manuais INTEGER DEFAULT 0);

-- comentarios
CREATE TABLE comentarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chamado_id TEXT,
        data TEXT,
        autor TEXT,
        texto TEXT,
        FOREIGN KEY(chamado_id) REFERENCES chamados(id) ON DELETE CASCADE
    );

-- mapa_config
CREATE TABLE mapa_config (
        id TEXT PRIMARY KEY,
        config_json TEXT
    );

-- mapa_pins
CREATE TABLE mapa_pins (
        id TEXT PRIMARY KEY,
        predio_id TEXT,
        pavimento_id INTEGER,
        sala TEXT,
        x INTEGER,
        y INTEGER,
        descricao TEXT
    );

-- equipamentos_doados
CREATE TABLE equipamentos_doados (
        patrimonio TEXT,
        modelo TEXT,
        serial_number TEXT,
        equipamento TEXT,
        tipo_movimentacao TEXT,
        data_movimentacao TEXT,
        chamado TEXT,
        ssd TEXT,
        motivo_baixa TEXT
    );

-- faqs
CREATE TABLE faqs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                tipo_faq TEXT NOT NULL,
                url TEXT NOT NULL UNIQUE,
                data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
            , conteudo TEXT);

-- plantoes_matutino
CREATE TABLE plantoes_matutino (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        data_iso TEXT,
        dia_semana TEXT,
        servidor TEXT,
        telefone TEXT,
        UNIQUE(ano, data_iso, servidor) ON CONFLICT REPLACE
    );

-- plantoes_semanal
CREATE TABLE plantoes_semanal (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        mes TEXT,
        periodo_str TEXT,
        data_inicio TEXT,
        data_fim TEXT,
        service_desk TEXT,
        manutencao TEXT,
        infraestrutura TEXT,
        desenvolvimento TEXT,
        UNIQUE(ano, data_inicio, manutencao) ON CONFLICT REPLACE
    );

-- notificacoes
CREATE TABLE notificacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL,
        titulo TEXT NOT NULL,
        mensagem TEXT NOT NULL,
        data_evento TEXT,
        link_pagina TEXT,
        lida INTEGER DEFAULT 0,
        data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(tipo, titulo, data_evento) ON CONFLICT IGNORE
    );

-- impressoras
CREATE TABLE impressoras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE,
        servidor TEXT,
        tipo TEXT,
        modelo TEXT,
        localizacao TEXT,
        ip_host TEXT,
        status TEXT,
        total_paginas INTEGER DEFAULT 0,
        filas_relacionadas TEXT,
        detalhes_extra TEXT,
        data_atualizacao TEXT
    );

-- garantia
CREATE TABLE garantia (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        patrimonio TEXT,
        serial_number TEXT,
        equipamento TEXT,
        modelo TEXT,
        fornecedor TEXT,
        contrato TEXT,
        n_saj TEXT,
        data_inicio TEXT,
        data_fim TEXT,
        status_garantia TEXT,
        dias_restantes INTEGER,
        localizacao TEXT,
        observacao TEXT,
        data_atualizacao TEXT
    );

-- garantia_contratos
CREATE TABLE garantia_contratos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contrato TEXT,
        pu_saj TEXT,
        item TEXT,
        contratacao_por TEXT,
        fornecedor TEXT,
        termo_referencia TEXT,
        termo_recebimento TEXT,
        nota_fiscal TEXT,
        garantia_inicio TEXT,
        garantia_fim TEXT,
        status_garantia TEXT,
        dias_restantes INTEGER,
        link_suporte TEXT,
        data_atualizacao TEXT
    );

-- garantia_chamados
CREATE TABLE garantia_chamados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item TEXT,
        status TEXT,
        numero_serie TEXT,
        patrimonio TEXT,
        chamado_mpm TEXT,
        chamado_externo TEXT,
        defeito TEXT,
        solucao TEXT,
        nota_no_chamado TEXT,
        chamado_dmp TEXT,
        data_atualizacao TEXT
    , data_abertura TEXT);

-- eventos_manuais
CREATE TABLE eventos_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        data_inicio TEXT NOT NULL,
        data_fim TEXT,
        descricao TEXT,
        autor TEXT,
        data_criacao TEXT
    );

-- unidades_manuais
CREATE TABLE unidades_manuais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        data_atualizacao TEXT
    );

-- unidades
CREATE TABLE unidades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cidade TEXT,
        tipo TEXT,
        setor TEXT,
        sigla TEXT,
        titular TEXT,
        unidade_predio TEXT,
        telefone TEXT,
        url TEXT,
        origem TEXT DEFAULT 'web',
        data_atualizacao TEXT
    );

-- ramais_mpms
CREATE TABLE ramais_mpms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        localidade TEXT,
        setor_nome TEXT,
        telefone_ramal TEXT,
        tipo TEXT,
        data_atualizacao TEXT
    );

-- ramais_config
CREATE TABLE ramais_config (
        chave TEXT PRIMARY KEY,
        valor TEXT
    );

-- faq_videos
CREATE TABLE faq_videos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        categoria TEXT NOT NULL,
        nome_arquivo TEXT NOT NULL,
        url TEXT NOT NULL UNIQUE,
        caminho_relativo TEXT,
        tamanho_bytes INTEGER DEFAULT 0,
        data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
    );

-- faq_imagens
CREATE TABLE faq_imagens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        categoria TEXT NOT NULL,
        nome_arquivo TEXT NOT NULL,
        url TEXT NOT NULL UNIQUE,
        caminho_relativo TEXT,
        tamanho_bytes INTEGER DEFAULT 0,
        data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
    );

-- system_settings
CREATE TABLE system_settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            is_secret INTEGER DEFAULT 0,
            category TEXT,
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

-- viagens
CREATE TABLE viagens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quem_foi TEXT,
        chamado TEXT,
        saida_iso TEXT,
        retorno_iso TEXT,
        saida_br TEXT,
        retorno_br TEXT,
        localidade TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    , chamado_diaria TEXT);

-- whatsapp_destinatarios
CREATE TABLE whatsapp_destinatarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT UNIQUE NOT NULL,
            ativo INTEGER DEFAULT 1,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );

-- whatsapp_disparos_log
CREATE TABLE whatsapp_disparos_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo_evento TEXT NOT NULL,
            evento_id TEXT NOT NULL,
            data_evento TEXT,
            destinatario TEXT NOT NULL,
            mensagem TEXT,
            status TEXT DEFAULT 'enviado',
            response_payload TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(tipo_evento, evento_id, destinatario) ON CONFLICT IGNORE
        );

-- cron_schedules
CREATE TABLE cron_schedules (
            task_id TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            categoria TEXT,
            ativo INTEGER DEFAULT 1,
            tipo_agendamento TEXT DEFAULT 'intervalo',
            intervalo_valor INTEGER DEFAULT 2,
            intervalo_unidade TEXT DEFAULT 'horas',
            horario_fixo TEXT DEFAULT '12:00',
            apenas_dias_uteis INTEGER DEFAULT 0,
            descricao TEXT,
            ultima_execucao DATETIME,
            proxima_execucao DATETIME,
            ultimo_status TEXT DEFAULT 'pendente',
            ultimo_log TEXT,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );

-- cron_logs
CREATE TABLE cron_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id TEXT NOT NULL,
            inicio DATETIME NOT NULL,
            fim DATETIME,
            duracao_segundos REAL,
            status TEXT NOT NULL,
            mensagem TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );

-- ad_cache_ous
CREATE TABLE ad_cache_ous (
        dn TEXT PRIMARY KEY,
        name TEXT,
        parent_dn TEXT,
        description TEXT,
        canonical_name TEXT,
        updated_at TEXT
    );

-- ad_cache_users
CREATE TABLE ad_cache_users (
        sam_account_name TEXT PRIMARY KEY,
        display_name TEXT,
        mail TEXT,
        department TEXT,
        title TEXT,
        dn TEXT,
        parent_ou_dn TEXT,
        is_active BOOLEAN,
        user_account_control INTEGER,
        when_created TEXT,
        last_logon TEXT,
        updated_at TEXT
    , pager TEXT, telephone_number TEXT, mobile TEXT, office TEXT, company TEXT, manager TEXT, description TEXT);

-- ad_cache_groups
CREATE TABLE ad_cache_groups (
        sam_account_name TEXT PRIMARY KEY,
        display_name TEXT,
        dn TEXT,
        description TEXT,
        group_type TEXT,
        member_count INTEGER DEFAULT 0,
        updated_at TEXT
    );

-- ad_cache_group_members
CREATE TABLE ad_cache_group_members (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_dn TEXT,
        member_dn TEXT,
        updated_at TEXT
    );

-- ad_sync_meta
CREATE TABLE ad_sync_meta (
        id INTEGER PRIMARY KEY,
        last_sync_at TEXT,
        status TEXT,
        total_ous INTEGER DEFAULT 0,
        total_users INTEGER DEFAULT 0,
        total_groups INTEGER DEFAULT 0,
        error_message TEXT
    , total_computers INTEGER DEFAULT 0);

-- ad_cache_computers
CREATE TABLE ad_cache_computers (
        name TEXT PRIMARY KEY,
        dns_hostname TEXT,
        operating_system TEXT,
        os_version TEXT,
        description TEXT,
        managed_by TEXT,
        dn TEXT,
        parent_ou_dn TEXT,
        is_active BOOLEAN,
        user_account_control INTEGER,
        when_created TEXT,
        last_logon TEXT,
        updated_at TEXT
    );

-- sccm_cache_devices
CREATE TABLE sccm_cache_devices (
        resource_id TEXT PRIMARY KEY,
        name TEXT,
        last_logon_user TEXT,
        ip_addresses TEXT,
        mac_addresses TEXT,
        manufacturer TEXT,
        model TEXT,
        operating_system TEXT,
        os_version TEXT,
        client_version TEXT,
        client_active INTEGER DEFAULT 1,
        last_active_time TEXT,
        ad_site_name TEXT,
        distinguished_name TEXT,
        raw_json TEXT,
        updated_at TEXT
    , processor TEXT, memory_ram TEXT, disk_drives TEXT);

-- sccm_cache_users
CREATE TABLE sccm_cache_users (
        resource_id TEXT PRIMARY KEY,
        user_name TEXT,
        full_user_name TEXT,
        user_group_name TEXT,
        windows_nt_domain TEXT,
        distinguished_name TEXT,
        raw_json TEXT,
        updated_at TEXT
    );

-- sccm_cache_collections
CREATE TABLE sccm_cache_collections (
        collection_id TEXT PRIMARY KEY,
        name TEXT,
        collection_type TEXT,
        member_count INTEGER DEFAULT 0,
        comment TEXT,
        last_refresh_time TEXT,
        updated_at TEXT
    );

-- ferias_bancada
CREATE TABLE ferias_bancada (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER,
        membro TEXT,
        tipo_escala TEXT, -- 'ferias', 'licencas_compensacoes', 'recesso_forense'
        mes INTEGER, -- 1 a 12 ou NULL para recesso
        mes_nome TEXT,
        periodo_bruto TEXT,
        data_inicio_iso TEXT,
        data_fim_iso TEXT,
        data_inicio_br TEXT,
        data_fim_br TEXT,
        dias INTEGER,
        status TEXT, -- 'Confirmada', 'Prevista', 'Usufruto', 'Venda', 'Cancelada', 'Recesso'
        cor_hex TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

-- garantia_agendamentos
CREATE TABLE garantia_agendamentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tecnico TEXT,
        data_prevista TEXT,
        hora_prevista TEXT,
        chamado_a_atender TEXT,
        data_atualizacao TEXT
    );

-- senhas_cofre
CREATE TABLE senhas_cofre (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        categoria TEXT NOT NULL,
        url_sistema TEXT DEFAULT '',
        usuario TEXT NOT NULL,
        senha_cifrada TEXT NOT NULL,
        observacoes TEXT DEFAULT '',
        data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

-- central_telefonica
CREATE TABLE "central_telefonica" (
"acd_group_directory_number" TEXT,
  "acd_type" TEXT,
  "add_on_module_1" TEXT,
  "add_on_module_2" TEXT,
  "add_on_module_3" TEXT,
  "advancedconfiguration" TEXT,
  "analogtype" TEXT,
  "annu_first_name" TEXT,
  "annu_name" TEXT,
  "associated_dect" TEXT,
  "associated_rsi" TEXT,
  "associated_station_directory_number" TEXT,
  "associated_videophone" TEXT,
  "attached_multitandem" TEXT,
  "cpctimer" TEXT,
  "ctq_ra_type" TEXT,
  "call_announce_group" TEXT,
  "call_by_name_mini_messaging_right" TEXT,
  "call_conference_group" TEXT,
  "call_priority" TEXT,
  "caller_category" TEXT,
  "clickandph" TEXT,
  "cmf_2600" TEXT,
  "connection_category_id" TEXT,
  "cost_center_id" TEXT,
  "cost_center_name" TEXT,
  "dev_routing_flag" TEXT,
  "dm_profile" TEXT,
  "data_connection_category_id" TEXT,
  "destination_node" TEXT,
  "dialled_number_masked_in_ticket" TEXT,
  "directory_number" TEXT,
  "display_compatibility_type_of_clip_set" TEXT,
  "display_name" TEXT,
  "domain_identifier" TEXT,
  "ext_gw" TEXT,
  "emulation_paging_right" TEXT,
  "entity_number" TEXT,
  "equipment_address_board" TEXT,
  "equipment_address_rack" TEXT,
  "equipment_address_terminal" TEXT,
  "explicit_prio" TEXT,
  "external_forwarding_category_id" TEXT,
  "external_keyboard_presence" TEXT,
  "external_login" TEXT,
  "facility_category_id" TEXT,
  "function" TEXT,
  "gfu_incoming_access" TEXT,
  "gfu_list_number" TEXT,
  "gfu_outgoing_access" TEXT,
  "gfu_preferential" TEXT,
  "gw_type" TEXT,
  "ghost_z" TEXT,
  "ghost_z_feature" TEXT,
  "implicit_prio" TEXT,
  "internal_keyboard" TEXT,
  "ith_act_mode" TEXT,
  "ith_service_type" TEXT,
  "ith_user_list" TEXT,
  "key_profile" TEXT,
  "language_id" TEXT,
  "mac_presence" TEXT,
  "mail_address" TEXT,
  "message_led" TEXT,
  "metering_category" TEXT,
  "modem_trunk_group_info" TEXT,
  "msg_class_nb" TEXT,
  "multi_line_facility" TEXT,
  "multi_line_station" TEXT,
  "nativeencryption" TEXT,
  "nb_error_codaffaire" TEXT,
  "nb_error_codsecret" TEXT,
  "nomadic_dec" TEXT,
  "number_of_users" TEXT,
  "occupation_type" TEXT,
  "ope_cca" TEXT,
  "opex_license" TEXT,
  "pcbt_associated" TEXT,
  "pin" TEXT,
  "paging_beeper_directory_number" TEXT,
  "paging_trunk_group_number" TEXT,
  "phone_applications_cos" TEXT,
  "phone_book_first_name" TEXT,
  "phone_book_name" TEXT,
  "pickup_group_name" TEXT,
  "pin_group_control" TEXT,
  "primary_line_prio" TEXT,
  "prio_presentation" TEXT,
  "private_calling_number" TEXT,
  "profile_type" TEXT,
  "public_network_category_id" TEXT,
  "recorder_directory_number" TEXT,
  "remote_ua" TEXT,
  "reserved_time_slot" TEXT,
  "right_profile_name" TEXT,
  "robinet_category" TEXT,
  "robinet_right" TEXT,
  "room_type" TEXT,
  "routing_number" TEXT,
  "s0_facility" TEXT,
  "s0_presence" TEXT,
  "sipgroup1srtpworkingmode" TEXT,
  "sip_authentification_login" TEXT,
  "sip_passwd" TEXT,
  "secondary_line_prio" TEXT,
  "secret_code" TEXT,
  "secret_code_access_for_suu" TEXT,
  "secretary_directory_number" TEXT,
  "set_role" TEXT,
  "stap" TEXT,
  "station_group_directory_number" TEXT,
  "station_sub_type" TEXT,
  "station_type" TEXT,
  "supplementary_service_right" TEXT,
  "supportutf8" TEXT,
  "tapipremiumserver" TEXT,
  "tandem" TEXT,
  "tele_marketing_actor" TEXT,
  "teleservice" TEXT,
  "teleservice_incident" TEXT,
  "terminalclip" TEXT,
  "text_msg_number" TEXT,
  "transparency_vsi" TEXT,
  "type_of_keyboard" TEXT,
  "ua_3g_features" TEXT,
  "ua_mr2_3g_features" TEXT,
  "url_domain" TEXT,
  "url_username" TEXT,
  "utf8_comment1" TEXT,
  "utf8_comment2" TEXT,
  "utf8_comment3" TEXT,
  "utf8_comment4" TEXT,
  "utf8_phone_book_first_name" TEXT,
  "utf8_phone_book_name" TEXT,
  "uniqueid" TEXT,
  "urgent_call" TEXT,
  "use_private_calling_number" TEXT,
  "user_pin_group" TEXT,
  "v24_presence" TEXT,
  "vip" TEXT,
  "vad_use" TEXT,
  "valid_for_call_by_name" TEXT,
  "voice_mail_box_a4610iv" TEXT,
  "voice_mail_box_emb_vm" TEXT,
  "voice_mail_box_octel" TEXT,
  "voice_mail_box_sympac" TEXT,
  "voice_mail_directory_number" TEXT,
  "voice_mail_secret_code" TEXT,
  "voice_mail_type" TEXT,
  "z_alarm_physical_configuration" TEXT,
  "z_subdevice" TEXT,
  "z_ivr" TEXT,
  "videosupportprofile" TEXT,
  "x25dte" TEXT,
  "ramal" TEXT,
  "nome_titular" TEXT,
  "complemento" TEXT,
  "tipo_de_estacao" TEXT,
  "subtipo" TEXT,
  "funcao_role" TEXT,
  "grupo_de_captura" TEXT,
  "cat_rede_pública" TEXT,
  "login_externo" TEXT,
  "e-mail" TEXT,
  "rack" TEXT,
  "placa" TEXT,
  "terminal" TEXT,
  "endereco_ip" TEXT,
  "mac_address" TEXT,
  "nome_exibido" TEXT,
  "endereco_equipamento" TEXT,
  "categoria_dispositivo" TEXT,
  "status_equipamento" TEXT
);

-- fiscalizacao_indicacoes
CREATE TABLE "fiscalizacao_indicacoes" (
"nº Saj" TEXT,
  "Fiscal titular" TEXT,
  "Fiscal suplente" TEXT,
  "Objeto" TEXT,
  "Contrato" TEXT
);

-- fiscalizacao_publicacoes
CREATE TABLE "fiscalizacao_publicacoes" (
"nº Saj" TEXT,
  "Fiscal titular" TEXT,
  "Fiscal suplente" TEXT,
  "Objeto" TEXT,
  "Portaria" TEXT,
  "Data Portaria" TIMESTAMP
);

-- fiscalizacao_contador
CREATE TABLE "fiscalizacao_contador" (
"Quantidade Fiscal" TEXT,
  "Qtd Titular" INTEGER,
  "Qtd Suplente" INTEGER
);

