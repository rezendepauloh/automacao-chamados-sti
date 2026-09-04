from .connection import DB_PATH, get_connection

from .ramais_db import (
    setup_ramais_table,
    save_ramais_to_db,
    get_ramais_df,
    setup_ramais_config_table,
    get_ramais_config,
    save_ramais_config
)

from .tickets_db import (
    setup_database,
    save_tickets_to_db,
    close_missing_tickets,
    close_missing_tickets_by_base,
    update_ticket_status,
    update_ticket_andamento,
    update_ticket_tag,
    update_ticket_location_details,
    update_ticket_title,
    save_comments_to_db,
    get_comments_by_ticket,
    sync_closed_tickets_to_train_dataset,
    load_data
)

from .map_db import (
    setup_map_tables,
    save_map_config,
    get_map_config,
    get_map_pins
)

from .donations_db import (
    setup_donations_table,
    sync_donations_from_excel,
    get_donations_data
)

from .settings_db import (
    setup_settings_table,
    get_setting,
    set_setting,
    get_all_settings,
    seed_settings_from_env_if_empty
)

from .plantoes_db import (
    setup_plantoes_tables,
    save_plantoes_matutino,
    save_plantoes_semanal,
    get_plantoes_matutino,
    get_plantoes_semanal
)

from .notifications_db import (
    setup_notifications_table,
    add_notification,
    get_notifications,
    get_unread_notifications_count,
    mark_notification_as_read,
    mark_notification_as_unread,
    mark_all_notifications_as_read
)

from .impressoras_db import (
    setup_impressoras_table,
    save_impressoras_to_db,
    get_impressoras_df
)

from .central_db import (
    get_central_telefonica_df
)

from .garantia_db import (
    setup_garantia_tables,
    sync_garantia_from_excel,
    get_garantia_contratos_df,
    get_garantia_chamados_df
)

from .events_db import (
    setup_eventos_manuais_table,
    save_evento_manual,
    get_eventos_manuais
)

from .unidades_db import (
    setup_unidades_tables,
    setup_unidades_manuais_table,
    get_unidades_manuais,
    save_unidades_to_db,
    get_unidades_df,
    add_unidade_manual,
    update_unidade_manual_by_id,
    delete_unidade_manual
)

from .fiscalizacao_db import (
    sync_fiscalizacao_from_excel,
    get_fiscalizacao_indicacoes_df,
    get_fiscalizacao_publicacoes_df,
    get_fiscalizacao_contador_df
)

from .viagens_db import (
    setup_viagens_table,
    sync_viagens_from_excel,
    get_viagens_df
)

from .whatsapp_db import (
    setup_whatsapp_tables,
    get_whatsapp_destinatarios,
    add_whatsapp_destinatario,
    toggle_whatsapp_destinatario_status,
    has_whatsapp_been_sent,
    log_whatsapp_dispatch,
    get_whatsapp_disparos_log
)

from .cron_db import (
    setup_cron_tables,
    get_cron_schedules,
    get_cron_schedule_by_id,
    update_cron_schedule,
    log_cron_execution_start,
    log_cron_execution_end,
    get_recent_cron_logs
)

__all__ = [
    "DB_PATH",
    "get_connection",
    "setup_ramais_table",
    "save_ramais_to_db",
    "get_ramais_df",
    "setup_database",
    "save_tickets_to_db",
    "close_missing_tickets",
    "close_missing_tickets_by_base",
    "update_ticket_status",
    "update_ticket_andamento",
    "update_ticket_tag",
    "update_ticket_location_details",
    "update_ticket_title",
    "save_comments_to_db",
    "get_comments_by_ticket",
    "sync_closed_tickets_to_train_dataset",
    "load_data",
    "setup_map_tables",
    "save_map_config",
    "get_map_config",
    "get_map_pins",
    "setup_donations_table",
    "sync_donations_from_excel",
    "get_donations_data",
    "setup_plantoes_tables",
    "save_plantoes_matutino",
    "save_plantoes_semanal",
    "get_plantoes_matutino",
    "get_plantoes_semanal",
    "setup_notifications_table",
    "add_notification",
    "get_notifications",
    "get_unread_notifications_count",
    "mark_notification_as_read",
    "mark_notification_as_unread",
    "mark_all_notifications_as_read",
    "setup_impressoras_table",
    "save_impressoras_to_db",
    "get_impressoras_df",
    "get_central_telefonica_df",
    "setup_garantia_tables",
    "sync_garantia_from_excel",
    "get_garantia_contratos_df",
    "get_garantia_chamados_df",
    "setup_eventos_manuais_table",
    "save_evento_manual",
    "get_eventos_manuais",
    "setup_unidades_tables",
    "setup_unidades_manuais_table",
    "get_unidades_manuais",
    "save_unidades_to_db",
    "get_unidades_df",
    "add_unidade_manual",
    "update_unidade_manual_by_id",
    "delete_unidade_manual",
    "sync_fiscalizacao_from_excel",
    "get_fiscalizacao_indicacoes_df",
    "get_fiscalizacao_publicacoes_df",
    "get_fiscalizacao_contador_df",
    "setup_viagens_table",
    "sync_viagens_from_excel",
    "get_viagens_df",
    "setup_whatsapp_tables",
    "get_whatsapp_destinatarios",
    "add_whatsapp_destinatario",
    "toggle_whatsapp_destinatario_status",
    "has_whatsapp_been_sent",
    "log_whatsapp_dispatch",
    "get_whatsapp_disparos_log",
    "setup_cron_tables",
    "get_cron_schedules",
    "get_cron_schedule_by_id",
    "update_cron_schedule",
    "log_cron_execution_start",
    "log_cron_execution_end",
    "get_recent_cron_logs"
]
